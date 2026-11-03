"""HWP 검수 워크플로우와 DOCX 헬퍼 테스트."""

import asyncio
import base64
import io
import zipfile
from datetime import timedelta

import mistralai.workflows as workflows
import pytest
from mistralai.workflows.testing import create_test_worker

import worker.workflows.hwp_document_review as review_module
from worker.workflows.hwp_document_review import HwpDocumentReviewWorkflow
from worker.workflows.hwp_review_docx import (
    apply_review_edits,
    convert_hancom_to_docx,
    extract_document,
)
from worker.workflows.hwp_review_models import (
    BodyFootnoteDifference,
    DriveFileInfo,
    ExtractedDocument,
    HanjaTranslation,
    HwpReviewOutput,
    SourceAnnotation,
    TextEdit,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOCUMENT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""


def _paragraph(text: str) -> str:
    return f'<w:p xmlns:w="{W_NS}"><w:r><w:t>{text}</w:t></w:r></w:p>'


DOCUMENT_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}"><w:body>{_paragraph("첫 번째 문단입니다.")}{_paragraph("두 번째 문단, 오탈자 문장이다.")}<w:sectPr/></w:body></w:document>"""

FOOTNOTES_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W_NS}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>주석 1: 『번역서』의 주석이다.</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>주석 2: 子曰、學而時習之、不亦說乎。</w:t></w:r></w:p></w:footnote>
</w:footnotes>"""


def make_docx() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        archive.writestr("_rels/.rels", ROOT_RELS_XML)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)
        archive.writestr("word/document.xml", DOCUMENT_XML)
        archive.writestr("word/footnotes.xml", FOOTNOTES_XML)
    return buffer.getvalue()


def test_extract_document_reads_body_and_footnotes() -> None:
    document = extract_document(make_docx(), source_format="hwp")

    assert [block.text for block in document.body] == [
        "첫 번째 문단입니다.",
        "두 번째 문단, 오탈자 문장이다.",
    ]
    assert [footnote.footnote_id for footnote in document.footnotes] == ["1", "2"]
    assert not review_module.contains_hanja(document.footnotes[0].text)
    assert review_module.contains_hanja(document.footnotes[1].text)


def test_apply_review_edits_updates_body_and_footnotes() -> None:
    edited = apply_review_edits(
        make_docx(),
        [
            TextEdit(
                item_id="p1",
                original="두 번째 문단, 오탈자 문장이다.",
                revised="두 번째 문단, 오탈자를 고친 문장이다.",
                category="typo",
                reason="오탈자 수정",
            )
        ],
        [
            TextEdit(
                item_id="1",
                original="주석 1: 『번역서』의 주석이다.",
                revised="주석 1: 〈번역서〉의 주석이다.",
                category="book_title",
                reason="서지명 표기 통일",
            )
        ],
        [
            SourceAnnotation(
                item_id="1", source="홍길동, 〈번역서〉, 출판사, 2020 (추정)"
            )
        ],
        [
            HanjaTranslation(
                item_id="2",
                original="子曰、學而時習之。",
                translation="공자가 말하였다. 배우고 때로 익히면 즐겁지 않은가.",
            )
        ],
    )

    document = extract_document(edited, source_format="hwp")
    assert document.body[1].text == "두 번째 문단, 오탈자를 고친 문장이다."
    assert document.footnotes[0].text == (
        "주석 1: 〈번역서〉의 주석이다. [출처: 홍길동, 〈번역서〉, 출판사, 2020 (추정)]"
    )
    assert "번역: 공자가 말하였다." in document.footnotes[1].text


def test_apply_review_edits_ignores_unknown_ids() -> None:
    edited = apply_review_edits(
        make_docx(),
        [
            TextEdit(
                item_id="p999", original="x", revised="y", category="typo", reason="r"
            )
        ],
        [
            TextEdit(
                item_id="999", original="x", revised="y", category="typo", reason="r"
            )
        ],
        [],
        [],
    )
    document = extract_document(edited, source_format="hwp")
    assert len(document.body) == 2
    assert len(document.footnotes) == 2


def test_convert_hancom_to_docx_rejects_unknown_format() -> None:
    with pytest.raises(ValueError):
        convert_hancom_to_docx(b"not-a-document", "docx")


async def test_hwp_document_review_end_to_end(temporal_env, monkeypatch) -> None:
    docx_bytes = make_docx()
    content_b64 = base64.b64encode(docx_bytes).decode()
    copied_titles: list[str] = []
    uploaded: list[str] = []

    @workflows.activity(name="hwp-find-drive-file")
    async def fake_find_drive_file(file_name: str) -> DriveFileInfo:
        return DriveFileInfo(
            file_id="source-id", title="번역서.hwp", parent_id="folder-1"
        )

    @workflows.activity(name="hwp-copy-drive-file")
    async def fake_copy_drive_file(
        file_id: str, title: str, parent_id: str | None
    ) -> DriveFileInfo:
        copied_titles.append(title)
        return DriveFileInfo(file_id="copy-id", title=title, parent_id=parent_id)

    @workflows.activity(name="hwp-download-drive-file")
    async def fake_download_drive_file(file_id: str) -> str:
        return content_b64

    @workflows.activity(name="hwp-upload-drive-file")
    async def fake_upload_drive_file(
        title: str, content: str, parent_id: str | None
    ) -> DriveFileInfo:
        uploaded.append(title)
        return DriveFileInfo(
            file_id=f"upload-{len(uploaded)}", title=title, parent_id=parent_id
        )

    @workflows.activity(name="hwp-convert-and-extract")
    async def fake_convert_and_extract(
        content: str, file_title: str
    ) -> ExtractedDocument:
        return extract_document(docx_bytes, source_format="hwp")

    async def fake_parse(model_class, request):
        prompt = request.messages[-1].content
        if model_class is review_module._EditBatch:
            if "본문 문단 목록" in prompt:
                return review_module._EditBatch(
                    edits=[
                        TextEdit(
                            item_id="p1",
                            original="두 번째 문단, 오탈자 문장이다.",
                            revised="두 번째 문단, 오탈자를 고친 문장이다.",
                            category="typo",
                            reason="오탈자 수정",
                        )
                    ]
                )
            return review_module._EditBatch(
                edits=[
                    TextEdit(
                        item_id="1",
                        original="주석 1: 『번역서』의 주석이다.",
                        revised="주석 1: 〈번역서〉의 주석이다.",
                        category="book_title",
                        reason="서지명 표기 통일",
                    )
                ]
            )
        if model_class is review_module._SourceBatch:
            return review_module._SourceBatch(
                annotations=[
                    SourceAnnotation(item_id="1", source="홍길동, 〈번역서〉 (추정)")
                ]
            )
        if model_class is review_module._HanjaBatch:
            return review_module._HanjaBatch(
                translations=[
                    HanjaTranslation(
                        item_id="2",
                        original="子曰、學而時習之。",
                        translation="공자가 말하였다. 배우고 때로 익히면 즐겁다.",
                    )
                ]
            )
        return review_module._DifferenceBatch(
            differences=[
                BodyFootnoteDifference(
                    body_block_id="p0",
                    footnote_id="1",
                    difference="본문 인용과 주석 근거문이 일치하지 않는다.",
                )
            ]
        )

    monkeypatch.setattr(review_module, "find_drive_file", fake_find_drive_file)
    monkeypatch.setattr(review_module, "copy_drive_file", fake_copy_drive_file)
    monkeypatch.setattr(review_module, "download_drive_file", fake_download_drive_file)
    monkeypatch.setattr(review_module, "upload_drive_file", fake_upload_drive_file)
    monkeypatch.setattr(review_module, "convert_and_extract", fake_convert_and_extract)
    monkeypatch.setattr(review_module, "chat_parse_to_model", fake_parse)

    activities = [
        fake_find_drive_file,
        fake_copy_drive_file,
        fake_download_drive_file,
        fake_upload_drive_file,
        fake_convert_and_extract,
        review_module.proofread_body,
        review_module.polish_footnotes,
        review_module.annotate_footnote_sources,
        review_module.translate_hanja_footnotes,
        review_module.compare_body_footnotes,
        review_module.build_corrected_docx,
        review_module.build_review_report,
    ]

    async with create_test_worker(
        temporal_env, workflows=[HwpDocumentReviewWorkflow], activities=activities
    ):
        handle = await temporal_env.client.start_workflow(
            "hwp-document-review",
            {"file_name": "번역서"},
            id="test-hwp-document-review",
            task_queue="test-task-queue",
            execution_timeout=timedelta(seconds=30),
        )
        result = await asyncio.wait_for(handle.result(), timeout=45)

    output = HwpReviewOutput.model_validate(result)
    assert output.source_title == "번역서.hwp"
    assert output.working_copy_title == "번역서_검수사본"
    assert copied_titles == ["번역서_검수사본"]
    assert uploaded == ["번역서_교정본.docx", "번역서_교정리포트.docx"]
    assert output.stats["body_edits"] == 1
    assert output.stats["footnote_edits"] == 1
    assert output.stats["source_annotations"] == 1
    assert output.stats["hanja_translations"] == 1
    assert output.stats["body_footnote_differences"] == 1
    assert (
        "교정리포트" in output.summary
        or output.report_file.title == "번역서_교정리포트.docx"
    )
