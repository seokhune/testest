"""PDF 검수 워크플로우 단위 테스트."""

import asyncio
import base64
from datetime import timedelta

import mistralai.workflows as sdk_workflows
import pytest
from mistralai.workflows.testing import create_test_worker

import worker.workflows.pdf_document_review as review_module
from worker.workflows.hwp_review_models import (
    DriveFileInfo,
    HanjaTranslation,
    PdfReviewOutput,
    TextEdit,
)
from worker.workflows.pdf_document_review import (
    PdfDocumentReviewWorkflow,
    detect_year,
    extract_page_blocks,
    group_by_year,
    split_pdf_pages,
)


def make_pdf_text() -> str:
    parts = []
    for page in range(6, 18):
        parts.append(f"- {page} -")
        parts.append(f"{page}번째 페이지 첫 문단입니다. 조선의 사적을 따른다.")
        parts.append("")
        parts.append(f"{page}번째 페이지 둘째 문단, 태조 원년의 기사이다.")
        parts.append("")
    return "\n".join(parts)


def test_split_pdf_pages_reads_markers() -> None:
    pages = split_pdf_pages(make_pdf_text())
    assert [number for number, _ in pages] == list(range(6, 18))
    assert "6번째 페이지 첫 문단입니다." in pages[0][1]


def test_extract_page_blocks_slices_range() -> None:
    blocks, total_pages = extract_page_blocks(
        make_pdf_text(), page_start=8, page_end=15
    )
    assert total_pages == 17
    assert {block.page for block in blocks} == set(range(8, 16))
    assert len([b for b in blocks if b.page == 8]) == 2
    assert blocks[0].block_id == "p0"
    assert blocks[0].text == "8번째 페이지 첫 문단입니다. 조선의 사적을 따른다."


def test_extract_page_blocks_without_range_keeps_all() -> None:
    blocks, _ = extract_page_blocks(make_pdf_text())
    assert len(blocks) == 24
    assert blocks[0].page == 6


def test_extract_page_blocks_unknown_range_raises() -> None:
    with pytest.raises(Exception):
        extract_page_blocks(make_pdf_text(), page_start=50, page_end=60)


def test_extract_page_blocks_without_markers_single_page() -> None:
    text = "표식 없는 문서\n\n두 번째 문단"
    blocks, total_pages = extract_page_blocks(text)
    assert total_pages == 1
    assert [b.page for b in blocks] == [1, 1]


def test_detect_year_reads_reign_year_headings() -> None:
    assert detect_year("태종 6년(무인) 봄 정월에 ...") == "태종 6년(무인)"
    assert detect_year("2\n6년(임오) 상왕이 ...") == "6년(임오)"
    assert detect_year("태조 원년 임신에 ...") == "태조 원년"
    assert detect_year("그해 겨울에 큰 눈이 내렸다.") is None


def test_group_by_year_buckets_edits() -> None:
    from worker.workflows.hwp_review_models import BodyBlock

    blocks = [
        BodyBlock(block_id="p0", text="태종 6년(무인) 봄 정월에 초하루 병술에", page=8),
        BodyBlock(block_id="p1", text="김첨이를 의정부에 보내았다.", page=8),
        BodyBlock(block_id="p2", text="태종 7년(기묘) 가을에", page=9),
        BodyBlock(block_id="p3", text="하륜이 상서하였다.", page=9),
    ]
    edits = [
        TextEdit(item_id="p1", original="a", revised="b", category="typo", reason="r"),
        TextEdit(item_id="p3", original="c", revised="d", category="spacing", reason="s"),
    ]
    translations = [
        HanjaTranslation(item_id="p0", original="漢文", translation="한문 번역"),
    ]
    from worker.workflows.hwp_review_models import BodyFootnoteDifference

    differences = [
        BodyFootnoteDifference(
            body_block_id="p3", footnote_id="5", difference="인용 문구가 다르다"
        )
    ]
    sections = group_by_year(blocks, edits, translations, differences)
    assert [year for year, _, _, _ in sections] == ["태종 6년(무인)", "태종 7년(기묘)"]
    assert [e.item_id for e in sections[0][1]] == ["p1"]
    assert [t.item_id for t in sections[0][2]] == ["p0"]
    assert [e.item_id for e in sections[1][1]] == ["p3"]
    assert [d.footnote_id for d in sections[1][3]] == ["5"]


def test_reconcile_edits_remaps_invented_ids() -> None:
    from worker.workflows.hwp_review_models import BodyBlock
    from worker.workflows.pdf_document_review import reconcile_edits

    blocks = [
        BodyBlock(block_id="p0", text="태종 6년(무인) 봄 정월에", page=8),
        BodyBlock(block_id="p1", text="김첨을 의정부에 보내았다.", page=8),
    ]
    edits = [
        TextEdit(
            item_id="p12",
            original="김첨을 의정부에 보내았다.",
            revised="김첨을 의정부에 보내었다.",
            category="typo",
            reason="어미 교정",
        ),
        TextEdit(
            item_id="p99",
            original="매칭 안 되는 문장",
            revised="무시된다",
            category="typo",
            reason="r",
        ),
    ]
    result = reconcile_edits(blocks, edits)
    assert len(result) == 1
    assert result[0].item_id == "p1"
    assert result[0].revised == "김첨을 의정부에 보내었다."


def _docx_bytes_from_paragraphs(paragraphs: list[str]) -> bytes:
    import io

    from docx import Document

    document = Document()
    for paragraph in paragraphs:
        document.add_paragraph(paragraph)
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def test_apply_edits_by_text_preserves_formatting() -> None:
    import io

    from docx import Document
    from worker.workflows.hwp_review_docx import apply_edits_by_text

    document = Document()
    paragraph = document.add_paragraph()
    bold = paragraph.add_run("태조원년에 ")
    bold.bold = True
    paragraph.add_run("개국하였다. 子曰學而時習之。")
    buffer = io.BytesIO()
    document.save(buffer)

    edits = [
        TextEdit(
            item_id="p0",
            original="태조원년에 개국하였다. 子曰學而時習之。",
            revised="태조 원년에 개국하였다. 子曰學而時習之。",
            category="spacing",
            reason="띄어쓰기",
        )
    ]
    translations = [
        HanjaTranslation(
            item_id="p0",
            original="태조원년에 개국하였다. 子曰學而時習之。",
            translation="공자가 말하기를, 배우고 때로 익히면 즐겁다.",
        )
    ]
    result, applied = apply_edits_by_text(buffer.getvalue(), edits, translations)
    assert applied == 1

    document = Document(io.BytesIO(result))
    paragraph = document.paragraphs[0]
    assert paragraph.text == "태조 원년에 개국하였다. 子曰學而時習之。"
    bold_runs = [run for run in paragraph.runs if run.bold]
    assert bold_runs, "굵은 서식이 보존되어야 한다"
    assert "번역:" not in paragraph.text, "본문에 번역이 끼워지지 않아야 한다"


def test_build_structured_docx_preserves_pages_and_applies_edits() -> None:
    from worker.workflows.hwp_review_models import BodyBlock
    from worker.workflows.hwp_review_docx import build_structured_docx

    pages = [
        (8, "첫 문단입니다.\n\n고려 태조 원년의 기사이다."),
        (9, "둘째 페이지 문단입니다."),
    ]
    blocks = [
        BodyBlock(block_id="p0", text="첫 문단입니다.", page=8),
        BodyBlock(block_id="p1", text="고려 태조 원년의 기사이다.", page=8),
        BodyBlock(block_id="p2", text="둘째 페이지 문단입니다.", page=9),
    ]
    edits = [
        TextEdit(
            item_id="p1",
            original="원문과 달라도 상관없다",
            revised="고려 태조 원년의 기사다.",
            category="grammar",
            reason="문맥상 종결형",
        )
    ]
    result, applied = build_structured_docx(pages, blocks, edits, [])
    assert applied == 1

    import io

    from docx import Document

    document = Document(io.BytesIO(result))
    texts = [p.text for p in document.paragraphs]
    assert "- 8 -" in texts and "- 9 -" in texts
    assert "고려 태조 원년의 기사다." in texts


def test_unwrap_text_frames_keeps_footnote_text() -> None:
    import io

    from docx import Document
    from docx.oxml.ns import qn
    from worker.workflows.hwp_review_docx import unwrap_text_frames

    document = Document()
    document.add_paragraph("본문 문단입니다.")
    footnote = document.add_paragraph("1) 주석 인용문입니다.")
    footnote_paragraph = footnote._p
    frame_pr = footnote_paragraph.get_or_add_pPr().makeelement(
        qn("w:framePr"), {}
    )
    footnote_paragraph.get_or_add_pPr().append(frame_pr)
    buffer = io.BytesIO()
    document.save(buffer)

    result = unwrap_text_frames(buffer.getvalue())
    document = Document(io.BytesIO(result))
    texts = [p.text for p in document.paragraphs]
    assert "본문 문단입니다." in texts
    assert "1) 주석 인용문입니다." in texts
    assert all(
        p._p.find(qn("w:pPr")) is None
        or p._p.find(qn("w:pPr")).find(qn("w:framePr")) is None
        for p in document.paragraphs
    )


def test_apply_edits_skips_large_deletions() -> None:
    from worker.workflows.hwp_review_docx import apply_edits_by_text

    docx_bytes = _docx_bytes_from_paragraphs(
        ["본문 문단입니다. 1) 주석 인용문이 여기에 길게 이어집니다. 2) 또 다른 주석입니다."]
    )
    edits = [
        TextEdit(
            item_id="p0",
            original="본문 문단입니다. 1) 주석 인용문이 여기에 길게 이어집니다. 2) 또 다른 주석입니다.",
            revised="본문 문단입니다.",  # 주석을 통째로 지우는 교정
            category="style",
            reason="삭제 시도",
        )
    ]
    result, applied = apply_edits_by_text(docx_bytes, edits, [])
    assert applied == 0

    import io

    from docx import Document

    document = Document(io.BytesIO(result))
    assert "주석 인용문" in document.paragraphs[0].text


def test_append_review_report_groups_by_year() -> None:
    from worker.workflows.hwp_review_docx import append_review_report

    docx_bytes = _docx_bytes_from_paragraphs(["본문 문단입니다."])
    from worker.workflows.hwp_review_models import BodyFootnoteDifference

    sections = [
        (
            "태종 6년(무인)",
            [TextEdit(item_id="p1", original="원문", revised="수정", category="typo", reason="사유")],
            [HanjaTranslation(item_id="p1", original="漢文", translation="번역")],
            [
                BodyFootnoteDifference(
                    body_block_id="p1", footnote_id="2", difference="본문과 주석 인용이 다르다"
                )
            ],
        ),
    ]
    result = append_review_report(
        docx_bytes, "조선사.pdf", sections, {"body_paragraphs": 2, "body_edits": 1, "hanja_translations": 1}
    )

    import io

    from docx import Document

    document = Document(io.BytesIO(result))
    texts = [p.text for p in document.paragraphs]
    assert "본문 문단입니다." in texts
    assert "조선사.pdf 검수 리포트" in texts
    assert "태종 6년(무인)" in texts
    assert "수정: 수정" in texts
    assert "번역: 번역" in texts
    assert "[주석 2] 본문과 주석 인용이 다르다" in texts


async def test_pdf_document_review_end_to_end(temporal_env, monkeypatch) -> None:
    pdf_text = make_pdf_text()
    uploaded: list[str] = []

    @sdk_workflows.activity(name="pdf-find-drive-file")
    async def fake_find_pdf_file(file_name: str) -> DriveFileInfo:
        return DriveFileInfo(file_id="pdf-id", title="조선사.pdf", parent_id="folder-1")

    @sdk_workflows.activity(name="pdf-read-text")
    async def fake_read_pdf_text(file_id: str) -> str:
        return pdf_text

    @sdk_workflows.activity(name="pdf-upload-drive-file")
    async def fake_upload_drive_file(
        title: str, content: str, parent_id: str | None
    ) -> DriveFileInfo:
        uploaded.append(title)
        assert base64.b64decode(content)[:2] == b"PK"
        return DriveFileInfo(
            file_id=f"upload-{len(uploaded)}", title=title, parent_id=parent_id
        )

    @sdk_workflows.activity(name="pdf-build-result")
    async def fake_build_result(
        source: DriveFileInfo,
        pages: list[tuple[int, str]],
        blocks: list,
        total_pages: int,
        body_edits: list[TextEdit],
        hanja_translations: list[HanjaTranslation],
        year_sections: list,
        stats: dict,
    ) -> str:
        assert year_sections, "년도 구간이 계산되어야 한다"
        docx_bytes = _docx_bytes_from_paragraphs(["결과 문단입니다."])
        return base64.b64encode(docx_bytes).decode()

    @sdk_workflows.activity(name="pdf-annotate-source-titles")
    async def fake_annotate_source_titles(blocks: list) -> list:
        return []

    @sdk_workflows.activity(name="pdf-compare-quotes")
    async def fake_compare_quotes(blocks: list) -> list:
        return []

    async def fake_parse(model_class, request):
        prompt = request.messages[-1].content
        if model_class is review_module._DifferenceBatch:
            return review_module._DifferenceBatch(differences=[])
        if model_class is review_module._EditBatch:
            if fake_parse.edit_calls:
                return review_module._EditBatch(edits=[])
            fake_parse.edit_calls.append(1)
            return review_module._EditBatch(
                edits=[
                    TextEdit(
                        item_id="p1",
                        original="6번째 페이지 첫 문단입니다. 조선의 사적을 따른다.",
                        revised="8번째 페이지 첫 문단입니다. 조선의 사적을 따른다.",
                        category="typo",
                        reason="페이지 번호 오기 수정",
                    )
                ]
            )
        assert model_class is review_module._HanjaBatch
        return review_module._HanjaBatch(translations=[])

    fake_parse.edit_calls = []

    monkeypatch.setattr(review_module, "find_pdf_file", fake_find_pdf_file)
    monkeypatch.setattr(review_module, "read_pdf_text", fake_read_pdf_text)
    monkeypatch.setattr(review_module, "annotate_source_titles", fake_annotate_source_titles)
    monkeypatch.setattr(review_module, "compare_quotes", fake_compare_quotes)
    monkeypatch.setattr(review_module, "upload_drive_file", fake_upload_drive_file)
    monkeypatch.setattr(review_module, "chat_parse_to_model", fake_parse)
    monkeypatch.setattr(review_module, "build_result_docx", fake_build_result)

    async with create_test_worker(
        temporal_env,
        workflows=[PdfDocumentReviewWorkflow],
        activities=[
            fake_find_pdf_file,
            fake_read_pdf_text,
            fake_upload_drive_file,
            fake_build_result,
            fake_annotate_source_titles,
            fake_compare_quotes,
            review_module.proofread_body,
            review_module.translate_hanja_blocks,
        ],
    ) as _worker:
        handle = await temporal_env.client.start_workflow(
            "pdf-document-review",
            {"file_name": "조선사", "page_start": 8, "page_end": 15},
            id="test-pdf-document-review",
            task_queue="test-task-queue",
            execution_timeout=timedelta(seconds=30),
        )
        result = await asyncio.wait_for(handle.result(), timeout=45)

    output = PdfReviewOutput.model_validate(result)
    assert output.source_title == "조선사.pdf"
    assert output.stats["page_start"] == 8 and output.stats["page_end"] == 15
    assert output.stats["total_pages"] == 17
    assert output.stats["body_paragraphs"] == 16
    assert output.stats["body_edits"] == 1
    assert "8~15페이지" in output.summary
    assert "인용 차이" in output.summary
    assert uploaded == ["조선사_p8-15_검수결과.docx"]
