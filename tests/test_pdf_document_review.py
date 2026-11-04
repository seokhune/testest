"""PDF 검수 워크플로우 단위 테스트."""

import asyncio
import base64
import io
from datetime import timedelta

import mistralai.workflows as sdk_workflows
import pytest
from mistralai.workflows.testing import create_test_worker

import worker.workflows.pdf_document_review as review_module
from worker.workflows.hwp_review_models import (
    DriveFileInfo,
    ExtractedDocument,
    HanjaTranslation,
    PdfReviewOutput,
    TextEdit,
)
from worker.workflows.pdf_document_review import (
    PdfDocumentReviewWorkflow,
    extract_page_blocks,
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


async def test_pdf_document_review_end_to_end(temporal_env, monkeypatch) -> None:
    pdf_text = make_pdf_text()
    uploaded: list[str] = []

    @sdk_workflows.activity(name="pdf-find-drive-file")
    async def fake_find_pdf_file(file_name: str) -> DriveFileInfo:
        return DriveFileInfo(
            file_id="pdf-id", title="조선사.pdf", parent_id="folder-1"
        )

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

    async def fake_parse(model_class, request):
        prompt = request.messages[-1].content
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
    monkeypatch.setattr(review_module, "upload_drive_file", fake_upload_drive_file)
    monkeypatch.setattr(review_module, "chat_parse_to_model", fake_parse)

    async with create_test_worker(
        temporal_env,
        workflows=[PdfDocumentReviewWorkflow],
        activities=[
            fake_find_pdf_file,
            fake_read_pdf_text,
            fake_upload_drive_file,
            review_module.proofread_body,
            review_module.translate_hanja_blocks,
            review_module.build_output_docx,
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
    assert uploaded == [
        "조선사_p8-15_교정본.docx",
        "조선사_p8-15_검수리포트.docx",
    ]
