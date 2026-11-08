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
    blocks_from_pages,
    detect_year,
    group_by_year,
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


def test_blocks_from_pages_merges_lines_into_paragraphs() -> None:
    from worker.workflows.pdf_document_review import blocks_from_pages

    pages = [
        (
            8,
            "태조 19년(임신) 정월에 하왕이\n즉위하였다.\n하왕이 폐세자를 서인으로\n보내였다.\n1) <보기> 인용문이다.",
        ),
        (9, "- 9 -\n한 문장입니다."),
    ]
    blocks = blocks_from_pages(pages)
    assert [(b.page, b.text) for b in blocks] == [
        (8, "태조 19년(임신) 정월에 하왕이 즉위하였다."),
        (8, "하왕이 폐세자를 서인으로 보내였다."),
        (8, "1) <보기> 인용문이다."),
        (9, "한 문장입니다."),
    ]
    assert blocks[0].block_id == "p0"
    assert blocks[3].block_id == "p3"


def test_detect_year_reads_reign_year_headings() -> None:
    assert detect_year("태종 6년(무인) 봄 정월에 ...") == "태종 6년(무인)"
    assert detect_year("2\n6년(임오) 상왕이 ...") == "6년(임오)"
    assert detect_year("태조 원년 임신에 ...") == "태조 원년"
    assert detect_year("그해 겨울에 큰 눈이 내렸다.") is None


def test_group_by_year_buckets_edits() -> None:
    from worker.workflows.hwp_review_models import BodyBlock

    blocks = [
        BodyBlock(block_id="p0", text="태종 6년(무인) 봄 정월에 초하루 병술에", page=8),
        BodyBlock(block_id="p1", text="김첨을 의정부에 보내았다.", page=8),
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
    assert [section["year"] for section in sections] == [
        "태종 6년(무인)",
        "태종 7년(기묘)",
    ]
    assert sections[0]["heading"].startswith("태종 6년(무인)")
    assert [e.item_id for e in sections[0]["edits"]] == ["p1"]
    assert [t.item_id for t in sections[0]["translations"]] == ["p0"]
    assert [e.item_id for e in sections[1]["edits"]] == ["p3"]
    assert [d.footnote_id for d in sections[1]["differences"]] == ["5"]


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


def test_insert_year_reports_places_full_report_at_year_ends() -> None:
    import io

    from docx import Document
    from worker.workflows.hwp_review_models import BodyFootnoteDifference
    from worker.workflows.pdf_document_review import insert_year_reports

    document = Document()
    document.add_paragraph("태조 19년(임신) 정월에 하왕이 즉위하였다.")
    document.add_paragraph("하왕이 폐세자를 서인으로 보내였다.")
    document.add_paragraph("태조 20년(계유) 봄 정월에 하왕이 죽으니 태종이 즉위하였다.")
    document.add_paragraph("태종이 하륜을 문하부에 보내였다.")
    buffer = io.BytesIO()
    document.save(buffer)

    sections = [
        {
            "year": "태조 19년(임신)",
            "heading": "태조 19년(임신) 정월에 하왕이 즉위하였다.",
            "edits": [
                TextEdit(
                    item_id="p1",
                    original="하왕이 폐세자를 서인으로 보내였다.",
                    revised="하왕이 폐세자를 서인으로 보내었다.",
                    category="typo",
                    reason="어미 교정",
                )
            ],
            "translations": [
                HanjaTranslation(
                    item_id="p0", original="漢文 주석", translation="한문 번역문"
                )
            ],
            "differences": [
                BodyFootnoteDifference(
                    body_block_id="p1", footnote_id="2", difference="인용 불일치"
                )
            ],
            "annotations": [],
        },
        {
            "year": "태조 20년(계유)",
            "heading": "태조 20년(계유) 봄 정월에 하왕이 죽으니 태종이 즉위하였다.",
            "edits": [
                TextEdit(
                    item_id="p3",
                    original="태종이 하륜을 문하부에 보내였다.",
                    revised="태종이 하륜을 문하부에 보내었다.",
                    category="typo",
                    reason="어미 교정",
                )
            ],
            "translations": [],
            "differences": [],
            "annotations": [],
        },
    ]
    result = insert_year_reports(buffer.getvalue(), sections)
    document = Document(io.BytesIO(result))
    texts = [p.text for p in document.paragraphs]

    assert "〔태조 19년(임신) 검수 리포트〕" in texts
    assert "〔태조 20년(계유) 검수 리포트〕" in texts
    # 19년 리포트는 20년 제목 직전에 위치해야 한다
    index_19 = texts.index("〔태조 19년(임신) 검수 리포트〕")
    index_20_heading = texts.index(
        "태조 20년(계유) 봄 정월에 하왕이 죽으니 태종이 즉위하였다."
    )
    assert index_19 < index_20_heading
    # 리포트에 네 항목이 모두 들어 있어야 한다
    body_19 = texts[index_19:index_20_heading]
    assert "1) 교정 내역" in body_19
    assert "2) 한문 번역 - 주석번역" in body_19
    assert "3) 본문-주석 내용 차이" in body_19
    assert "4) 주석 출처 보완 제안" in body_19
    # 교정 내역은 문단 전체가 아니라 변경 구간 발췌만 표시한다.
    assert any("—수정:" in text for text in body_19)
    # 발췌는 전체 문단보다 짧아야 한다
    assert any(
        len(text) < len("—수정: 하왕이 폐세자를 서인으로 보내었다.") for text in body_19 if text.startswith("—수정:")
    )
    assert any("번역: 한문 번역문" in text for text in body_19)
    assert any("인용 불일치" in text for text in body_19)
    # 마지막 연도(20년) 리포트는 본문 끝에
    assert texts.index("〔태조 20년(계유) 검수 리포트〕") > texts.index(
        "태종이 하륜을 문하부에 보내였다."
    )


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


async def test_pdf_document_review_end_to_end(temporal_env, monkeypatch) -> None:
    pdf_text = make_pdf_text()
    uploaded: list[str] = []

    @sdk_workflows.activity(name="pdf-find-drive-file")
    async def fake_find_pdf_file(file_name: str) -> DriveFileInfo:
        return DriveFileInfo(file_id="pdf-id", title="조선사.pdf", parent_id="folder-1")

    @sdk_workflows.activity(name="pdf-load-pages")
    async def fake_load_pages(
        file_id: str, page_start: int | None, page_end: int | None
    ) -> tuple[list[tuple[int, str]], int]:
        pages = [
            (number, f"{number}번째 페이지 첫 문단입니다. 조선의 사적을 따른다.\n\n{number}번째 페이지 둘째 문단, 태조 원년의 기사이다.")
            for number in range(8, 16)
        ]
        return pages, 143

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
        first_page: int,
        last_page: int,
        blocks: list,
        total_pages: int,
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

    @sdk_workflows.activity(name="pdf-annotate-footnote-sources")
    async def fake_annotate_footnote_sources(blocks: list) -> list:
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
    monkeypatch.setattr(review_module, "load_pages", fake_load_pages)
    monkeypatch.setattr(review_module, "annotate_source_titles", fake_annotate_source_titles)
    monkeypatch.setattr(review_module, "compare_quotes", fake_compare_quotes)
    monkeypatch.setattr(review_module, "annotate_footnote_sources", fake_annotate_footnote_sources)
    monkeypatch.setattr(review_module, "upload_drive_file", fake_upload_drive_file)
    monkeypatch.setattr(review_module, "chat_parse_to_model", fake_parse)
    monkeypatch.setattr(review_module, "build_result_docx", fake_build_result)

    async with create_test_worker(
        temporal_env,
        workflows=[PdfDocumentReviewWorkflow],
        activities=[
            fake_find_pdf_file,
            fake_load_pages,
            fake_upload_drive_file,
            fake_build_result,
            fake_annotate_source_titles,
            fake_compare_quotes,
            fake_annotate_footnote_sources,
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
    assert output.stats["total_pages"] == 143
    assert output.stats["body_paragraphs"] == 16
    assert output.stats["body_edits"] == 1
    assert "8~15페이지" in output.summary
    assert "내용 차이" in output.summary
    assert uploaded == ["조선사_p8-15_검수결과.docx"]


def test_excerpt_edit_pair_shows_only_changed_span() -> None:
    from worker.workflows.pdf_document_review import _excerpt_edit_pair

    original = (
        "왕이 삼군을 거느리고 개경으로 향하였다. 적병이 크게 궤멸되었다. "
        "장군 흔달·견달 등 3천 2백인을 사로잡고, 5천 7백 급을 참수하였다."
    )
    revised = original.replace("3천 2백인", "3천 2백 명")
    original_excerpt, revised_excerpt = _excerpt_edit_pair(original, revised)

    # 전체 문단이 아니라 변경 구간 근처만 나와야 한다.
    assert len(original_excerpt) <= 160
    assert "3천 2백인" in original_excerpt
    assert "3천 2백 명" in revised_excerpt
    assert "왕이 삼군을 거느리고 개경으로 향하였다" not in original_excerpt


def test_excerpt_edit_pair_handles_identical_text() -> None:
    from worker.workflows.pdf_document_review import _excerpt_edit_pair

    original_excerpt, revised_excerpt = _excerpt_edit_pair("동일 문단", "동일 문단")
    assert original_excerpt == revised_excerpt == "동일 문단"


def test_year_report_edits_use_excerpts_not_full_paragraphs() -> None:
    from worker.workflows.pdf_document_review import _year_reports

    long_text = (
        "앞부분 문맥이 매우 길게 이어지는 문단이다. " * 5
        + "중간의 오타를 고친다. "
        + "뒷부분 문맥도 매우 길게 이어진다. " * 5
    )
    section = {
        "year": "태조 19년(임신)",
        "edits": [
            TextEdit(
                item_id="p1",
                original=long_text.replace("오타를", "오타아를"),
                revised=long_text,
                category="typo",
                reason="오타",
            )
        ],
        "translations": [],
        "differences": [],
        "annotations": [],
    }
    lines = _year_reports(section)
    joined = "\n".join(lines)
    assert "—원문: " in joined and "—수정: " in joined
    # 문단 전체가 그대로 들어가면 안 된다.
    assert long_text not in joined
    assert "오타아를" in joined
    assert "오타를 고친다" in joined
