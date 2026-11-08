"""전 권 자동 검수 워크플로우 (한 번 실행으로 664페이지 전체).

내부에서 페이지 구간(약 24페이지씩)을 순회하며 기존 검수 파이프라인
(교정·번역·출처 + 품질 검증기)을 구간마다 실행하고, 검증을 통과한
결과 DOCX를 구간별로 자동 업로드한다. 사용자는 실행 1회로 전 권의
초안을 받는다. 품질 수리(스킬)는 이 초안 위에서 수행한다.
"""

from __future__ import annotations

import mistralai.workflows as workflows
import pydantic
from mistralai.workflows import Depends

from worker.workflows.hwp_review_models import DriveFileInfo, PdfReviewOutput
from worker.workflows.pdf_document_review import (
    PdfReviewInput,
    PdfDocumentReviewWorkflow,
    annotate_footnote_sources,
    blocks_from_pages,
    build_result_docx,
    find_pdf_file,
    group_by_year,
    load_pages,
    proofread_body,
    reconcile_edits,
    reconcile_translations,
    translate_hanja_blocks,
    upload_drive_file,
)

SEGMENT_SIZE = 24
FIRST_PAGE = 1
LAST_PAGE = 664


class FullBookInput(pydantic.BaseModel):
    file_name: str
    output_parent_id: str | None = None


class FullBookOutput(pydantic.BaseModel):
    result_files: list[DriveFileInfo]
    stats: dict[str, int]
    summary: str


@workflows.workflow.define(name="full-book-review", on_behalf_of=True)
class FullBookReviewWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: FullBookInput) -> FullBookOutput:
        source = await find_pdf_file(data.file_name)
        stem = source.title.rsplit(".", 1)[0] if "." in source.title else source.title
        output_parent_id = data.output_parent_id or source.parent_id

        result_files: list[DriveFileInfo] = []
        failed_segments: list[str] = []

        start = FIRST_PAGE
        while start <= LAST_PAGE:
            end = min(start + SEGMENT_SIZE - 1, LAST_PAGE)
            try:
                pages, total_pages = await load_pages(source.file_id, start, end)
                if not pages:
                    start = end + 1
                    continue
                blocks = blocks_from_pages(pages)

                body_edits = await proofread_body(blocks)
                hanja_translations = await translate_hanja_blocks(blocks)
                source_annotations = await annotate_footnote_sources(blocks)
                # 본문-주석 내용 차이는 스킬(Work)이 수리 단계에서 담당한다.

                body_edits = reconcile_edits(blocks, body_edits)
                hanja_translations = reconcile_translations(blocks, hanja_translations)
                source_annotations = [
                    annotation
                    for annotation in source_annotations
                    if annotation.item_id in {block.block_id for block in blocks}
                ]
                year_sections = group_by_year(
                    blocks,
                    body_edits,
                    hanja_translations,
                    [],
                    source_annotations,
                )
                stats = {
                    "body_paragraphs": len(blocks),
                    "body_edits": len(body_edits),
                    "hanja_translations": len(hanja_translations),
                    "quote_differences": 0,
                    "source_annotations": len(source_annotations),
                    "total_pages": total_pages,
                    "page_start": start,
                    "page_end": end,
                }
                result_b64, verification_errors = await build_result_docx(
                    source,
                    pages[0][0],
                    pages[-1][0],
                    blocks,
                    total_pages,
                    year_sections,
                    stats,
                )
                if verification_errors:
                    # 검증 실패 구간은 업로드하지 않고 기록만 남긴다.
                    failed_segments.append(
                        f"p{start}-{end}: " + "; ".join(verification_errors[:3])
                    )
                else:
                    result_file = await upload_drive_file(
                        f"{stem}_p{start}-{end}_검수초안.docx",
                        result_b64,
                        output_parent_id,
                    )
                    result_files.append(result_file)
            except Exception as error:
                failed_segments.append(f"p{start}-{end}: {type(error).__name__}")

            start = end + 1

        return FullBookOutput(
            result_files=result_files,
            stats={
                "segments_done": len(result_files),
                "segments_failed": len(failed_segments),
            },
            summary=(
                f"전 권 초안 {len(result_files)}개 구간 완료. "
                f"실패 {len(failed_segments)}개: {' | '.join(failed_segments[:5])}"
                if failed_segments
                else f"전 권 초안 {len(result_files)}개 구간 완료."
            ),
        )
