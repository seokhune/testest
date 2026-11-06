"""PDF 문서 검수 워크플로우.

Google Drive의 PDF 파일에서 텍스트 추출 도구로 본문을 읽어 지정한 페이지
범위만 교정(오타, 문법, 띄어쓰기, 한문 번역)하고, 교정본 DOCX와 검수
리포트 DOCX를 만들어 Drive에 업로드한다.

PDF 바이너리를 다운로드하지 않고 텍스트 표현을 읍어 오므로 커넥터의
base64 2MB 전송 한도보다 큰 파일도 처리할 수 있다.
"""

from __future__ import annotations

import re
from datetime import timedelta
from typing import Any

import mistralai.workflows as workflows
import pydantic
from mistralai.workflows import Depends
from mistralai.workflows.plugins.mistralai import (
    ChatCompletionRequest,
    SystemMessage,
    UserMessage,
    chat_parse_to_model,
)
from mistralai.workflows.plugins.mistralai.connectors import (
    ToolCallClient,
    connector,
    uses_connectors,
)

from worker.workflows.hwp_review_models import (
    BodyBlock,
    DriveFileInfo,
    HanjaTranslation,
    PdfReviewInput,
    PdfReviewOutput,
    TextEdit,
)
from worker.workflows.hwp_review_docx import (
    apply_edits_by_text,
    build_structured_docx,
    contains_hanja,
    convert_pdf_to_docx,
)

REVIEW_MODEL = "mistral-medium-latest"
LLM_CHUNK_SIZE = 8
LLM_TIMEOUT = timedelta(seconds=300)
IO_TIMEOUT = timedelta(seconds=120)
CONVERT_TIMEOUT = timedelta(minutes=10)
MAX_TEXT_CHARS = 1_900_000
MAX_TRANSFER_RAW_BYTES = 1_500_000

drive = connector("google_drive_mcp")

PAGE_MARKER_RE = re.compile(r"^\s*-\s*(\d+)\s*-\s*$")


def _tool_result_value(result: Any) -> Any:
    import json

    for item in result.get("content") or []:
        item_type = item.get("type")
        if item_type == "text" and item.get("text"):
            try:
                return json.loads(item["text"])
            except json.JSONDecodeError:
                return item["text"]
        if item_type == "resource":
            resource = item.get("resource") or {}
            text = resource.get("text")
            if text:
                try:
                    return json.loads(text)
                except json.JSONDecodeError:
                    return text
    return None


def _chunk(items: list, size: int) -> list[list]:
    return [items[start : start + size] for start in range(0, len(items), size)]


def split_pdf_pages(text: str) -> list[tuple[int, str]]:
    """PDF 텍스트를 페이지 표식(`- 12 -` 등) 기준으로 나눈다."""
    pages: dict[int, list[str]] = {1: []}
    current = 1
    for line in text.splitlines():
        marker = PAGE_MARKER_RE.match(line)
        if marker:
            current = int(marker.group(1))
            pages.setdefault(current, [])
        else:
            pages.setdefault(current, []).append(line)
    return [
        (number, "\n".join(lines).strip())
        for number, lines in sorted(pages.items())
        if "\n".join(lines).strip()
    ]


def slice_pages(
    text: str,
    page_start: int | None = None,
    page_end: int | None = None,
) -> tuple[list[tuple[int, str]], int]:
    """PDF 텍스트를 페이지 단위로 슬라이싱한다."""
    pages = split_pdf_pages(text)
    total_pages = max((number for number, _ in pages), default=1)
    selected = pages
    if page_start is not None or page_end is not None:
        start = page_start or 1
        end = page_end or total_pages
        selected = [
            (number, page_text) for number, page_text in pages if start <= number <= end
        ]
        if not selected:
            raise workflows.WorkflowError(
                f"지정한 페이지 범위({start}~{end})의 본문을 찾지 못했습니다. "
                f"PDF에 페이지 표식(- 8 - 형태)이 있는지 확인해 주세요.",
                non_retryable=True,
            )
    return selected, total_pages


def extract_page_blocks(
    text: str,
    page_start: int | None = None,
    page_end: int | None = None,
) -> tuple[list[BodyBlock], int]:
    """텍스트를 페이지별 문단 블록으로 나누고, 페이지 범위로 슬라이싱한다."""
    selected, total_pages = slice_pages(text, page_start, page_end)
    blocks: list[BodyBlock] = []
    for page_number, page_text in selected:
        for paragraph in re.split(r"\n\s*\n", page_text):
            cleaned = paragraph.strip()
            if cleaned:
                blocks.append(
                    BodyBlock(
                        block_id=f"p{len(blocks)}", text=cleaned, page=page_number
                    )
                )
    return blocks, total_pages


class _EditBatch(pydantic.BaseModel):
    edits: list[TextEdit]


class _HanjaBatch(pydantic.BaseModel):
    translations: list[HanjaTranslation]


BODY_PROOFREAD_PROMPT = """아래는 번역서 본문 문단 목록입니다.

규칙:
1. 오타자(맞춤법, 띄어쓰기, 오타)를 수정하십시오.
2. 문법에 맞게 윤문되되, 번역된 내용(의미)은 절대 바꾸지 마십시오.
3. 변경이 필요한 문단만 edits에 포함하십시오. 변경이 없으면 빈 목록을 반환하십시오.
4. revised는 해당 문단의 전체 최종 텍스트여야 합니다.

각 항목은 item_id(대괄호 안 문단 번호), original(원문 전체), revised(수정문 전체), category(typo|grammar|spacing|style 중 하나), reason(변경 이유 한 문장)을 포함합니다.

문단 목록:
"""

HANJA_TRANSLATE_PROMPT = """아래는 번역서 문단 목록입니다. 한문이 포함된 문단만 골라, 한문 구절의 한글 번역을 만드십시오.

규칙:
1. 한문 구절을 자연스러운 한국어로 번역하십시오.
2. 한문이 없는 문단은 포함하지 마십시오. 결과가 없으면 빈 목록을 반환하십시오.
3. translation에는 번역만 담고 원문 그대로는 포함하지 마십시오.

각 항목은 item_id(대괄호 안 문단 번호), original(한문 구절이 포함된 문단 전체), translation(한문 번역)을 포함합니다.

문단 목록:
"""


async def _parse(model_class: type, prompt: str) -> Any:
    request = ChatCompletionRequest(
        model=REVIEW_MODEL,
        messages=[
            SystemMessage(content="당신은 한국어 번역 학술서 교정 전문가입니다."),
            UserMessage(content=prompt),
        ],
    )
    return await chat_parse_to_model(model_class, request)


@workflows.activity(name="pdf-find-drive-file", start_to_close_timeout=IO_TIMEOUT)
async def find_pdf_file(
    file_name: str, client: ToolCallClient = Depends(drive)
) -> DriveFileInfo:
    result = await client.call_tool(
        tool_name="search_files",
        arguments={"query": f"title contains '{file_name}'", "page_size": 25},
    )
    value = _tool_result_value(result)
    files = value.get("files") if isinstance(value, dict) else value
    files = files or []
    candidates = [
        file
        for file in files
        if (file.get("title") or "").lower().endswith(".pdf")
        and file_name.lower() in (file.get("title") or "").lower()
    ]
    if not candidates:
        raise workflows.WorkflowError(
            f"드라이브에서 PDF 문서를 찾지 못했습니다: {file_name}",
            non_retryable=True,
        )
    file = candidates[0]
    return DriveFileInfo(
        file_id=file.get("id") or "",
        title=file.get("title") or "",
        parent_id=file.get("parent_id"),
        mime_type=file.get("mime_type"),
        view_url=file.get("view_url"),
        file_size=int(file.get("file_size") or 0) or None,
    )


@workflows.activity(name="pdf-read-text", start_to_close_timeout=IO_TIMEOUT)
async def read_pdf_text(
    file_id: str, client: ToolCallClient = Depends(drive)
) -> str:
    result = await client.call_tool(
        tool_name="read_file_content",
        arguments={"file_id": file_id},
    )
    value = _tool_result_value(result)
    text = value.get("file_content") if isinstance(value, dict) else value
    if not isinstance(text, str) or not text.strip():
        raise workflows.WorkflowError(
            "PDF에서 텍스트를 추출하지 못했습니다. 텍스트 레이어가 있는 PDF인지 확인해 주세요.",
            non_retryable=True,
        )
    if len(text) > MAX_TEXT_CHARS:
        raise workflows.WorkflowError(
            f"PDF 텍스트가 너무 깁니다({len(text):,}자). 파일을 나눠 업로드해 주세요.",
            non_retryable=True,
        )
    return text


@workflows.activity(name="pdf-proofread-body", start_to_close_timeout=LLM_TIMEOUT)
async def proofread_body(blocks: list[BodyBlock]) -> list[TextEdit]:
    edits: list[TextEdit] = []
    for chunk in _chunk(blocks, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        result = await _parse(_EditBatch, BODY_PROOFREAD_PROMPT + listing)
        edits.extend(result.edits)
    return edits


@workflows.activity(name="pdf-translate-hanja", start_to_close_timeout=LLM_TIMEOUT)
async def translate_hanja_blocks(blocks: list[BodyBlock]) -> list[HanjaTranslation]:
    targets = [block for block in blocks if contains_hanja(block.text)]
    translations: list[HanjaTranslation] = []
    for chunk in _chunk(targets, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        result = await _parse(_HanjaBatch, HANJA_TRANSLATE_PROMPT + listing)
        translations.extend(result.translations)
    return translations


@workflows.activity(name="pdf-build-outputs", start_to_close_timeout=IO_TIMEOUT)
async def build_output_docx(
    source_title: str,
    blocks: list[BodyBlock],
    body_edits: list[TextEdit],
    hanja_translations: list[HanjaTranslation],
) -> tuple[str, str]:
    """교정본 DOCX와 검수 리포트 DOCX를 만들어 base64로 반환한다."""
    import base64
    import io

    from docx import Document

    edits_by_id = {edit.item_id: edit for edit in body_edits}
    hanja_by_id = {translation.item_id: translation for translation in hanja_translations}

    corrected = Document()
    corrected.add_heading(f"{source_title} 교정본", level=1)
    for block in blocks:
        edit = edits_by_id.get(block.block_id)
        corrected.add_paragraph(edit.revised if edit else block.text)
        translation = hanja_by_id.get(block.block_id)
        if translation:
            corrected.add_paragraph(f"번역: {translation.translation}")

    report = Document()
    report.add_heading(f"{source_title} 검수 리포트", level=1)
    report.add_heading("1. 교정 요약", level=2)
    report.add_paragraph(f"검수 문단 {len(blocks)}개 중 {len(body_edits)}개 수정.")
    if hanja_translations:
        report.add_paragraph(f"한문 번역 {len(hanja_translations)}건 추가.")
    report.add_heading("2. 교정 내역", level=2)
    if not body_edits:
        report.add_paragraph("수정된 문단이 없습니다.")
    for edit in body_edits:
        report.add_paragraph(f"[{edit.item_id}] ({edit.category}) {edit.reason}")
        report.add_paragraph(f"원문: {edit.original}")
        report.add_paragraph(f"수정: {edit.revised}")

    def _to_b64(document: Document) -> str:
        buffer = io.BytesIO()
        document.save(buffer)
        return base64.b64encode(buffer.getvalue()).decode()

    return _to_b64(corrected), _to_b64(report)


@workflows.activity(
    name="pdf-build-inplace", start_to_close_timeout=CONVERT_TIMEOUT
)
async def build_inplace_edition(
    source: DriveFileInfo,
    pages: list[tuple[int, str]],
    blocks: list[BodyBlock],
    total_pages: int,
    body_edits: list[TextEdit],
    hanja_translations: list[HanjaTranslation],
    client: ToolCallClient = Depends(drive),
) -> tuple[str, int]:
    """원본 PDF 양식을 그대로 유지한 수정반영본을 만든다.

    다운로드→변환→교정을 한 액티비티 안에서 처리하고 작은 DOCX만
    반환한다(액티비티 결과가 base64 2MB 이벤트 한도를 넘지 않게).
    검수 범위에 해당하는 물리 페이지만 pdf2docx로 변환한다. 변환에
    실패하면 원본 페이지 구조를 재구성한 DOCX로 대체한다.
    """
    import asyncio
    import base64

    try:
        result = await client.call_tool(
            tool_name="download_file_content",
            arguments={"file_id": source.file_id},
        )
        value = _tool_result_value(result)
        content = value.get("content") if isinstance(value, dict) else value
        if not content:
            raise RuntimeError("PDF 다운로드 결과가 비었습니다")
        pdf_bytes = base64.b64decode(content)

        start = (pages[0][0] - 1) if pages else 0
        end = pages[-1][0] if pages else None
        docx_bytes = await asyncio.to_thread(
            convert_pdf_to_docx, pdf_bytes, start, end
        )
        edited, applied = await asyncio.to_thread(
            apply_edits_by_text, docx_bytes, body_edits, hanja_translations, blocks
        )
        return base64.b64encode(edited).decode(), applied
    except Exception as error:
        import logging

        logging.getLogger(__name__).warning(
            "원본 양식 변환 실패, 구조 재구성으로 대체: %s", error
        )
    edited, applied = await asyncio.to_thread(
        build_structured_docx, pages, blocks, body_edits, hanja_translations
    )
    return base64.b64encode(edited).decode(), applied


@workflows.activity(
    name="pdf-upload-drive-file",
    start_to_close_timeout=IO_TIMEOUT,
    retry_policy_max_attempts=1,
)
async def upload_drive_file(
    title: str,
    content_b64: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> DriveFileInfo:
    arguments: dict[str, Any] = {
        "title": title,
        "base64_content": content_b64,
        "content_mime_type": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        "disable_conversion_to_google_type": True,
    }
    if parent_id:
        arguments["parent_id"] = parent_id
    result = await client.call_tool(tool_name="create_file", arguments=arguments)
    file = _tool_result_value(result) or {}
    if not file.get("id"):
        raise workflows.WorkflowError("드라이브 업로드에 실패했습니다")
    return DriveFileInfo(
        file_id=file.get("id") or "",
        title=file.get("title") or title,
        parent_id=file.get("parent_id"),
        view_url=file.get("view_url"),
    )


@workflows.workflow.define(name="pdf-document-review", on_behalf_of=True)
@uses_connectors(drive)
class PdfDocumentReviewWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: PdfReviewInput) -> PdfReviewOutput:
        source = await find_pdf_file(data.file_name)
        stem = source.title.rsplit(".", 1)[0] if "." in source.title else source.title
        output_parent_id = data.output_parent_id or source.parent_id

        text = await read_pdf_text(source.file_id)
        pages, total_pages = slice_pages(text, data.page_start, data.page_end)
        blocks = [
            BodyBlock(block_id=f"p{index}", text=paragraph.strip(), page=page_number)
            for index, (page_number, page_text) in enumerate(pages)
            for paragraph in re.split(r"\n\s*\n", page_text)
            if paragraph.strip()
        ]

        body_edits = await proofread_body(blocks)
        hanja_translations = await translate_hanja_blocks(blocks)

        corrected_b64, report_b64 = await build_output_docx(
            source.title, blocks, body_edits, hanja_translations
        )

        range_suffix = ""
        if data.page_start is not None or data.page_end is not None:
            range_start = data.page_start or 1
            range_end = data.page_end or total_pages
            range_suffix = f"_p{range_start}-{range_end}"

        corrected = await upload_drive_file(
            f"{stem}{range_suffix}_교정본.docx", corrected_b64, output_parent_id
        )
        report = await upload_drive_file(
            f"{stem}{range_suffix}_검수리포트.docx", report_b64, output_parent_id
        )

        inplace_b64, applied = await build_inplace_edition(
            source, pages, blocks, total_pages, body_edits, hanja_translations
        )
        inplace_file = await upload_drive_file(
            f"{stem}{range_suffix}_수정반영본.docx",
            inplace_b64,
            output_parent_id,
        )
        inplace_note = f" 수정반영본에 {applied}건을 반영했습니다."

        page_note = ""
        if range_suffix:
            page_note = (
                f" (총 {total_pages}페이지 중 "
                f"{data.page_start or 1}~{data.page_end or total_pages}페이지)"
            )
        stats = {
            "body_paragraphs": len(blocks),
            "body_edits": len(body_edits),
            "hanja_translations": len(hanja_translations),
            "total_pages": total_pages,
            "page_start": data.page_start or 0,
            "page_end": data.page_end or 0,
        }
        summary = (
            f"본문 {len(blocks)}문단을 검수해{page_note} "
            f"{len(body_edits)}건을 교정하고 한문 번역 {len(hanja_translations)}건을 추가했습니다."
        )
        return PdfReviewOutput(
            source_title=source.title,
            corrected_file=corrected,
            report_file=report,
            inplace_file=inplace_file,
            stats=stats,
            summary=summary + inplace_note,
        )
