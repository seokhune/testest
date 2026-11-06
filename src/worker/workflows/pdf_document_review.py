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
    BodyFootnoteDifference,
    DriveFileInfo,
    HanjaTranslation,
    PdfReviewInput,
    PdfReviewOutput,
    SourceAnnotation,
    TextEdit,
)
from worker.workflows.hwp_review_docx import (
    append_review_report,
    apply_edits_by_text,
    build_structured_docx,
    contains_hanja,
    convert_pdf_to_docx,
    unwrap_text_frames,
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


YEAR_RE = re.compile(r"\d{1,3}\s*년\s*\([^)]*\)|\d{1,3}\s*년|원년")


def detect_year(text: str) -> str | None:
    """문단에서 년도 표기(예: '태종 6년(무인)')를 찾아 라벨로 반환한다."""
    match = YEAR_RE.search(text[:60])
    if not match:
        return None
    prefix = text[: match.start()]
    king = re.search(r"[가-힣]{1,4}\s*$", prefix)
    return ((king.group(0) if king else "") + match.group(0)).strip()


def group_by_year(
    blocks: list[BodyBlock],
    body_edits: list[TextEdit],
    hanja_translations: list[HanjaTranslation],
    differences: list[BodyFootnoteDifference] | None = None,
    source_annotations: list[SourceAnnotation] | None = None,
) -> list[tuple]:
    """블록 순서대로 년도 구간을 따라가며 검수 결과를 년도별로 묶는다."""
    sections: list[tuple] = []
    current = "기타"
    edits_by_block: dict[str, list[TextEdit]] = {}
    for edit in body_edits:
        edits_by_block.setdefault(edit.item_id, []).append(edit)
    translations_by_block: dict[str, list[HanjaTranslation]] = {}
    for translation in hanja_translations:
        translations_by_block.setdefault(translation.item_id, []).append(translation)
    differences_by_block: dict[str, list[BodyFootnoteDifference]] = {}
    for difference in differences or []:
        differences_by_block.setdefault(difference.body_block_id or "", []).append(
            difference
        )
    annotations_by_block: dict[str, list[SourceAnnotation]] = {}
    for annotation in source_annotations or []:
        annotations_by_block.setdefault(annotation.item_id, []).append(annotation)
    section_index: dict[str, int] = {}

    def section_for(label: str) -> int:
        if label not in section_index:
            section_index[label] = len(sections)
            sections.append((label, [], [], [], []))
        return section_index[label]

    for block in blocks:
        year = detect_year(block.text)
        if year:
            current = year
        index = section_for(current)
        edits, translations, diffs, annotations = (
            sections[index][1],
            sections[index][2],
            sections[index][3],
            sections[index][4],
        )
        edits.extend(edits_by_block.get(block.block_id, []))
        translations.extend(translations_by_block.get(block.block_id, []))
        diffs.extend(differences_by_block.pop(block.block_id, []))
        annotations.extend(annotations_by_block.pop(block.block_id, []))
    if differences_by_block:
        index = section_for("기타")
        for leftover in differences_by_block.values():
            sections[index][3].extend(leftover)
    if annotations_by_block:
        index = section_for("기타")
        for leftover in annotations_by_block.values():
            sections[index][4].extend(leftover)
    return sections


def _norm_text(text: str) -> str:
    return "".join(text.split())


def reconcile_edits(
    blocks: list[BodyBlock], edits: list[TextEdit]
) -> list[TextEdit]:
    """LLM 교정 항목을 블록에 정합한다.

    LLM이 실제 블록과 다른 item_id를 붙일 수 있어, item_id가 없으면
    원문 텍스트 일치로 블록을 찾아 id를 다시 매긴다. 매칭 실패 항목은
    버리고, 같은 블록에 여러 항목이 있으면 마지막 것을 유지한다.
    """
    by_id = {block.block_id: block for block in blocks}
    by_text = {_norm_text(block.text): block for block in blocks}
    matched: dict[str, TextEdit] = {}
    for edit in edits:
        block = by_id.get(edit.item_id)
        if block is None:
            block = by_text.get(_norm_text(edit.original))
        if block is None:
            needle = _norm_text(edit.original)[:60]
            if needle:
                for candidate in blocks:
                    if _norm_text(candidate.text).startswith(needle):
                        block = candidate
                        break
        if block is None:
            continue
        matched[block.block_id] = edit.model_copy(update={"item_id": block.block_id})
    return list(matched.values())


def reconcile_translations(
    blocks: list[BodyBlock], translations: list[HanjaTranslation]
) -> list[HanjaTranslation]:
    """LLM 한문 번역 항목을 블록에 정합한다 (reconcile_edits와 같은 규칙)."""
    by_id = {block.block_id: block for block in blocks}
    by_text = {_norm_text(block.text): block for block in blocks}
    matched: dict[str, HanjaTranslation] = {}
    for translation in translations:
        block = by_id.get(translation.item_id)
        if block is None:
            block = by_text.get(_norm_text(translation.original))
        if block is None:
            needle = _norm_text(translation.original)[:60]
            if needle:
                for candidate in blocks:
                    if _norm_text(candidate.text).startswith(needle):
                        block = candidate
                        break
        if block is None:
            continue
        matched[block.block_id] = translation.model_copy(
            update={"item_id": block.block_id}
        )
    return list(matched.values())


class _EditBatch(pydantic.BaseModel):
    edits: list[TextEdit]


class _HanjaBatch(pydantic.BaseModel):
    translations: list[HanjaTranslation]


class _DifferenceBatch(pydantic.BaseModel):
    differences: list[BodyFootnoteDifference]


class _SourceBatch(pydantic.BaseModel):
    annotations: list[SourceAnnotation]


BODY_PROOFREAD_PROMPT = """아래는 번역서 본문 문단 목록입니다.

규칙:
1. 오타자(맞춤법, 띄어쓰기, 오타)를 수정하십시오.
2. 문법에 맞게 윤문되되, 번역된 내용(의미)은 절대 바꾸지 마십시오.
3. 주석 번호(1), 2) 등)와 주석 인용문, 괄호 안 출처 표기는 절대 삭제하거나 변경하지 마십시오.
4. revised는 해당 문단의 전체 최종 텍스트여야 하며, 원문의 모든 내용(주석 포함)을 빠짐없이 포함해야 합니다.
5. 변경이 필요한 문단만 edits에 포함하십시오. 변경이 없으면 빈 목록을 반환하십시오.

각 항목은 item_id(대괄호 안 문단 번호), original(원문 전체), revised(수정문 전체), category(typo|grammar|spacing|style 중 하나), reason(변경 이유 한 문장)을 포함합니다.

문단 목록:
"""

HANJA_TRANSLATE_PROMPT = """아래는 번역서 문단 목록입니다. 한문(한자문)으로 작성된 주석만 골라 자연스러운 한국어로 번역하십시오.

규칙:
1. 한문 구절을 자연스러운 한국어로 번역하십시오.
2. 한자가 일부 포함되었어도 대부분 한국어라면 번역하지 말고 건너뛰십시오. 이미 한국어로만 작성된 주석도 건너뛰십시오.
3. 결과가 없으면 빈 목록을 반환하십시오.
4. translation에는 번역만 담고 원문 그대로는 포함하지 마십시오.

각 항목은 item_id(대괄호 안 문단 번호), original(한문 구절이 포함된 문단 전체), translation(한문 번역)을 포함합니다.

문단 목록:
"""

SOURCE_TITLE_PROMPT = """아래는 번역서 문단 목록입니다. 주석의 서지명 표기를 교정하십시오.

규칙:
1. 책·논문·신문 등 서지명은 『 』(겹낫표)로 감싸십시오(예: 『조선왕조실록』). 『 』나 다른 괄호로 표기된 서지명도 『 』로 통일하고, category를 book_title로 지정하십시오.
2. 서지명이 아닌 인용구에는 적용하지 마십시오.
3. 교정이 필요한 문단만 edits에 포함하십시오. 필요 없으면 빈 목록을 반환하십시오.
4. revised는 해당 문단의 전체 최종 텍스트여야 하며, 서지명 표기 외의 내용은 절대 바꾸지 마십시오.
5. 주석 번호와 주석 인용문은 절대 삭제하지 마십시오.

각 항목은 item_id(대괄호 안 문단 번호), original(원문 전체), revised(수정문 전체), category(book_title 고정), reason(변경 이유 한 문장)을 포함합니다.

문단 목록:
"""

FOOTNOTE_SOURCE_PROMPT = """아래는 번역서 문단 목록입니다. 주석의 출처(서지 정보)를 확인하십시오.

규칙:
1. 각 주석의 내용(근거문)을 근거로 출처(저자, 서지명, 출판사, 출판 연도, 쪽수 등)를 판독할 수 있을 때만 annotations에 포함하십시오.
2. 근거가 불확실하면 source 끝에 "(추정)"을 붙이십시오.
3. 근거를 전혀 찾을 수 없거나 출처가 이미 완전하게 표기된 주석은 건너뛰십시오.
4. source는 한국어 서지 표기 관례를 따르십시오.

각 항목은 item_id(대괄호 안 문단 번호), source(보완할 출처 표기)을 포함합니다.

문단 목록:
"""

QUOTE_DIFFERENCE_PROMPT = """아래는 번역서 문단 목록입니다. 본문 내용과 주석 내용을 비교하여 다음을 찾아 differences에 정리하십시오.

- 주석의 인용이 본문의 특정 표현과 맞지 않는 경우(인용 불일치, 대상 불일치, 수치·연도 불일치 등)
- 본문은 언급했는데 주석이 관련 내용을 담고 있지 않은 경우
- 주석과 본문의 서술이 상충하는 경우

차이가 없으면 빈 목록을 반환하십시오.

각 항목은 body_block_id(대괄호 안 문단 번호), footnote_id(해당 주석 번호, 모르면 빈 문자열), difference(차이점에 대한 구체적 서술)을 포함합니다.

문단 목록:
"""


def _log_llm_failure(pass_name: str, error: Exception) -> None:
    import logging

    logging.getLogger(__name__).warning(
        "LLM 응답 파싱 실패(%s), 해당 청크 건너뜀: %s",
        pass_name,
        str(error)[:500],
    )


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
        try:
            result = await _parse(_EditBatch, BODY_PROOFREAD_PROMPT + listing)
            edits.extend(result.edits)
        except Exception as error:
            _log_llm_failure("본문 교정", error)
    return edits


@workflows.activity(name="pdf-translate-hanja", start_to_close_timeout=LLM_TIMEOUT)
async def translate_hanja_blocks(blocks: list[BodyBlock]) -> list[HanjaTranslation]:
    targets = [block for block in blocks if contains_hanja(block.text)]
    translations: list[HanjaTranslation] = []
    for chunk in _chunk(targets, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        try:
            result = await _parse(_HanjaBatch, HANJA_TRANSLATE_PROMPT + listing)
            translations.extend(result.translations)
        except Exception as error:
            _log_llm_failure("한문 번역", error)
    return translations


@workflows.activity(name="pdf-annotate-source-titles", start_to_close_timeout=LLM_TIMEOUT)
async def annotate_source_titles(blocks: list[BodyBlock]) -> list[TextEdit]:
    edits: list[TextEdit] = []
    for chunk in _chunk(blocks, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        try:
            result = await _parse(_EditBatch, SOURCE_TITLE_PROMPT + listing)
            edits.extend(result.edits)
        except Exception as error:
            _log_llm_failure("서지명 표기", error)
    return edits


@workflows.activity(name="pdf-compare-quotes", start_to_close_timeout=LLM_TIMEOUT)
async def compare_quotes(blocks: list[BodyBlock]) -> list[BodyFootnoteDifference]:
    differences: list[BodyFootnoteDifference] = []
    for chunk in _chunk(blocks, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        try:
            result = await _parse(_DifferenceBatch, QUOTE_DIFFERENCE_PROMPT + listing)
            differences.extend(result.differences)
        except Exception as error:
            _log_llm_failure("본문-주석 비교", error)
    return differences


@workflows.activity(name="pdf-annotate-footnote-sources", start_to_close_timeout=LLM_TIMEOUT)
async def annotate_footnote_sources(blocks: list[BodyBlock]) -> list[SourceAnnotation]:
    annotations: list[SourceAnnotation] = []
    for chunk in _chunk(blocks, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        try:
            result = await _parse(_SourceBatch, FOOTNOTE_SOURCE_PROMPT + listing)
            annotations.extend(result.annotations)
        except Exception as error:
            _log_llm_failure("주석 출처 판독", error)
    return annotations


@workflows.activity(name="pdf-build-result", start_to_close_timeout=CONVERT_TIMEOUT)
async def build_result_docx(
    source: DriveFileInfo,
    pages: list[tuple[int, str]],
    blocks: list[BodyBlock],
    total_pages: int,
    body_edits: list[TextEdit],
    hanja_translations: list[HanjaTranslation],
    year_sections: list[tuple[str, list, list]],
    stats: dict[str, int],
    client: ToolCallClient = Depends(drive),
) -> str:
    """검수 결과 한 파일(원본 양식 본문 + 년도별 리포트)을 만들어 base64로 반환한다.

    다운로드→변환→교정을 한 액티비티 안에서 처리하고 작은 DOCX만
    반환한다(액티비티 결과가 base64 2MB 이벤트 한도를 넘지 않게).
    원본 PDF를 pdf2docx로 변환해 양식을 유지하고, 그 뒤에 년도별
    교정 내역/한문 번역 리포트를 붙인다. 변환에 실패하면 페이지
    구조를 재구성한 본문으로 대체한다.
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
        docx_bytes = await asyncio.to_thread(unwrap_text_frames, docx_bytes)
        edited, _applied = await asyncio.to_thread(
            apply_edits_by_text, docx_bytes, body_edits, hanja_translations, blocks
        )
    except Exception as error:
        import logging

        logging.getLogger(__name__).warning(
            "원본 양식 변환 실패, 구조 재구성으로 대체: %s", error
        )
        edited, _applied = await asyncio.to_thread(
            build_structured_docx, pages, blocks, body_edits, hanja_translations
        )

    final = await asyncio.to_thread(
        append_review_report, edited, source.title, year_sections, stats
    )
    return base64.b64encode(final).decode()


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
        source_title_edits = await annotate_source_titles(blocks)
        hanja_translations = await translate_hanja_blocks(blocks)
        source_annotations = await annotate_footnote_sources(blocks)
        differences = await compare_quotes(blocks)

        body_edits = reconcile_edits(blocks, body_edits + source_title_edits)
        hanja_translations = reconcile_translations(blocks, hanja_translations)
        source_annotations = [
            annotation
            for annotation in source_annotations
            if annotation.item_id in {block.block_id for block in blocks}
        ]
        valid_ids = {block.block_id for block in blocks}
        differences = [
            difference
            for difference in differences
            if (difference.body_block_id or "") in valid_ids
        ]
        year_sections = group_by_year(
            blocks, body_edits, hanja_translations, differences, source_annotations
        )
        stats = {
            "body_paragraphs": len(blocks),
            "body_edits": len(body_edits),
            "hanja_translations": len(hanja_translations),
            "quote_differences": len(differences),
            "source_annotations": len(source_annotations),
            "total_pages": total_pages,
            "page_start": data.page_start or 0,
            "page_end": data.page_end or 0,
        }

        result_b64 = await build_result_docx(
            source,
            pages,
            blocks,
            total_pages,
            body_edits,
            hanja_translations,
            year_sections,
            stats,
        )

        range_suffix = ""
        if data.page_start is not None or data.page_end is not None:
            range_start = data.page_start or 1
            range_end = data.page_end or total_pages
            range_suffix = f"_p{range_start}-{range_end}"

        result_file = await upload_drive_file(
            f"{stem}{range_suffix}_검수결과.docx", result_b64, output_parent_id
        )

        page_note = ""
        if range_suffix:
            page_note = (
                f" (총 {total_pages}페이지 중 "
                f"{data.page_start or 1}~{data.page_end or total_pages}페이지)"
            )
        year_note = " · ".join(
            year
            for year, edits, translations, diffs, annotations in year_sections
            if edits or translations or diffs or annotations
        )
        summary = (
            f"본문 {len(blocks)}문단을 검수해{page_note} "
            f"{len(body_edits)}건을 교정하고 한문 번역 {len(hanja_translations)}건을 추가했으며, "
            f"본문-주석 인용 차이 {len(differences)}건과 주석 출처 보완 제안 "
            f"{len(source_annotations)}건을 확인했습니다. "
            f"년도 구간: {year_note}"
        )
        return PdfReviewOutput(
            source_title=source.title,
            result_file=result_file,
            stats=stats,
            summary=summary,
        )
