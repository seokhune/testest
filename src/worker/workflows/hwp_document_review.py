"""HWP/HWPX 번역문 검수 워크플로우.

Google Drive의 한컴 문서를 검색해 사본을 만들고, DOCX로 변환한 뒤
여섯 단계의 검수(본문 오탈자 교정, 본문 윤문, 주석 서지명 〈 〉 처리,
주석 출처 보완, 한문 주석 번역, 본문-주석 차이점 대조)를 수행하고
교정본 DOCX와 검수 리포트를 Drive에 업로드한다. 원본 파일은 수정하지
않는다.
"""

from __future__ import annotations

import asyncio
import base64
import json
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

from worker.workflows.hwp_review_docx import (
    DOCX_MIME_TYPE,
    HANCOM_EXTENSIONS,
    apply_review_edits,
    build_report_docx,
    contains_hanja,
    convert_hancom_to_docx,
    extract_document,
)
from worker.workflows.hwp_review_models import (
    BodyBlock,
    BodyFootnoteDifference,
    DriveFileInfo,
    ExtractedDocument,
    FootnoteItem,
    HanjaTranslation,
    HwpReviewInput,
    HwpReviewOutput,
    SourceAnnotation,
    TextEdit,
)

REVIEW_MODEL = "mistral-medium-latest"
LLM_CHUNK_SIZE = 8
COMPARE_BODY_CHUNK_SIZE = 20
COMPARE_FOOTNOTE_CHAR_LIMIT = 8000
LLM_TIMEOUT = timedelta(seconds=300)
IO_TIMEOUT = timedelta(seconds=120)
CONVERT_TIMEOUT = timedelta(minutes=10)

drive = connector("google_drive_mcp")

_SOFFICE_LOCK = asyncio.Lock()


def _tool_result_value(result: Any) -> Any:
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


def _truncate(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + "…(이하 생략)"


async def _parse(model_class: type, prompt: str) -> Any:
    request = ChatCompletionRequest(
        model=REVIEW_MODEL,
        messages=[
            SystemMessage(content="당신은 한국어 번역 학술서 교정 전문가입니다."),
            UserMessage(content=prompt),
        ],
    )
    return await chat_parse_to_model(model_class, request)


@workflows.activity(name="hwp-find-drive-file", start_to_close_timeout=IO_TIMEOUT)
async def find_drive_file(
    file_name: str, client: ToolCallClient = Depends(drive)
) -> DriveFileInfo:
    result = await client.call_tool(
        tool_name="search_files",
        arguments={"query": f"title contains '{file_name}'", "page_size": 25},
    )
    files = _tool_result_value(result) or []
    candidates = [
        file
        for file in files
        if (file.get("title") or "").lower().endswith(HANCOM_EXTENSIONS)
        and file_name.lower() in (file.get("title") or "").lower()
    ]
    if not candidates:
        raise workflows.WorkflowError(
            f"드라이브에서 한컴 문서를 찾지 못했습니다: {file_name}",
            non_retryable=True,
        )
    file = candidates[0]
    return DriveFileInfo(
        file_id=file.get("id") or "",
        title=file.get("title") or "",
        parent_id=file.get("parent_id"),
        mime_type=file.get("mime_type"),
        view_url=file.get("view_url"),
    )


@workflows.activity(
    name="hwp-copy-drive-file",
    start_to_close_timeout=IO_TIMEOUT,
    retry_policy_max_attempts=1,
)
async def copy_drive_file(
    file_id: str,
    title: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> DriveFileInfo:
    arguments: dict[str, Any] = {"file_id": file_id, "title": title}
    if parent_id:
        arguments["parent_id"] = parent_id
    result = await client.call_tool(tool_name="copy_file", arguments=arguments)
    file = _tool_result_value(result) or {}
    if not file.get("id"):
        raise workflows.WorkflowError("드라이브 파일 복제에 실패했습니다")
    return DriveFileInfo(
        file_id=file.get("id") or "",
        title=file.get("title") or title,
        parent_id=file.get("parent_id"),
        view_url=file.get("view_url"),
    )


@workflows.activity(name="hwp-download-drive-file", start_to_close_timeout=IO_TIMEOUT)
async def download_drive_file(
    file_id: str, client: ToolCallClient = Depends(drive)
) -> str:
    result = await client.call_tool(
        tool_name="download_file_content",
        arguments={"file_id": file_id},
    )
    value = _tool_result_value(result)
    content = value.get("content") if isinstance(value, dict) else value
    if not content:
        raise workflows.WorkflowError("파일 내용을 다운로드하지 못했습니다")
    return content


@workflows.activity(
    name="hwp-upload-drive-file",
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
        "content_mime_type": DOCX_MIME_TYPE,
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


@workflows.activity(
    name="hwp-convert-and-extract", start_to_close_timeout=CONVERT_TIMEOUT
)
async def convert_and_extract(content_b64: str, file_title: str) -> ExtractedDocument:
    source_bytes = base64.b64decode(content_b64)
    ext = file_title.rsplit(".", 1)[-1].lower() if "." in file_title else ""
    async with _SOFFICE_LOCK:
        docx_bytes = await asyncio.to_thread(convert_hancom_to_docx, source_bytes, ext)
    return extract_document(docx_bytes, source_format=ext)


class _EditBatch(pydantic.BaseModel):
    edits: list[TextEdit]


BODY_PROOFREAD_PROMPT = """아래는 번역서 본문 문단 목록입니다.

규칙:
1. 오탈자(맞춤법, 띄어쓰기, 오타)를 수정하십시오.
2. 문법에 맞게 윤문하되, 번역된 내용(의미)은 절대 바꾸지 마십시오. 의미를 유지하는 표현·문법·띄어쓰기·맞춤법만 손질하십시오.
3. 변경이 필요한 문단만 edits에 포함하십시오. 변경이 없으면 빈 목록을 반환하십시오.
4. revised는 해당 문단의 전체 최종 텍스트여야 합니다.

각 항목은 item_id(대괄호 안 문단 번호), original(원문 전체), revised(수정문 전체), category(typo|grammar|spacing|style 중 하나), reason(변경 이유 한 문장)을 포함합니다.

문단 목록:
"""


@workflows.activity(name="hwp-proofread-body", start_to_close_timeout=LLM_TIMEOUT)
async def proofread_body(body: list[BodyBlock]) -> list[TextEdit]:
    edits: list[TextEdit] = []
    for chunk in _chunk(body, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        result = await _parse(_EditBatch, BODY_PROOFREAD_PROMPT + listing)
        edits.extend(result.edits)
    return edits


FOOTNOTE_POLISH_PROMPT = """아래는 번역서의 주석 목록입니다.

규칙:
1. 오탈자(맞춤법, 띄어쓰기, 오타)를 수정하십시오.
2. 문법에 맞게 윤문하되, 번역된 내용(의미)은 절대 바꾸지 마십시오.
3. 책·논문·신문 등 서지명은 〈 〉로 감싸십시오(예: 〈조선왕조실록〉). 〈 〉나 다른 괄호로 표기된 서지명도 〈 〉로 통일하고, category를 book_title로 지정하십시오. 서지명이 아닌 인용구에는 적용하지 마십시오.
4. 변경이 필요한 주석만 edits에 포함하십시오. 변경이 없으면 빈 목록을 반환하십시오.
5. revised는 해당 주석의 전체 최종 텍스트여야 합니다.

각 항목은 item_id(대괄호 안 주석 번호), original(원문 전체), revised(수정문 전체), category(typo|grammar|spacing|style|book_title 중 하나), reason(변경 이유 한 문장)을 포함합니다.

주석 목록:
"""


@workflows.activity(name="hwp-polish-footnotes", start_to_close_timeout=LLM_TIMEOUT)
async def polish_footnotes(footnotes: list[FootnoteItem]) -> list[TextEdit]:
    edits: list[TextEdit] = []
    for chunk in _chunk(footnotes, LLM_CHUNK_SIZE):
        listing = "\n".join(
            f"[{footnote.footnote_id}] {footnote.text}" for footnote in chunk
        )
        result = await _parse(_EditBatch, FOOTNOTE_POLISH_PROMPT + listing)
        edits.extend(result.edits)
    return edits


class _SourceBatch(pydantic.BaseModel):
    annotations: list[SourceAnnotation]


SOURCE_PROMPT = """아래는 번역서의 주석 목록입니다. 각 주석의 내용(근거문)을 근거로 출처(저자, 서지명, 출판사, 출판 연도, 쪽수 등)를 판단할 수 있을 때만 annotations에 포함하십시오.

- 근거가 불확실하면 source 끝에 "(추정)"을 붙이십시오.
- 근거를 전혀 찾을 수 없거나 출처가 이미 완전하게 표기된 주석은 건너뛰십시오.
- source는 한국어 서지 표기 관례를 따르십시오.
- item_id는 대괄호 안 주석 번호입니다.

주석 목록:
"""


@workflows.activity(
    name="hwp-annotate-footnote-sources", start_to_close_timeout=LLM_TIMEOUT
)
async def annotate_footnote_sources(
    footnotes: list[FootnoteItem],
) -> list[SourceAnnotation]:
    annotations: list[SourceAnnotation] = []
    for chunk in _chunk(footnotes, LLM_CHUNK_SIZE):
        listing = "\n".join(
            f"[{footnote.footnote_id}] {_truncate(footnote.text, 1200)}"
            for footnote in chunk
        )
        result = await _parse(_SourceBatch, SOURCE_PROMPT + listing)
        annotations.extend(result.annotations)
    return annotations


class _HanjaBatch(pydantic.BaseModel):
    translations: list[HanjaTranslation]


HANJA_PROMPT = """아래는 번역서의 주석 목록입니다. 한문(한자문)으로 작성된 주석만 골라 자연스러운 한국어로 번역하십시오.

- item_id는 대괄호 안 주석 번호, original은 한문 주석 전체, translation은 한국어 번역문입니다.
- 한자가 일부 포함되었어도 대부분 한국어라면 번역하지 말고 건너뛰십시오.
- 이미 한국어로만 작성된 주석은 건너뛰십시오.

주석 목록:
"""


@workflows.activity(
    name="hwp-translate-hanja-footnotes", start_to_close_timeout=LLM_TIMEOUT
)
async def translate_hanja_footnotes(
    footnotes: list[FootnoteItem],
) -> list[HanjaTranslation]:
    hanja_footnotes = [f for f in footnotes if contains_hanja(f.text)]
    translations: list[HanjaTranslation] = []
    for chunk in _chunk(hanja_footnotes, LLM_CHUNK_SIZE):
        listing = "\n".join(
            f"[{footnote.footnote_id}] {_truncate(footnote.text, 1200)}"
            for footnote in chunk
        )
        result = await _parse(_HanjaBatch, HANJA_PROMPT + listing)
        translations.extend(result.translations)
    return translations


class _DifferenceBatch(pydantic.BaseModel):
    differences: list[BodyFootnoteDifference]


COMPARE_PROMPT = """아래는 번역서의 본문 문단 목록과 주석 목록입니다. 본문 내용과 주석 내용을 비교하여 다음을 찾아 differences에 정리하십시오.

- 주석의 인용이 본문의 특정 표현과 맞지 않는 경우(인용 불일치, 대상 불일치, 수치·연도 불일치 등)
- 본문은 언급했는데 주석이 관련 내용을 담고 있지 않은 경우
- 주석과 본문의 서술이 상충하는 경우

각 항목은 body_block_id(관련 본문 문단 번호, 없으면 null), footnote_id(관련 주석 번호, 없으면 null), difference(차이점에 대한 구체적 서술)을 포함합니다. 차이가 없으면 빈 목록을 반환하십시오.

본문 문단 목록:
"""


@workflows.activity(
    name="hwp-compare-body-footnotes", start_to_close_timeout=LLM_TIMEOUT
)
async def compare_body_footnotes(
    body: list[BodyBlock], footnotes: list[FootnoteItem]
) -> list[BodyFootnoteDifference]:
    differences: list[BodyFootnoteDifference] = []
    footnote_listing = "\n".join(
        f"[{footnote.footnote_id}] {_truncate(footnote.text, 600)}"
        for footnote in footnotes
    )
    footnote_listing = _truncate(footnote_listing, COMPARE_FOOTNOTE_CHAR_LIMIT)
    for chunk in _chunk(body, COMPARE_BODY_CHUNK_SIZE):
        body_listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        prompt = COMPARE_PROMPT + body_listing + "\n\n주석 목록:\n" + footnote_listing
        result = await _parse(_DifferenceBatch, prompt)
        differences.extend(result.differences)
    return differences


@workflows.activity(
    name="hwp-build-corrected-docx", start_to_close_timeout=CONVERT_TIMEOUT
)
async def build_corrected_docx(
    docx_b64: str,
    body_edits: list[TextEdit],
    footnote_edits: list[TextEdit],
    source_annotations: list[SourceAnnotation],
    hanja_translations: list[HanjaTranslation],
) -> str:
    docx_bytes = base64.b64decode(docx_b64)
    edited = await asyncio.to_thread(
        apply_review_edits,
        docx_bytes,
        body_edits,
        footnote_edits,
        source_annotations,
        hanja_translations,
    )
    return base64.b64encode(edited).decode()


@workflows.activity(name="hwp-build-review-report", start_to_close_timeout=IO_TIMEOUT)
async def build_review_report(
    source_title: str,
    body_edits: list[TextEdit],
    footnote_edits: list[TextEdit],
    source_annotations: list[SourceAnnotation],
    hanja_translations: list[HanjaTranslation],
    differences: list[BodyFootnoteDifference],
) -> str:
    report_bytes = await asyncio.to_thread(
        build_report_docx,
        source_title,
        body_edits,
        footnote_edits,
        source_annotations,
        hanja_translations,
        differences,
    )
    return base64.b64encode(report_bytes).decode()


@workflows.workflow.define(name="hwp-document-review", on_behalf_of=True)
@uses_connectors(drive)
class HwpDocumentReviewWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: HwpReviewInput) -> HwpReviewOutput:
        source = await find_drive_file(data.file_name)
        stem = source.title.rsplit(".", 1)[0] if "." in source.title else source.title
        output_parent_id = data.output_parent_id or source.parent_id

        working_copy = await copy_drive_file(
            source.file_id, f"{stem}_검수사본", output_parent_id
        )
        content_b64 = await download_drive_file(source.file_id)
        document = await convert_and_extract(content_b64, source.title)

        body_edits = await proofread_body(document.body)
        footnote_edits = await polish_footnotes(document.footnotes)
        source_annotations = await annotate_footnote_sources(document.footnotes)
        hanja_translations = await translate_hanja_footnotes(document.footnotes)
        differences = await compare_body_footnotes(document.body, document.footnotes)

        corrected_b64 = await build_corrected_docx(
            document.docx_b64,
            body_edits,
            footnote_edits,
            source_annotations,
            hanja_translations,
        )
        report_b64 = await build_review_report(
            source.title,
            body_edits,
            footnote_edits,
            source_annotations,
            hanja_translations,
            differences,
        )

        corrected = await upload_drive_file(
            f"{stem}_교정본.docx", corrected_b64, output_parent_id
        )
        report = await upload_drive_file(
            f"{stem}_교정리포트.docx", report_b64, output_parent_id
        )

        stats = {
            "body_paragraphs": len(document.body),
            "footnotes": len(document.footnotes),
            "body_edits": len(body_edits),
            "footnote_edits": len(footnote_edits),
            "source_annotations": len(source_annotations),
            "hanja_translations": len(hanja_translations),
            "body_footnote_differences": len(differences),
        }
        summary = (
            f"본문 {len(document.body)}문단, 주석 {len(document.footnotes)}건을 검수해 "
            f"본문 {len(body_edits)}건, 주석 {len(footnote_edits)}건을 교정하고 "
            f"출처 {len(source_annotations)}건을 보완했으며, 한문 주석 "
            f"{len(hanja_translations)}건을 번역했습니다. 본문-주석 차이점 "
            f"{len(differences)}건은 리포트를 확인하십시오."
        )
        return HwpReviewOutput(
            source_title=source.title,
            working_copy_title=working_copy.title,
            corrected_file=corrected,
            report_file=report,
            stats=stats,
            summary=summary,
        )
