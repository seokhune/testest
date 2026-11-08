"""구간별 원본 양식 DOCX 변환 워크플로우 (LLM 호출 없음, 비용 0).

지정한 페이지 구간을 pdf2docx로 변환해 원본 양식(글꼴·굵기·각주)이
보존된 DOCX만 만들어 업로드한다. 검수 내용(교정·번역·차이·출처)은
이 틀에 들어가지 않는다 — 그 내용은 Work의 검수 스킬이 수행하고,
결과는 별도의 패치 워크플로우로 이 틀에 병합된다.
"""

from __future__ import annotations

from datetime import timedelta

import mistralai.workflows as workflows
import pydantic
from mistralai.workflows import Depends
from mistralai.workflows.plugins.mistralai.connectors import (
    ToolCallClient,
    connector,
    uses_connectors,
)

from worker.workflows.hwp_review_models import DriveFileInfo
from worker.workflows.pdf_document_review import (
    _tool_result_value,
    find_pdf_file,
)
from worker.workflows.hwp_review_docx import (
    convert_pdf_to_docx,
    unwrap_text_frames,
)

drive = connector(name="google_drive_mcp")

CONVERT_TIMEOUT = timedelta(minutes=10)


class SegmentConvertInput(pydantic.BaseModel):
    file_name: str
    page_start: int
    page_end: int
    output_parent_id: str | None = None


class SegmentConvertOutput(pydantic.BaseModel):
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


@workflows.activity(name="segment-convert", start_to_close_timeout=CONVERT_TIMEOUT)
async def convert_segment(
    file_id: str,
    first_page: int,
    last_page: int,
    title: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> DriveFileInfo:
    import asyncio
    import base64

    result = await client.call_tool(
        tool_name="download_file_content", arguments={"file_id": file_id}
    )
    value = _tool_result_value(result)
    content = value.get("content") if isinstance(value, dict) else value
    if not content:
        raise workflows.WorkflowError("PDF 다운로드에 실패했습니다")
    pdf_bytes = base64.b64decode(content)

    def _convert() -> bytes:
        docx_bytes = convert_pdf_to_docx(pdf_bytes, first_page - 1, last_page)
        return unwrap_text_frames(docx_bytes)

    docx_bytes = await asyncio.to_thread(_convert)

    arguments = {
        "title": title,
        "base64_content": base64.b64encode(docx_bytes).decode("ascii"),
        "content_mime_type": (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        "disable_conversion_to_google_type": True,
    }
    if parent_id:
        arguments["parent_id"] = parent_id
    upload = await client.call_tool(tool_name="create_file", arguments=arguments)
    file = _tool_result_value(upload) or {}
    if not file.get("id"):
        raise workflows.WorkflowError("드라이브 업로드에 실패했습니다")
    return DriveFileInfo(
        file_id=file.get("id") or "",
        title=file.get("title") or title,
        parent_id=file.get("parent_id"),
        view_url=file.get("view_url"),
    )


@workflows.workflow.define(name="segment-convert", on_behalf_of=True)
@uses_connectors(drive)
class SegmentConvertWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: SegmentConvertInput) -> SegmentConvertOutput:
        source = await find_pdf_file(data.file_name)
        parent = data.output_parent_id or source.parent_id
        stem = source.title.rsplit(".", 1)[0] if "." in source.title else source.title
        title = f"{stem}_p{data.page_start}-{data.page_end}_원본틀.docx"
        result_file = await convert_segment(
            source.file_id,
            data.page_start,
            data.page_end,
            title,
            parent,
        )
        return SegmentConvertOutput(
            result_file=result_file,
            stats={"pages": data.page_end - data.page_start + 1},
            summary=(
                f"p{data.page_start}~{data.page_end} 원본 양식 DOCX 변환 완료. "
                f"LLM 호출은 없었습니다."
            ),
        )
