"""틀 DOCX에 검수 리포트를 병합한다 (LLM 호출 없음, 비용 0).

원본 양식 보존 틀(segment-convert 결과)의 맨 뒤에 검수 리포트 문단을
추가한 결과물 DOCX를 만든다. 리포트 내용은 입력으로 받은 줄 목록
그대로 삽입하며, 어떤 판단도 하지 않는다(내용 판단은 리포트 작성자의 몫).
"""

from __future__ import annotations

from datetime import timedelta
from typing import Any

import mistralai.workflows as workflows
import pydantic
from mistralai.workflows import Depends
from mistralai.workflows.plugins.mistralai.connectors import (
    ToolCallClient,
    connector,
    uses_connectors,
)

from worker.workflows.hwp_review_models import DriveFileInfo
from worker.workflows.pdf_document_review import _tool_result_value

drive = connector(name="google_drive_mcp")

PATCH_TIMEOUT = timedelta(minutes=5)


class DocxAppendInput(pydantic.BaseModel):
    file_name: str  # 틀 파일명 (예: ..._p8-15_원본틀.docx)
    lines: list[str]
    output_title: str | None = None
    output_parent_id: str | None = None


class DocxAppendOutput(pydantic.BaseModel):
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


@workflows.activity(name="docx-find-frame", start_to_close_timeout=PATCH_TIMEOUT)
async def find_frame_file(
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
        if (file.get("title") or "").lower().endswith(".docx")
        and file_name.lower() in (file.get("title") or "").lower()
    ]
    if not candidates:
        raise workflows.WorkflowError(
            f"드라이브에서 틀 문서를 찾지 못했습니다: {file_name}",
            non_retryable=True,
        )

    def _created(file):
        return str(file.get("created_time") or file.get("modified_time") or "")

    candidates.sort(key=_created, reverse=True)
    file = candidates[0]
    return DriveFileInfo(
        file_id=file.get("id") or "",
        title=file.get("title") or "",
        parent_id=file.get("parent_id"),
        mime_type=file.get("mime_type"),
        view_url=file.get("view_url"),
        file_size=int(file.get("file_size") or 0) or None,
    )


@workflows.activity(name="docx-append-report", start_to_close_timeout=PATCH_TIMEOUT)
async def append_report(
    file_id: str,
    lines: list[str],
    title: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> tuple[DriveFileInfo, int]:
    import asyncio
    import base64
    import io

    from docx import Document
    from docx.oxml.ns import qn

    result = await client.call_tool(
        tool_name="download_file_content", arguments={"file_id": file_id}
    )
    value = _tool_result_value(result)
    content = value.get("content") if isinstance(value, dict) else value
    if not content:
        raise workflows.WorkflowError("틀 DOCX 다운로드에 실패했습니다")
    docx_bytes = base64.b64decode(content)

    def _append() -> bytes:
        document = Document(io.BytesIO(docx_bytes))
        body = document.element.body
        anchor = body.find(qn("w:sectPr"))
        inserted = 0
        if anchor is None:
            anchor = body[-1]
        for line in lines:
            paragraph = document.add_paragraph(line)
            if anchor is not None and anchor.tag == qn("w:sectPr"):
                anchor.addprevious(paragraph._p)
            else:
                body.append(paragraph._p)
            inserted += 1
        buffer = io.BytesIO()
        document.save(buffer)
        return buffer.getvalue()

    new_bytes = await asyncio.to_thread(_append)

    arguments: dict[str, Any] = {
        "title": title,
        "base64_content": base64.b64encode(new_bytes).decode("ascii"),
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
    return (
        DriveFileInfo(
            file_id=file.get("id") or "",
            title=file.get("title") or title,
            parent_id=file.get("parent_id"),
            view_url=file.get("view_url"),
        ),
        inserted,
    )


@workflows.workflow.define(name="docx-append-report", on_behalf_of=True)
@uses_connectors(drive)
class DocxAppendReportWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: DocxAppendInput) -> DocxAppendOutput:
        frame = await find_frame_file(data.file_name)
        parent = data.output_parent_id or frame.parent_id
        title = data.output_title or frame.title.replace(
            "_원본틀.docx", "_검수결과.docx"
        )
        result_file, inserted = await append_report(
            frame.file_id, data.lines, title, parent
        )
        return DocxAppendOutput(
            result_file=result_file,
            stats={"inserted_lines": inserted},
            summary=(
                f"틀에 검수 리포트 {inserted}줄을 병합했습니다. "
                f"LLM 호출은 없었습니다."
            ),
        )
