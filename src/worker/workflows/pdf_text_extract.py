"""PDF 페이지 텍스트 추출 전용 워크플로우.

LLM 호출이 없으므로 실행 비용이 0이다. 지정한 페이지 범위의 원문 텍스트를
그대로 Drive의 텍스트 파일로 저장한다. 검수 리포트를 사람(또는 에이전트)이
직접 읽고 수행하는 작업의 원재료로 쓰기 위한 것이다.
"""

from __future__ import annotations

from datetime import timedelta
from typing import Any

import mistralai.workflows as workflows
from mistralai.workflows import Depends
from mistralai.workflows.plugins.mistralai.connectors import (
    ToolCallClient,
    connector,
    uses_connectors,
)

from worker.workflows.hwp_review_models import DriveFileInfo
from worker.workflows.pdf_document_review import (
    find_pdf_file,
    load_pages,
    _tool_result_value,
)

import pydantic

drive = connector(name="google_drive_mcp")

EXTRACT_TIMEOUT = timedelta(minutes=10)


class PdfTextExtractInput(pydantic.BaseModel):
    file_name: str
    output_parent_id: str | None = None
    page_start: int | None = None
    page_end: int | None = None


class PdfTextExtractOutput(pydantic.BaseModel):
    source_title: str
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


@workflows.activity(name="pdf-upload-text-file", start_to_close_timeout=EXTRACT_TIMEOUT)
async def upload_text_file(
    title: str,
    text: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> DriveFileInfo:
    import base64

    arguments: dict[str, Any] = {
        "title": title,
        "base64_content": base64.b64encode(text.encode("utf-8")).decode("ascii"),
        "content_mime_type": "text/plain",
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


@workflows.workflow.define(name="pdf-text-extract", on_behalf_of=True)
@uses_connectors(drive)
class PdfTextExtractWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: PdfTextExtractInput) -> PdfTextExtractOutput:
        source = await find_pdf_file(data.file_name)
        stem = source.title.rsplit(".", 1)[0] if "." in source.title else source.title
        output_parent_id = data.output_parent_id or source.parent_id

        pages, total_pages = await load_pages(
            source.file_id, data.page_start, data.page_end
        )

        lines: list[str] = []
        for number, text in pages:
            lines.append(f"===== 물리 페이지 {number} =====")
            lines.append(text.strip())
            lines.append("")
        extracted = "\n".join(lines)

        start = data.page_start or (pages[0][0] if pages else 0)
        end = data.page_end or (pages[-1][0] if pages else 0)
        result_file = await upload_text_file(
            f"{stem}_p{start}-{end}_원문텍스트.txt", extracted, output_parent_id
        )

        return PdfTextExtractOutput(
            source_title=source.title,
            result_file=result_file,
            stats={
                "pages_extracted": len(pages),
                "total_pages": total_pages,
                "chars": len(extracted),
            },
            summary=(
                f"{len(pages)}개 페이지({start}~{end})의 원문 텍스트를 추출해 "
                f"드라이브에 저장했습니다. LLM 호출은 없었습니다."
            ),
        )
