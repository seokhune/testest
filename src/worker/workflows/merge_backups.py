"""검수 결과 백업 파일들을 순서대로 하나로 병합한다 (LLM 호출 없음, 비용 0).

입력으로 받은 순서대로 드라이브의 텍스트 백업 파일을 내려받아 이어붙이고,
하나의 통합 결과 파일로 업로드한다. 어떤 판단도 하지 않는다.
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

MERGE_TIMEOUT = timedelta(minutes=10)


class MergeBackupsInput(pydantic.BaseModel):
    file_ids: list[str]  # 병합 순서대로
    labels: list[str] | None = None  # 각 파일 앞에 붙일 구간 라벨 (선택)
    output_title: str
    output_parent_id: str | None = None


class MergeBackupsOutput(pydantic.BaseModel):
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


@workflows.activity(name="merge-backups", start_to_close_timeout=MERGE_TIMEOUT)
async def merge_backups(
    file_ids: list[str],
    labels: list[str] | None,
    title: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> tuple[DriveFileInfo, int]:
    import base64

    async def _merge() -> tuple[bytes, int]:
        parts: list[str] = []
        for index, file_id in enumerate(file_ids):
            result = await client.call_tool(
                tool_name="download_file_content",
                arguments={"file_id": file_id},
            )
            value = _tool_result_value(result)
            content = value.get("content") if isinstance(value, dict) else value
            if not content:
                raise workflows.WorkflowError(
                    f"{index + 1}번째 백업 파일을 내려받지 못했습니다: {file_id}"
                )
            text = base64.b64decode(content).decode("utf-8", errors="replace")
            label = labels[index] if labels and index < len(labels) else None
            header = f"\n\n{label}\n{'=' * 60}\n" if label else ""
            parts.append(header + text.strip())
        merged = ("\n\n" + "=" * 70 + "\n\n").join(parts)
        return merged.encode("utf-8"), len(file_ids)

    merged_bytes, count = await _merge()

    arguments: dict[str, Any] = {
        "title": title,
        "base64_content": base64.b64encode(merged_bytes).decode("ascii"),
        "content_mime_type": "text/plain",
        "disable_conversion_to_google_type": True,
    }
    if parent_id:
        arguments["parent_id"] = parent_id
    upload = await client.call_tool(tool_name="create_file", arguments=arguments)
    file = _tool_result_value(upload) or {}
    if not file.get("id"):
        raise workflows.WorkflowError("병합 결과 업로드에 실패했습니다")
    return (
        DriveFileInfo(
            file_id=file.get("id") or "",
            title=file.get("title") or title,
            parent_id=file.get("parent_id"),
            view_url=file.get("view_url"),
        ),
        count,
    )


@workflows.workflow.define(name="merge-backups", on_behalf_of=True)
@uses_connectors(drive)
class MergeBackupsWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: MergeBackupsInput) -> MergeBackupsOutput:
        if not data.file_ids:
            raise workflows.WorkflowError("병합할 파일이 없습니다")
        parent = data.output_parent_id or "0AOuygiQzrnEVUk9PVA"
        result_file, count = await merge_backups(
            data.file_ids, data.labels, data.output_title, parent
        )
        return MergeBackupsOutput(
            result_file=result_file,
            stats={"files_merged": count},
            summary=f"백업 {count}개를 순서대로 병합했습니다. LLM 호출은 없었습니다.",
        )
