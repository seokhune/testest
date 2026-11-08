"""검수결과 DOCX의 3번 섹션(본문-주석 내용 차이)에 매핑 검수 결과를 병합한다.

LLM 호출이 없으므로 실행 비용이 0이다. 기존 검수결과 DOCX에서 3번 섹션의
안내 문구(placeholder)를 찾아, 매핑 검수로 확정된 차이 내용으로 교체한 뒤
새 파일로 업로드한다. 원본 서식은 건드리지 않고 해당 문단만 교체한다.
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
from worker.workflows.pdf_document_review import (
    _tool_result_value,
    find_pdf_file,
)

drive = connector(name="google_drive_mcp")

PATCH_TIMEOUT = timedelta(minutes=5)

SECTION3_HEADING = "3) 본문-주석 내용 차이"
SECTION4_HEADING = "4) 주석 출처 보완 제안"


class DocxPatchInput(pydantic.BaseModel):
    file_name: str
    lines: list[str]
    output_parent_id: str | None = None


class DocxPatchOutput(pydantic.BaseModel):
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


@workflows.activity(name="docx-patch-section3", start_to_close_timeout=PATCH_TIMEOUT)
async def patch_section3(
    file_id: str,
    title: str,
    lines: list[str],
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> tuple[DriveFileInfo, int, int]:
    """DOCX를 받아 3번 섹션을 교체한 새 DOCX를 업로드한다.

    반환값: (업로드된 파일, 교체한 문단 수, 삽입한 줄 수)
    """
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
        raise workflows.WorkflowError("DOCX 다운로드 결과가 비었습니다")
    docx_bytes = base64.b64decode(content)

    def _patch() -> tuple[bytes, int, int]:
        document = Document(io.BytesIO(docx_bytes))
        paragraphs = list(document.element.body.iter(qn("w:p")))
        heading_index = None
        next_index = None
        for index, paragraph in enumerate(paragraphs):
            text = "".join(t.text or "" for t in paragraph.iter(qn("w:t")))
            stripped = text.strip()
            if heading_index is None and stripped.startswith(SECTION3_HEADING):
                heading_index = index
                continue
            if heading_index is not None and stripped.startswith(SECTION4_HEADING):
                next_index = index
                break
        if heading_index is None or next_index is None:
            raise workflows.WorkflowError(
                "3번/4번 섹션 경계를 찾지 못했습니다"
            )
        # 3번 헤딩과 4번 헤딩 사이의 모든 문단(구 방식의 차이 항목 포함)을 지운다
        removed = 0
        for paragraph in paragraphs[heading_index + 1 : next_index]:
            paragraph.getparent().remove(paragraph)
            removed += 1
        # 3번 헤딩 바로 뒤에 새 내용을 삽입한다
        anchor = paragraphs[heading_index]
        inserted = 0
        for line in lines:
            new_paragraph = document.add_paragraph(line)
            anchor.addnext(new_paragraph._p)
            anchor = new_paragraph._p
            inserted += 1
        buffer = io.BytesIO()
        document.save(buffer)
        return buffer.getvalue(), removed, inserted

    new_bytes, replaced, inserted = await asyncio.to_thread(_patch)

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
        replaced,
        inserted,
    )


@workflows.workflow.define(name="docx-patch-section3", on_behalf_of=True)
@uses_connectors(drive)
class DocxPatchSection3Workflow:
    @workflows.workflow.entrypoint
    async def run(self, data: DocxPatchInput) -> DocxPatchOutput:
        source = await find_pdf_file(data.file_name)
        stem = source.title.rsplit(".", 1)[0] if "." in source.title else source.title
        parent = data.output_parent_id or source.parent_id
        result_file, replaced, inserted = await patch_section3(
            source.file_id,
            f"{stem}_검수결과_최종.docx",
            data.lines,
            parent,
        )
        return DocxPatchOutput(
            result_file=result_file,
            stats={"replaced": replaced, "inserted": inserted},
            summary=(
                f"3번 섹션 안내 문구 {replaced}개 문단을 매핑 검수 결과 "
                f"{inserted}줄로 교체했습니다. LLM 호출은 없었습니다."
            ),
        )
