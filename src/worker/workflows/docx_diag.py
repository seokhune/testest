"""DOCX 변환 텍스트 진단 워크플로우 (LLM 호출 없음, 비용 0).

원본 PDF를 pdf2docx로 변환한 뒤, 지정한 텍스트(needle)가 변환 결과에
존재하는지 확인한다. 없으면 어느 지점에서 어긋나는지(가장 긴 일치
접두사의 끝 위치와 그 주변의 변환 결과 텍스트)를 보고한다.
검증기 오탐의 원인을 추측 없이 확인하기 위한 도구다.
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
    load_pages,
)
from worker.workflows.hwp_review_docx import (
    convert_pdf_to_docx,
    unwrap_text_frames,
)
import re
import unicodedata

drive = connector(name="google_drive_mcp")

DIAG_TIMEOUT = timedelta(minutes=10)

_HANJA_RE = re.compile(r"[\u3000-\u9fff\u3400-\u4dbf\uf000-\ufaff\ufe30-\ufe4f]")


def _norm(text: str) -> str:
    return unicodedata.normalize("NFC", "".join((text or "").split()))


def _korean_only(text: str) -> str:
    return _norm(_HANJA_RE.sub("", text or ""))


class DocxDiagInput(pydantic.BaseModel):
    file_name: str
    needle: str
    output_parent_id: str | None = None
    page_start: int | None = None
    page_end: int | None = None


class DocxDiagOutput(pydantic.BaseModel):
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


@workflows.activity(name="docx-diag", start_to_close_timeout=DIAG_TIMEOUT)
async def diagnose(
    file_id: str,
    first_page: int,
    last_page: int,
    needle: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> tuple[DriveFileInfo, str]:
    import asyncio
    import base64
    import io

    result = await client.call_tool(
        tool_name="download_file_content", arguments={"file_id": file_id}
    )
    value = _tool_result_value(result)
    content = value.get("content") if isinstance(value, dict) else value
    if not content:
        raise workflows.WorkflowError("PDF 다운로드에 실패했습니다")
    pdf_bytes = base64.b64decode(content)

    def _run() -> tuple[str, str]:
        docx_bytes = convert_pdf_to_docx(pdf_bytes, first_page - 1, last_page)
        docx_bytes = unwrap_text_frames(docx_bytes)
        from docx import Document
        from docx.oxml.ns import qn

        document = Document(io.BytesIO(docx_bytes))
        paragraphs = [
            "".join(t.text or "" for t in p.iter(qn("w:t")))
            for p in document.element.body.iter(qn("w:p"))
        ]
        whole = "\n".join(paragraphs)
        whole_norm = _norm(whole)
        whole_korean = _korean_only(whole)

        needle_norm = _norm(needle)
        needle_korean = _korean_only(needle)
        lines = []
        if needle_norm and needle_norm in whole_norm:
            lines.append("결과: 전체 텍스트가 변환 결과에 존재합니다 (정규 대조 통과).")
        elif needle_korean and needle_korean in whole_korean:
            lines.append("결과: 한글만 대조 시 존재합니다 (한자 인코딩 차이).")
        else:
            lines.append("결과: 존재하지 않습니다. 어긋난 지점을 찾습니다.")
            # 가장 긴 일치 접두사를 이분 탐색으로 찾는다
            probe_korean = needle_korean
            low, high = 0, len(probe_korean)
            while low < high:
                mid = (low + high + 1) // 2
                if probe_korean[:mid] in whole_korean:
                    low = mid
                else:
                    high = mid - 1
            matched_len = low
            lines.append(f"가장 긴 일치 접두사 길이(한글 기준): {matched_len} / {len(probe_korean)}")
            context_start = whole_korean.find(probe_korean[:matched_len]) if matched_len else -1
            if matched_len and context_start >= 0:
                ctx = whole_korean[max(0, context_start + matched_len - 60): context_start + matched_len + 120]
                lines.append(f"어긋난 지점 주변(변환 결과): …{ctx}…")
                tail = probe_korean[matched_len: matched_len + 120]
                lines.append(f"이 시점부터 찾아야 할 텍스트(원문): …{tail}…")
            else:
                lines.append("접두사가 아예 일치하지 않습니다(첫 글자부터 다름).")
                lines.append(f"원문 첫 80자(한글): {probe_korean[:80]}")
        return "\n".join(lines), whole

    report, _whole = await asyncio.to_thread(_run)

    arguments = {
        "title": "docx_진단결과.txt",
        "base64_content": base64.b64encode(report.encode("utf-8")).decode("ascii"),
        "content_mime_type": "text/plain",
        "disable_conversion_to_google_type": True,
    }
    if parent_id:
        arguments["parent_id"] = parent_id
    upload = await client.call_tool(tool_name="create_file", arguments=arguments)
    file = _tool_result_value(upload) or {}
    if not file.get("id"):
        raise workflows.WorkflowError("진단 결과 업로드에 실패했습니다")
    return (
        DriveFileInfo(
            file_id=file.get("id") or "",
            title="docx_진단결과.txt",
            parent_id=file.get("parent_id"),
            view_url=file.get("view_url"),
        ),
        report,
    )


@workflows.workflow.define(name="docx-diag", on_behalf_of=True)
@uses_connectors(drive)
class DocxDiagWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: DocxDiagInput) -> DocxDiagOutput:
        source = await find_pdf_file(data.file_name)
        parent = data.output_parent_id or source.parent_id
        first = data.page_start or 1
        last = data.page_end or 999
        result_file, report = await diagnose(
            source.file_id, first, last, data.needle, parent
        )
        return DocxDiagOutput(
            result_file=result_file,
            stats={"report_chars": len(report)},
            summary=f"변환 결과 대조 진단을 완료했습니다. LLM 호출은 없었습니다.",
        )
