"""검수 리포트의 교정사항을 원본틀 DOCX에 반영해 최종본을 만든다 (LLM 0회, 비용 0).

입력: 원본틀 DOCX 파일명 + 검수 리포트(txt) 파일명.
리포트에서 '—원문: …' / '—수정: …' 쌍을 파싱해 원본틀의 해당 문단을
문자 단위 diff로 교체하고(서식 보존), 리포트의 년도별 섹션을
문서 뒤에 년도 리포트로 삽입한 최종 DOCX를 업로드한다.
"""

from __future__ import annotations

import re
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

from worker.workflows.hwp_review_models import DriveFileInfo, TextEdit
from worker.workflows.pdf_document_review import _tool_result_value
from worker.workflows.hwp_review_docx import apply_edits_by_text

drive = connector(name="google_drive_mcp")

APPLY_TIMEOUT = timedelta(minutes=10)

# 리포트의 교정 항목: —원문: … / —수정: … (사유 … pN) 형식
ORIG_RE = re.compile(r"^—원문[:：]\s*(.+)$", re.MULTILINE)
REV_RE = re.compile(r"^—수정[:：]\s*(.+)$", re.MULTILINE)
PAIR_RE = re.compile(
    r"—원문[:：]\s*(?P<orig>.+?)\n—수정[:：]\s*(?P<rev>.+?)(?:\n|$)",
    re.MULTILINE | re.DOTALL,
)


def parse_year_sections(report_text: str) -> list[dict]:
    """리포트를 년도별 섹션으로 파싱한다.

    〔(년도) 검수 리포트〕 헤더를 기준으로 분리한다.
    헤더가 없으면 빈 리스트를 반환한다.
    """
    sections: list[dict] = []
    pattern = re.compile(r"〔(.+?) 검수 리포트〕")
    current_year = None
    current_lines: list[str] = []
    for line in report_text.splitlines():
        match = pattern.search(line)
        if match:
            if current_year is not None:
                sections.append(
                    {"year": current_year, "lines": list(current_lines)}
                )
            current_year = match.group(1)
            current_lines = [line]
        elif current_year is not None:
            current_lines.append(line)
    if current_year is not None:
        sections.append({"year": current_year, "lines": list(current_lines)})
    return sections


def parse_corrections(report_text: str) -> list[dict]:
    """리포트에서 교정 쌍을 파싱한다. 발췌용 … 생략 부호는 그대로 둔다."""
    pairs: list[dict] = []
    lines = report_text.splitlines()
    current_orig = None
    for line in lines:
        stripped = line.strip()
        orig_m = re.match(r"^—원문[:：]\s*(.*)$", stripped)
        rev_m = re.match(r"^—수정[:：]\s*(.*)$", stripped)
        if orig_m:
            current_orig = orig_m.group(1).strip()
        elif rev_m and current_orig is not None:
            pairs.append(
                {"original": current_orig, "revised": rev_m.group(1).strip()}
            )
            current_orig = None
        elif not stripped and current_orig is None:
            continue
    return pairs


@workflows.activity(name="apply-review", start_to_close_timeout=APPLY_TIMEOUT)
async def apply_review(
    frame_file_id: str,
    report_file_ids: list[str],
    title: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> tuple[DriveFileInfo, int, int]:
    import asyncio
    import base64
    import io

    async def _download(file_id: str) -> bytes:
        result = await client.call_tool(
            tool_name="download_file_content", arguments={"file_id": file_id}
        )
        value = _tool_result_value(result)
        content = value.get("content") if isinstance(value, dict) else value
        if not content:
            raise workflows.WorkflowError(f"파일 다운로드 실패: {file_id}")
        return base64.b64decode(content)

    frame_bytes = await _download(frame_file_id)
    report_parts: list[str] = []
    for report_file_id in report_file_ids:
        report_bytes = await _download(report_file_id)
        report_parts.append(report_bytes.decode("utf-8", errors="replace"))
    report_text = "\n\n======================================================================\n\n".join(report_parts)
    pairs = parse_corrections(report_text)

    def _apply() -> tuple[bytes, int]:
        from docx import Document
        from docx.oxml.ns import qn
        from worker.workflows.hwp_review_docx import _apply_text_edit

        document = Document(io.BytesIO(frame_bytes))
        paragraphs = list(document.element.body.iter(qn("w:p")))
        applied = 0
        used: set[int] = set()
        for pair in pairs:
            # 발췌의 … 를 떼고 핵심 구간만 남긴다.
            orig_core = pair["original"].strip().strip("…").strip()
            rev_core = pair["revised"].strip()
            rev_core = re.sub(r"^…\s*", "", rev_core)
            rev_core = re.sub(r"\s*…$", "", rev_core).strip()
            if not orig_core or orig_core == rev_core:
                continue
            for index, paragraph in enumerate(paragraphs):
                if index in used:
                    continue
                text = "".join(t.text or "" for t in paragraph.iter(qn("w:t")))
                if orig_core not in text:
                    continue
                # 문단 안에서 발췌 구간만 교체한다.
                new_text = text.replace(orig_core, rev_core, 1)
                if new_text == text:
                    continue
                _apply_text_edit(paragraph, new_text)
                used.add(index)
                applied += 1
                break
        buffer = io.BytesIO()
        document.save(buffer)
        return buffer.getvalue(), applied

    new_bytes, applied = await asyncio.to_thread(_apply)

    def _insert_reports() -> bytes:
        # 년도별 섹션을 파싱해서 해당 년도 헤딩 직전에 삽입한다.
        from docx import Document
        from docx.oxml.ns import qn

        year_sections = parse_year_sections(report_text)
        document = Document(io.BytesIO(new_bytes))
        paragraphs = list(document.element.body.iter(qn("w:p")))
        paragraph_texts = [
            "".join(t.text or "" for t in p.iter(qn("w:t")))
            for p in paragraphs
        ]

        placed_count = 0
        placed_indices: set[int] = set()
        for section in year_sections:
            year_norm = re.sub(r"\s+", "", section["year"])
            # 이 년도의 헤딩(예: '태조 19년')이 문서에 있는지 찾는다
            # 헤딩은 본문 중간에 있고, 그 직전이 삽입 위치다.
            # 다음 년도 헤딩을 찾아 그 직전에 삽입한다.
            pass

        # 각 년도 섹션을 다음 년도 헤딩 직전에 삽입한다.
        # 마지막 년도는 문서 끝에 삽입한다.
        for section in year_sections:
            year_core = section["year"].split("(")[0].strip()
            year_norm = re.sub(r"\s+", "", year_core)
            # 이 년도 다음에 오는 년도 헤딩을 찾는다 (년도 목록에서)
            next_year = None
            for other in year_sections:
                if other is section:
                    continue
                other_core = other["year"].split("(")[0].strip()
                if other_core > year_core:
                    if next_year is None or other_core < next_year:
                        next_year = other_core
            if next_year:
                # 다음 년도 헤딩 위치를 찾는다
                found = False
                for index, text in enumerate(paragraph_texts):
                    if index in placed_indices:
                        continue
                    text_norm = re.sub(r"\s+", "", text)
                    if next_year.replace(" ", "") in text_norm:
                        anchor = paragraphs[index]
                        for line in section["lines"]:
                            para = document.add_paragraph(line)
                            anchor.addprevious(para._p)
                        placed_indices.add(index)
                        placed_count += 1
                        found = True
                        break
                if not found:
                    # 못 찾으면 문서 끝에
                    body = document.element.body
                    anchor_el = body.find(qn("w:sectPr"))
                    for line in section["lines"]:
                        para = document.add_paragraph(line)
                        if anchor_el is not None:
                            anchor_el.addprevious(para._p)
                        else:
                            body.append(para._p)
                    placed_count += 1
            else:
                # 마지막 년도는 문서 끝에
                body = document.element.body
                anchor_el = body.find(qn("w:sectPr"))
                for line in section["lines"]:
                    para = document.add_paragraph(line)
                    if anchor_el is not None:
                        anchor_el.addprevious(para._p)
                    else:
                        body.append(para._p)
                placed_count += 1

        buffer = io.BytesIO()
        document.save(buffer)
        return buffer.getvalue()

    final_bytes = await asyncio.to_thread(_insert_reports)

    final_bytes = await _upload_tail(new_bytes)

    arguments: dict[str, Any] = {
        "title": title,
        "base64_content": base64.b64encode(final_bytes).decode("ascii"),
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
        raise workflows.WorkflowError("최종본 업로드에 실패했습니다")
    return (
        DriveFileInfo(
            file_id=file.get("id") or "",
            title=file.get("title") or title,
            parent_id=file.get("parent_id"),
            view_url=file.get("view_url"),
        ),
        applied,
        len(pairs),
    )


class ApplyReviewInput(pydantic.BaseModel):
    frame_file_id: str  # 원본틀 DOCX 파일 ID
    report_file_ids: list[str] = []  # 겹치는 검수 리포트 txt 파일 ID 목록
    output_title: str
    output_parent_id: str | None = None


class ApplyReviewOutput(pydantic.BaseModel):
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str

@workflows.workflow.define(name="apply-review", on_behalf_of=True)
@uses_connectors(drive)
class ApplyReviewWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: ApplyReviewInput) -> ApplyReviewOutput:
        parent = data.output_parent_id or "0AOuygiQzrnEVUk9PVA"
        result_file, applied, total = await apply_review(
            data.frame_file_id, data.report_file_ids, data.output_title, parent
        )
        return ApplyReviewOutput(
            result_file=result_file,
            stats={"corrections_parsed": total, "corrections_applied": applied},
            summary=(
                f"교정 {total}건 중 {applied}건을 원본틀에 반영하고 "
                f"년도 리포트를 삽입했습니다. LLM 호출은 없었습니다."
            ),
        )


