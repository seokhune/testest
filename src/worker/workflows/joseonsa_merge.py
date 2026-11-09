"""조선사 검수결과 28개 단일 DOCX 병합 워크플로우 (LLM 0회, 비용 0).

입력 file_ids 28개를 페이지 순서대로 받는다(첫 번째는 p1-24 TXT,
두 번째부터는 구간 검수결과 DOCX 27개). p1-24 TXT를 p25-48 DOCX
템플릿 서식(HCRBatang 11pt, 왼쪽 정렬, '====='/'〔' 시작 줄 굵게,
같은 페이지 여백)으로 변환한 뒤, 28개를 페이지 순서로 하나의 DOCX로
병합해 드라이브에 업로드한다.

병합은 docxcompose와 동등한 순수 표준 라이브러리 구현이다
(스타일·번호·이미지·섹션 속성 보존, rId 재매핑). 검증(문단/문자 수
비율, 구간 첫/끝 문단, 핵심 표지, 첫 문단)을 통과해야만 업로드한다.
검증 실패 시 업로드하지 않는다. 내용은 절대 수정하지 않는다.
"""

from __future__ import annotations

import asyncio
import base64
import io
import re
import zipfile
from datetime import timedelta
from typing import Any
from xml.dom import minidom

import pydantic

import mistralai.workflows as workflows
from mistralai.workflows import Depends
from mistralai.workflows.plugins.mistralai.connectors import (
    ToolCallClient,
    connector,
    uses_connectors,
)

drive = connector(name="google_drive_mcp")

MERGE_TIMEOUT = timedelta(minutes=20)
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DEFAULT_PARENT = "0AOuygiQzrnEVUk9PVA"
FIRST_PARA = "조선사 제3편 제1권 원고 검수 결과물"
MARKERS = ("114)", "謹按本國世代編年", "豈不謬乎")
P1_LINE_COUNT = 295

CT_MAP = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "gif": "image/gif", "bmp": "image/bmp", "tiff": "image/tiff",
    "tif": "image/tiff", "emf": "image/x-emf", "wmf": "image/x-wmf",
    "svg": "image/svg+xml",
}


class DriveFileInfo(pydantic.BaseModel):
    file_id: str
    title: str
    parent_id: str | None = None
    mime_type: str | None = None
    view_url: str | None = None
    file_size: int | None = None


class JoseonsaMergeInput(pydantic.BaseModel):
    # 28개 파일 ID, 페이지 순서. [0]=p1-24 TXT, [1]=p25-48 DOCX(변환 템플릿), ... [27]=p649-664
    file_ids: list[str]
    output_title: str
    output_parent_id: str | None = None


class JoseonsaMergeOutput(pydantic.BaseModel):
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


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


def _esc(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _body_split(docxml: str):
    """document.xml -> (prefix, 본문(마지막 sectPr 제외), 마지막 sectPr, suffix)."""
    match = re.search(r"<w:body>(.*)</w:body>", docxml, re.S)
    if not match:
        raise ValueError("document.xml에 body가 없습니다")
    prefix, suffix = docxml[: match.start(1)], docxml[match.end(1):]
    body = match.group(1)
    sects = list(re.finditer(r"<w:sectPr(?:\s[^>]*)?>.*?</w:sectPr>", body, re.S))
    if not sects:
        raise ValueError("document.xml에 sectPr가 없습니다")
    last = sects[-1]
    tail = body[last.end():].strip()
    if tail:
        raise ValueError("마지막 sectPr 뒤에 예외 데이터가 있습니다")
    return prefix, body[: last.start()], last.group(0), suffix


def _read_rels(rels_str: str) -> dict:
    out = {}
    for match in re.finditer(r"<Relationship\s[^>]*>", rels_str):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', match.group(0)))
        rid = attrs.get("Id")
        if rid:
            out[rid] = {
                "type": attrs.get("Type", ""),
                "target": attrs.get("Target", ""),
                "mode": attrs.get("TargetMode", ""),
            }
    return out


def _zip_bytes(parts: dict, extra: dict) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
        for name, data in extra.items():
            zf.writestr(name, data)
    return buffer.getvalue()


def build_p1_24_docx(txt_bytes: bytes, template_bytes: bytes) -> bytes:
    """p1-24 TXT(295줄)를 p25-48 DOCX 템플릿 위에 변환한다.

    서식: HCRBatang 11pt, 왼쪽 정렬, '====='·'〔' 시작 줄 굵게.
    페이지 여백·스타일은 템플릿(기존 구간)과 같다.
    """
    template = zipfile.ZipFile(io.BytesIO(template_bytes))
    parts = {name: template.read(name) for name in template.namelist()}
    docxml = parts["word/document.xml"].decode("utf-8")
    prefix, _content, final_sect, suffix = _body_split(docxml)

    lines = txt_bytes.decode("utf-8").splitlines()
    if len(lines) != P1_LINE_COUNT:
        raise ValueError(
            f"p1-24 TXT 줄 수가 {len(lines)}입니다({P1_LINE_COUNT}이어야 함)"
        )

    fonts = (
        '<w:rFonts w:ascii="HCRBatang" w:hAnsi="HCRBatang" '
        'w:eastAsia="HCRBatang" w:cs="HCRBatang"/>'
    )
    paragraphs = []
    for line in lines:
        bold = line.startswith("=====") or line.startswith("〔")
        bold_xml = "<w:b/><w:bCs/>" if bold else ""
        run_props = (
            f'<w:rPr>{fonts}{bold_xml}<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
        )
        if line == "":
            paragraph = f'<w:p><w:pPr><w:jc w:val="left"/>{run_props}</w:pPr></w:p>'
        else:
            paragraph = (
                f'<w:p><w:pPr><w:jc w:val="left"/>{run_props}</w:pPr>'
                f'<w:r>{run_props}<w:t xml:space="preserve">{_esc(line)}</w:t></w:r></w:p>'
            )
        paragraphs.append(paragraph)

    parts["word/document.xml"] = (
        prefix + "".join(paragraphs) + final_sect + suffix
    ).encode("utf-8")

    # 본문이 참조하지 않는 템플릿 이미지 rel/part 제거
    rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
    rel_map = _read_rels(rels)
    for rid, rel in rel_map.items():
        if rel["type"].endswith("/image"):
            rels = re.sub(
                r'<Relationship\s+Id="%s"[^>]*/>' % re.escape(rid), "", rels
            )
    parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    for name in [n for n in parts if n.startswith("word/media/")]:
        del parts[name]
    return _zip_bytes(parts, {})


def merge_docx(docx_blobs: list[bytes]) -> bytes:
    """28개 DOCX를 순서대로 하나로 병합한다 (docxcompose 동등 구현)."""
    first = zipfile.ZipFile(io.BytesIO(docx_blobs[0]))
    parts = {name: first.read(name) for name in first.namelist()}
    docxml = parts["word/document.xml"].decode("utf-8")
    prefix, master_content, master_sect, suffix = _body_split(docxml)
    rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
    ctypes = parts["[Content_Types].xml"].decode("utf-8")

    existing_ids = set(_read_rels(rels).keys())
    rel_counter = [0]
    new_rels: list[str] = []
    new_parts: dict = {}
    used_exts: set = set()
    docpr_counter = [100000]

    def next_rid() -> str:
        rel_counter[0] += 1
        rid = f"rIdM{rel_counter[0]}"
        while rid in existing_ids:
            rel_counter[0] += 1
            rid = f"rIdM{rel_counter[0]}"
        existing_ids.add(rid)
        return rid

    pieces = [master_content, f"<w:p><w:pPr>{master_sect}</w:pPr></w:p>"]

    for index, blob in enumerate(docx_blobs[1:]):
        source = zipfile.ZipFile(io.BytesIO(blob))
        document = source.read("word/document.xml").decode("utf-8")
        _prefix, content, final_sect, _suffix = _body_split(document)
        rel_map = _read_rels(
            source.read("word/_rels/document.xml.rels").decode("utf-8")
        )
        rid_map: dict = {}

        def sub_rid(match, index=index, source=source, rel_map=rel_map, rid_map=rid_map):
            attr, rid = match.group(1), match.group(2)
            if rid not in rel_map:
                raise ValueError(f"{index + 2}번째 문서에 rel 정의가 없음: {rid}")
            if rid not in rid_map:
                rel = rel_map[rid]
                new_id = next_rid()
                if rel["mode"] == "External" or rel["target"].startswith(
                    ("http://", "https://", "mailto:")
                ):
                    new_rels.append(
                        f'<Relationship Id="{new_id}" Type="{rel["type"]}" '
                        f'Target="{_esc(rel["target"])}" TargetMode="External"/>'
                    )
                else:
                    target = rel["target"]
                    base = target.rsplit("/", 1)[-1]
                    directory = target.rsplit("/", 1)[0] if "/" in target else ""
                    new_name = f"merge{index}_{base}"
                    new_target = f"{directory}/{new_name}" if directory else new_name
                    part_path = target if target.startswith("word/") else "word/" + target
                    if part_path not in source.namelist():
                        raise ValueError(f"{index + 2}번째 문서에 파트가 없음: {part_path}")
                    new_parts["word/" + new_target] = source.read(part_path)
                    used_exts.add(base.rsplit(".", 1)[-1].lower())
                    new_rels.append(
                        f'<Relationship Id="{new_id}" Type="{rel["type"]}" '
                        f'Target="{new_target}"/>'
                    )
                rid_map[rid] = new_id
            return f'{attr}="{rid_map[rid]}"'

        content = re.sub(r'(r:[A-Za-z]+)="(rId[0-9]+)"', sub_rid, content)

        def sub_id(match):
            docpr_counter[0] += 1
            return f'{match.group(1)}{docpr_counter[0]}{match.group(3)}'

        content = re.sub(r'(<wp:docPr[^>]*?\bid=")(\d+)(")', sub_id, content)
        content = re.sub(r'(<pic:cNvPr[^>]*?\bid=")(\d+)(")', sub_id, content)

        pieces.append(content)
        if index == len(docx_blobs) - 2:
            pieces.append(final_sect)  # 마지막 구간: 문서 최종 sectPr
        else:
            pieces.append(f"<w:p><w:pPr>{final_sect}</w:pPr></w:p>")

    parts["word/document.xml"] = (prefix + "".join(pieces) + suffix).encode("utf-8")
    rels = rels.replace("</Relationships>", "".join(new_rels) + "</Relationships>")
    parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    for ext in used_exts:
        if f'Extension="{ext}"' not in ctypes:
            ctype = CT_MAP.get(ext)
            if not ctype:
                raise ValueError(f"알 수 없는 미디어 확장자: {ext}")
            ctypes = ctypes.replace(
                "</Types>",
                f'<Default Extension="{ext}" ContentType="{ctype}"/></Types>',
            )
    parts["[Content_Types].xml"] = ctypes.encode("utf-8")

    merged = _zip_bytes(parts, new_parts)

    # 무결성 점검
    with zipfile.ZipFile(io.BytesIO(merged)) as zf:
        if zf.testzip() is not None:
            raise ValueError("병합 결과 zip이 깨졌습니다")
        names = set(zf.namelist())
        for rid, rel in _read_rels(
            parts["word/_rels/document.xml.rels"].decode("utf-8")
        ).items():
            if rel["mode"] == "External":
                continue
            import posixpath

            part = posixpath.normpath(posixpath.join("word", rel["target"]))
            if part not in names:
                raise ValueError(f"rel 대상 파트 누락: {rid} -> {rel['target']}")
        document_text = zf.read("word/document.xml").decode("utf-8")
        used = set(re.findall(r'r:[A-Za-z]+="(rId[^"]+)"', document_text))
        have = set(re.findall(r'Id="(rId[^"]+)"', zf.read("word/_rels/document.xml.rels").decode("utf-8")))
        if used - have:
            raise ValueError(f"정의되지 않은 rel 참조: {sorted(used - have)[:5]}")
        minidom.parseString(document_text)
    return merged


def _doc_stats(docx_bytes: bytes):
    dom = minidom.parseString(
        zipfile.ZipFile(io.BytesIO(docx_bytes)).read("word/document.xml")
    )
    paragraphs = dom.getElementsByTagName("w:p")
    texts = []
    for paragraph in paragraphs:
        buffer = []
        for node in paragraph.getElementsByTagName("w:t"):
            for child in node.childNodes:
                if child.nodeType == child.TEXT_NODE:
                    buffer.append(child.data)
        texts.append("".join(buffer))
    nonempty = [t for t in texts if t.strip()]
    first = nonempty[0] if nonempty else ""
    last = nonempty[-1] if nonempty else ""
    return len(paragraphs), sum(len(t) for t in texts), first, last, texts


def verify_merged(input_blobs: list, merged_bytes: bytes):
    """검증: 통과하지 못하면 (문제 목록, 통계)를 반환한다."""
    problems: list[str] = []
    total_paras = 0
    total_chars = 0
    first_last = []
    for blob in input_blobs:
        paras, chars, first, last, _texts = _doc_stats(blob)
        total_paras += paras
        total_chars += chars
        first_last.append((first, last))
    merged_paras, merged_chars, merged_first, _merged_last, merged_texts = _doc_stats(
        merged_bytes
    )
    if merged_paras < total_paras * 0.95:
        problems.append(
            f"문단 수 미달: {merged_paras} < {total_paras} * 0.95"
        )
    if merged_chars < total_chars * 0.98:
        problems.append(
            f"문자 수 미달: {merged_chars} < {total_chars} * 0.98"
        )
    joined = "\n".join(merged_texts)
    for index, (first, last) in enumerate(first_last):
        if first and first not in joined:
            problems.append(f"{index + 1}번째 구간 첫 문단 누락: {first[:40]!r}")
        if last and last not in joined:
            problems.append(f"{index + 1}번째 구간 끝 문단 누락: {last[:40]!r}")
    for marker in MARKERS:
        if marker not in joined:
            problems.append(f"핵심 표지 누락: {marker!r}")
    if merged_first != FIRST_PARA:
        problems.append(f"첫 문단 불일치: {merged_first[:50]!r}")
    stats = {
        "input_paras": total_paras,
        "input_chars": total_chars,
        "merged_paras": merged_paras,
        "merged_chars": merged_chars,
        "merged_bytes": len(merged_bytes),
        "segments": len(input_blobs),
        "verification_problems": len(problems),
    }
    return problems, stats


@workflows.activity(
    name="joseonsa-merge",
    start_to_close_timeout=MERGE_TIMEOUT,
    retry_policy_max_attempts=1,
)
async def joseonsa_merge(
    file_ids: list[str],
    title: str,
    parent_id: str | None,
    client: ToolCallClient = Depends(drive),
) -> tuple[DriveFileInfo, dict]:
    async def _download(file_id: str) -> bytes:
        result = await client.call_tool(
            tool_name="download_file_content", arguments={"file_id": file_id}
        )
        value = _tool_result_value(result)
        content = value.get("content") if isinstance(value, dict) else value
        if not content:
            raise workflows.WorkflowError(
                f"파일을 내려받지 못했습니다: {file_id}", non_retryable=True
            )
        return base64.b64decode(content)

    if len(file_ids) != 28:
        raise workflows.WorkflowError(
            f"file_ids는 28개여야 합니다(입력 {len(file_ids)}개)",
            non_retryable=True,
        )
    blobs = []
    for file_id in file_ids:
        blobs.append(await _download(file_id))
    txt_bytes = blobs[0]
    docx_blobs = blobs[1:]

    def _build_and_verify() -> tuple[bytes, list, dict]:
        p1_bytes = build_p1_24_docx(txt_bytes, docx_blobs[0])
        merged = merge_docx([p1_bytes] + docx_blobs)
        problems, stats = verify_merged([p1_bytes] + docx_blobs, merged)
        return merged, problems, stats

    merged_bytes, problems, stats = await asyncio.to_thread(_build_and_verify)
    if problems:
        # 검증 실패 시 업로드하지 않는다.
        raise workflows.WorkflowError(
            "검증 실패(업로드 생략): " + "; ".join(problems),
            non_retryable=True,
        )

    async def _unique_title(base_title: str) -> str:
        candidate = base_title
        version = 2
        while True:
            result = await client.call_tool(
                tool_name="search_files",
                arguments={"query": f"title = '{candidate}'", "page_size": 5},
            )
            value = _tool_result_value(result)
            files = value.get("files") if isinstance(value, dict) else value
            if not files:
                return candidate
            candidate = base_title.rsplit(".docx", 1)[0] + f"_v{version}.docx"
            version += 1

    final_title = await _unique_title(title)
    arguments: dict[str, Any] = {
        "title": final_title,
        "base64_content": base64.b64encode(merged_bytes).decode("ascii"),
        "content_mime_type": DOCX_MIME,
        "disable_conversion_to_google_type": True,
    }
    if parent_id:
        arguments["parent_id"] = parent_id
    upload = await client.call_tool(tool_name="create_file", arguments=arguments)
    file = _tool_result_value(upload) or {}
    if not file.get("id"):
        raise workflows.WorkflowError("드라이브 업로드에 실패했습니다")
    info = DriveFileInfo(
        file_id=file.get("id") or "",
        title=file.get("title") or final_title,
        parent_id=file.get("parent_id"),
        mime_type=file.get("mime_type"),
        view_url=file.get("view_url"),
    )
    return info, stats


@workflows.workflow.define(name="joseonsa-merge", on_behalf_of=True)
@uses_connectors(drive)
class JoseonsaMergeWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, data: JoseonsaMergeInput) -> JoseonsaMergeOutput:
        parent = data.output_parent_id or DEFAULT_PARENT
        result_file, stats = await joseonsa_merge(
            data.file_ids, data.output_title, parent
        )
        stats["files_merged"] = len(data.file_ids)
        return JoseonsaMergeOutput(
            result_file=result_file,
            stats=stats,
            summary=(
                f"검수결과 {len(data.file_ids)}개(p1-24 TXT 변환 포함)를 단일 DOCX로 "
                f"병합했습니다. 문단 {stats['merged_paras']:,} / 문자 {stats['merged_chars']:,}, "
                f"검증 문제 {stats['verification_problems']}건. LLM 호출은 없었습니다."
            ),
        )
