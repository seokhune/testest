"""PDF 문서 검수 워크플로우.

Google Drive의 PDF 파일에서 텍스트 추출 도구로 본문을 읽어 지정한 페이지
범위만 교정(오타, 문법, 띄어쓰기, 한문 번역)하고, 교정본 DOCX와 검수
리포트 DOCX를 만들어 Drive에 업로드한다.

PDF 바이너리를 다운로드하지 않고 텍스트 표현을 읍어 오므로 커넥터의
base64 2MB 전송 한도보다 큰 파일도 처리할 수 있다.
"""

from __future__ import annotations

import logging
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
    apply_edits_by_text,
    build_structured_docx,
    contains_hanja,
    convert_pdf_to_docx,
    unwrap_text_frames,
)

REVIEW_MODEL = "mistral-large-latest"
LLM_CHUNK_SIZE = 8
LLM_TIMEOUT = timedelta(seconds=900)
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


SENTENCE_ENDINGS = (".", "?", "!", "。")
FOOTNOTE_LINE_RE = re.compile(r"^\d{1,3}\)")


def _paragraphs_from_lines(lines: list[str]) -> list[str]:
    """pypdf의 줄 단위 텍스트를 문단으로 병합한다.

    문장 끝(./?/!/(。))에서 문단을 끊고, 각주 번호(1), 2) 등)로 시작하는
    줄은 자체 문단으로 분리한다. 연도 제목이 줄바꿈으로 잘리는 것도 이
    병합으로 하나의 문단으로 합쳐진다.
    """
    paragraphs: list[str] = []
    current: list[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if current:
                paragraphs.append(" ".join(current))
                current = []
            continue
        if PAGE_MARKER_RE.match(stripped):
            if current:
                paragraphs.append(" ".join(current))
                current = []
            continue
        if FOOTNOTE_LINE_RE.match(stripped) and current:
            paragraphs.append(" ".join(current))
            current = []
        current.append(stripped)
        if stripped.endswith(SENTENCE_ENDINGS) or FOOTNOTE_LINE_RE.match(stripped):
            paragraphs.append(" ".join(current))
            current = []
    if current:
        paragraphs.append(" ".join(current))
    return paragraphs


def blocks_from_pages(pages: list[tuple[int, str]]) -> list[BodyBlock]:
    """물리 페이지 텍스트에서 문단 블록을 만든다."""
    blocks: list[BodyBlock] = []
    for number, text in pages:
        for part in _paragraphs_from_lines(text.split("\n")):
            blocks.append(
                BodyBlock(block_id=f"p{len(blocks)}", text=part, page=number)
            )
    return blocks


# 고려·조선 왕호 (년도 헤딩은 '왕호 + 년수' 형태를 포함한다)
KING_NAMES = (
    "태조|혜종|정종|광종|경종|성종|목종|현종|덕종|숙종|예종|문종|순종|선종|헌종|"
    "인종|의종|명종|신종|희종|강종|고종|원종|충렬왕|충선왕|충숙왕|충혜왕|충목왕|"
    "충정왕|공민왕|공양왕|세종|단종|세조|연산군|중종|선조|광해군|인조|효종|"
    "영조|정조|순조|철종|순종|태종"
)
# 60갑자 간지 (예: 병신, 정유) - 헤딩 앞에 붙는 경우가 있다 ('병신 고려 태조 19년')
GANJI_RE = "[갑을병정무기경신임계][자축인묘진사오미신유술해]"
# 헤딩 판정: '[간지] [고려] 왕호 [한자] 년수(간지)' 형태.
# 왕호 없는 연도('청태 2년', '2년')는 헤딩이 될 수 없으므로 자동으로 제외된다.
# 문단 시작, 문장 끝(.), 닫는 괄호 뒤에만 매칭을 허용해
# 괄호 안 인용문·출전 표기 내부의 연도는 제외한다.
YEAR_HEADING_RE = re.compile(
    rf"(?:^|[.!?)\"\n]\s*)"
    rf"(?P<label>"
    rf"(?:{GANJI_RE}\s*년?\s*)?"
    rf"(?:고려\s*)?"
    rf"(?:{KING_NAMES})"
    rf"(?:\s*\([^)]*\))?"
    rf"\s*(?:\d{{1,3}}\s*년(?:\s*\([^)]*\))?|원년)"
    rf")"
)


def detect_year(text: str) -> str | None:
    """문단에서 년도 헤딩(예: '태조 19년(임신)', '병신 고려 태조 19년')을 찾아 반환한다.

    오탐 방지 규칙:
    - 반드시 왕호가 붙은 년수만 헤딩으로 인정한다. 각주·인용문 내부의
      '청태(淸泰) 2년', '스자쿠(朱雀) 2년' 같은 연호는 왕호가 아니므로 제외된다.
    - 각주 줄('27) ...' 형태)은 헤딩이 될 수 없다.
    - 여는 괄호 바로 뒤의 연도(괄호 안 인용문)는 제외된다.
    """
    stripped = text.strip()
    if FOOTNOTE_LINE_RE.match(stripped):
        return None
    match = YEAR_HEADING_RE.search(stripped[:200])
    if not match:
        return None
    return match.group("label").strip()


def group_by_year(
    blocks: list[BodyBlock],
    body_edits: list[TextEdit],
    hanja_translations: list[HanjaTranslation],
    differences: list[BodyFootnoteDifference] | None = None,
    source_annotations: list[SourceAnnotation] | None = None,
) -> list[dict]:
    """블록 순서대로 년도 구간을 따라가며 검수 결과를 년도별로 묶는다.

    각 구간은 {year, heading, edits, translations, differences, annotations}
    dict로, heading은 그 구간을 시작한 문단의 텍스트다.
    """
    sections: list[dict] = []
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

    def section_for(label: str) -> dict:
        if label not in section_index:
            section_index[label] = len(sections)
            sections.append(
                {
                    "year": label,
                    "heading": "",
                    "edits": [],
                    "translations": [],
                    "differences": [],
                    "annotations": [],
                }
            )
        return sections[section_index[label]]

    for block in blocks:
        year = detect_year(block.text)
        if year:
            current = year
        section = section_for(current)
        if year and not section["heading"]:
            section["heading"] = block.text[:MAX_HEADING_CHARS]
        section["edits"].extend(edits_by_block.get(block.block_id, []))
        section["translations"].extend(translations_by_block.get(block.block_id, []))
        section["differences"].extend(differences_by_block.pop(block.block_id, []))
        section["annotations"].extend(annotations_by_block.pop(block.block_id, []))
    if differences_by_block:
        section = section_for("기타")
        for leftover in differences_by_block.values():
            section["differences"].extend(leftover)
    if annotations_by_block:
        section = section_for("기타")
        for leftover in annotations_by_block.values():
            section["annotations"].extend(leftover)
    return sections


def _normalize_section_item(item, model_class):
    if isinstance(item, model_class):
        return item
    return model_class.model_validate(item)


def _norm_for_match(text: str) -> str:
    return "".join(text.split())


# 같은 문단 안에서 문단 전체가 아닌 변경 구간 발췌만 보여주기 위한 설정
EXCERPT_CONTEXT = 12  # 변경 지점 앞/뒤 문맥 문자 수
EXCERPT_MAX = 160  # 발췌 최대 길이


def _excerpt(text: str, start: int, end: int) -> str:
    """text[start:end] 변경 구간을 앞/뒤 문맥으로 감싸 발췌한다."""
    text = text.strip()
    if not text:
        return ""
    left = max(0, start - EXCERPT_CONTEXT)
    right = min(len(text), end + EXCERPT_CONTEXT)
    excerpt = text[left:right].strip()
    if len(excerpt) > EXCERPT_MAX:
        excerpt = excerpt[:EXCERPT_MAX].rstrip() + "…"
    prefix = "…" if left > 0 else ""
    suffix = "…" if right < len(text) else ""
    return f"{prefix}{excerpt}{suffix}"


def _excerpt_edit_pair(original: str, revised: str) -> tuple[str, str]:
    """original/revised 전체 문단에서 변경된 부분만 발췌하여 (original 발췌, revised 발췌)를 반환한다.

    SequenceMatcher로 변경 구간을 찾아 그 구간 주변만 보여준다. 원문/수정본 일치도가
    너무 낮으면 앞부분 중심 발췌로 대신한다.
    """
    import difflib

    original = (original or "").strip()
    revised = (revised or "").strip()
    if not original or not revised:
        return original[:EXCERPT_MAX], revised[:EXCERPT_MAX]

    matcher = difflib.SequenceMatcher(None, original, revised, autojunk=False)
    first_change = None
    last_change = None
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            if first_change is None:
                first_change = (i1, j1)
            last_change = (i2, j2)
    if first_change is None:
        # 차이가 없음(중복 교정 등) - 앞 부분 발췌
        return _excerpt(original, 0, EXCERPT_MAX), _excerpt(revised, 0, EXCERPT_MAX)

    if last_change[0] - first_change[0] > EXCERPT_MAX * 3:
        # 변경 구간이 지나치게 넓음 - 첫 변경 지점 중심 발췌
        mid = (first_change[0] + last_change[0]) // 2
        return _excerpt(original, mid - 8, mid + 8), _excerpt(
            revised, (first_change[1] + last_change[1]) // 2 - 8,
            (first_change[1] + last_change[1]) // 2 + 8,
        )

    return (
        _excerpt(original, first_change[0], last_change[0]),
        _excerpt(revised, first_change[1], last_change[1]),
    )


def _footnote_label(block_id: str, original_text: str) -> str:
    """주석 발췌 앞의 각주 번호('27) ...')를 찾아 '주석 27번' 라벨을 만든다.

    각주 번호를 못 찾으면 블록 ID로 대신한다.
    """
    match = re.match(r"^\s*(\d{1,3})\s*\)", (original_text or "").strip())
    if match:
        return f"주석 {match.group(1)}번"
    return block_id


def _year_reports(section: dict) -> list:
    year = section["year"]
    edits = section["edits"]
    translations = section["translations"]
    diffs = section["differences"]
    annotations = section["annotations"]
    lines = [
        f"〔{year} 검수 리포트〕",
        (
            f"교정 {len(edits)}건 · 한문 번역 {len(translations)}건 · "
            f"본문-주석 내용 차이 {len(diffs)}건 · 주석 출처 보완 제안 {len(annotations)}건"
        ),
        "1) 교정 내역",
    ]
    if edits:
        for edit in edits:
            original_excerpt, revised_excerpt = _excerpt_edit_pair(edit.original, edit.revised)
            lines.append(f"—원문: {original_excerpt}")
            lines.append(f"—수정: {revised_excerpt}")
    else:
        lines.append("수정된 문단이 없습니다.")
    lines.append("2) 한문 번역 - 주석번역")
    if translations:
        for translation in translations:
            label = (
                f"주석 {translation.footnote_number}번"
                if getattr(translation, "footnote_number", None)
                else _footnote_label(translation.item_id, translation.original)
            )
            lines.append(f"· 원문 - **{label}**: {translation.original}")
            lines.append(f"  번역: {translation.translation}")
    else:
        lines.append("한문 번역 대상이 없습니다.")
    lines.append("3) 본문-주석 내용 차이")
    if diffs:
        for difference in diffs:
            lines.append(f"· {difference.difference}")
    else:
        lines.append(
            "※ 본문-주석 내용 차이 검수는 각주 번호 1:1 매핑 검사(별도 수행)로 진행합니다."
        )
    lines.append("4) 주석 출처 보완 제안")
    if annotations:
        for annotation in annotations:
            label = (
                f"주석 {annotation.footnote_number}번"
                if annotation.footnote_number
                else annotation.item_id
            )
            lines.append(f"· [{label}] 출처: {annotation.source}")
    else:
        lines.append("주석 출처 보완 대상이 없습니다.")
    return lines


def insert_year_reports(docx_bytes: bytes, year_sections: list) -> bytes:
    """각 연도 구간의 내용이 모두 끝나는 지점(다음 연도 제목 직전)에
    그 연도의 검수 리포트 전체를 삽입한다.

    경계는 정규식이 아니라 다음 연도 구간의 제목 문단 텍스트(heading)를
    변환된 문서의 문단에서 내용으로 찾아 정한다. 찾지 못한 구간의
    리포트는 내용 유실을 막기 위해 본문 끝에 순서대로 추가한다.
    """
    import io
    import logging

    from docx import Document
    from docx.oxml.ns import qn

    def _normalize(item, model_class):
        if isinstance(item, model_class):
            return item
        return model_class.model_validate(item)

    sections = []
    for section in year_sections:
        normalized = {
            "year": section["year"],
            "heading": section.get("heading") or "",
            "edits": [_normalize(item, TextEdit) for item in section["edits"]],
            "translations": [
                _normalize(item, HanjaTranslation) for item in section["translations"]
            ],
            "differences": [
                _normalize(item, BodyFootnoteDifference)
                for item in section["differences"]
            ],
            "annotations": [
                _normalize(item, SourceAnnotation) for item in section["annotations"]
            ],
        }
        sections.append(normalized)

    content_sections = [
        section
        for section in sections
        if section["edits"]
        or section["translations"]
        or section["differences"]
        or section["annotations"]
    ]
    if not content_sections:
        return docx_bytes

    document = Document(io.BytesIO(docx_bytes))
    body = document.element.body
    paragraphs = list(body.iter(qn("w:p")))
    paragraph_texts = [
        _norm_for_match("".join(t.text or "" for t in p.iter(qn("w:t"))))
        for p in paragraphs
    ]

    def _find_heading_position(heading: str, start: int) -> int:
        needle = _norm_for_match(heading)[:24]
        if len(needle) < 6:
            return -1
        for index in range(start, len(paragraphs)):
            if needle in paragraph_texts[index]:
                return index
        return -1

    # 각 구간의 제목 문단 위치를 내용 매칭으로 찾는다
    positions: list[int] = []
    search_from = 0
    for section in content_sections:
        position = _find_heading_position(section["heading"], search_from)
        positions.append(position)
        if position >= 0:
            search_from = position + 1

    # 구간 k의 리포트는 구간 k+1의 제목 문단 직전에 삽입된다
    placed: set[int] = set()
    for index in range(len(content_sections) - 1):
        next_position = positions[index + 1]
        if next_position < 0:
            continue
        report_lines = _year_reports(content_sections[index])
        anchor_paragraph = paragraphs[next_position]
        for line in report_lines:
            paragraph = document.add_paragraph(line)
            anchor_paragraph.addprevious(paragraph._p)
        placed.add(index)

    matched = sum(1 for index in range(len(content_sections) - 1) if index in placed)
    logging.getLogger(__name__).info(
        "연도 경계 매칭: %d/%d (제목 미발견 구간은 본문 끝에 추가)",
        matched,
        len(content_sections) - 1,
    )

    # 첫 구간과 경로를 찾지 못한 구간의 리포트는 본문 끝에 순서대로 추가
    anchor = body.find(qn("w:sectPr"))
    if anchor is None:
        anchor = body[-1]
    for index, section in enumerate(content_sections):
        if index in placed:
            continue
        for line in _year_reports(section):
            anchor.addprevious(document.add_paragraph(line)._p)

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


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


def _footnote_number(text: str) -> str | None:
    """'27) 八月…' 형태에서 각주 번호('27')를 뽑는다."""
    match = re.match(r"^\s*(\d{1,3})\s*\)", (text or "").strip())
    return match.group(1) if match else None


def match_translation_to_block(
    translation: HanjaTranslation, blocks: list[BodyBlock]
) -> BodyBlock | None:
    """번역 항목을 원본 블록에 연결한다.

    LLM이 item_id를 다양한 형식('p27', '27', '주석 27번')으로 적어도
    각주 번호 기반으로 찾아낸다. 못 찾으면 텍스트 일치로 찾는다.
    """
    # 1) 원문의 각주 번호가 가장 신뢰할 수 있는 단서다('2) 丙申…').
    original_number = _footnote_number(translation.original or "")
    if original_number:
        try:
            wanted = int(original_number)
            for candidate in blocks:
                candidate_number = _footnote_number(candidate.text)
                if candidate_number and int(candidate_number) == wanted:
                    return candidate
        except ValueError:
            pass
    # 2) 블록 ID 정확 일치. 주의: 블록 ID('p3')에 들어간 숫자를 각주
    #    번호로 해석하면 안 된다 — 각주 3번은 다른 블록이다.
    by_id = {block.block_id: block for block in blocks}
    block = by_id.get(translation.item_id)
    if block is not None:
        return block
    # 3) item_id가 '주석 2번', '2'처럼 각주 번호를 직접 가리킬 때만
    #    번호 매칭을 쓴다('p3' 형태는 2)에서 이미 처리했다).
    if not re.fullmatch(r"p\d+", (translation.item_id or "").strip()):
        id_number = re.search(r"(\d{1,3})", translation.item_id or "")
        if id_number:
            try:
                wanted = int(id_number.group(1))
                for candidate in blocks:
                    candidate_number = _footnote_number(candidate.text)
                    if candidate_number and int(candidate_number) == wanted:
                        return candidate
            except ValueError:
                pass
    # 4) 텍스트 접두사 대조. 블록 쪽에서는 각주 번호 표기('2) ')를
    #    떼고 비교한다(LLM이 번호를 빼고 원문을 �려주는 경우 대비).
    needle = _norm_text(translation.original or "")[:60]
    if needle:
        for candidate in blocks:
            stripped = re.sub(r"^\s*\d{1,3}\s*\)\s*", "", candidate.text)
            if _norm_text(candidate.text).startswith(needle) or _norm_text(
                stripped
            ).startswith(needle):
                return candidate
    return None


def reconcile_translations(
    blocks: list[BodyBlock], translations: list[HanjaTranslation]
) -> list[HanjaTranslation]:
    """LLM 한문 번역 항목을 블록에 정합한다 (reconcile_edits와 같은 규칙)."""
    matched: dict[str, HanjaTranslation] = {}
    for translation in translations:
        block = match_translation_to_block(translation, blocks)
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
1. 본문의 오타자(맞춤법, 오타)를 수정하십시오.
2. 문법에 안 맞는 것은 교정하십시오. 내용을 바꾸지 말고 있는 그대로 교정과 문법에 맞게 윤문하면 됩니다.
3. 절대 다른 자료를 찾아 넣지 마십시오. 여기에 있는 내용만 가지고 하십시오.
4. 원고의 문체를 그대로 유지하십시오.
5. 주석 번호와 주석 인용문은 절대 삭제하거나 변경하지 마십시오.
6. revised는 해당 문단의 전체 최종 텍스트여야 합니다.
7. 변경이 필요한 문단만 edits에 포함하십시오. 변경이 없으면 빈 목록을 반환하십시오.

각 항목은 item_id(대괄호 안 문단 번호), original(원문 전체), revised(수정문 전체), category(교정 유형), reason(변경 이유)을 포함합니다.

문단 목록:
"""

HANJA_TRANSLATE_PROMPT = """아래는 한문 원문 각주 목록입니다. 목록의 모든 항목을 빠짐없이 한국어로 번역하십시오.

규칙:
1. 목록에 있는 각주를 하나도 빠뜨리지 말고 전부 번역하십시오. [N]으로 표시된 항목 수와 번역 결과 수가 같아야 합니다.
2. 한문 구절을 자연스러운 한국어로 번역하십시오.
3. translation에는 번역만 담고 원문 그대로는 포함하지 마십시오.
4. item_id에는 목록의 [N] 번호를 그대로 사용하십시오.

각 항목은 item_id(주석 번호, 예: "주석 23번"), original(한문 구절이 포함된 문단 전체), translation(한문 번역)을 포함합니다.

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

QUOTE_DIFFERENCE_PROMPT = """아래는 번역서 문단 목록입니다. 본문(한국어 번역)과 주석(한문 원문)을 대조하여, 본문이 주석의 원문 내용을 충실히 번역했는지 확인하십시오.

가장 중요한 검사입니다. 절대 다른 자료를 찾지 말고 여기에 있는 내용만으로 판단하십시오.

다음을 찾아 differences에 정리하십시오:
- 본문이 주석 원문의 내용을 누락했거나 추가했거나 왜곡한 경우
- 본문과 주석 원문의 수치·연도·인명·지명이 다른 경우
- 본문이 주석 원문의 의미를 반대로 해석했거나 과장/축소한 경우

차이점은 반드시 다음 형식으로 서술하십시오:
"본문에는 ~~~라고 했는데, 주석 원문에는 ~~~이다. 그러므로 본문은 주석의 내용과 큰 차이가 난다."
또는
"본문이 주석을 왜곡했다. 주석 원문의 ~~~ 부분을 ~~~로 잘못 번역했다."

차이가 없으면 빈 목록을 반환하십시오.

각 항목은 body_block_id(대괄호 안 문단 번호), footnote_id(해당 주석 번호, 모르면 빈 문자열), difference(위 형식의 구체적 서술)을 포함합니다.

문단 목록:
"""


MAX_EDITS = 40
MAX_TRANSLATIONS = 40
MAX_DIFFERENCES = 20
MAX_ANNOTATIONS = 40
# 원고에 존재할 수 없는 현대 서지 표기 - LLM이 지어낸 출처를 코드 레벨에서 차단한다.
FABRICATED_SOURCE_RE = re.compile(
    r"(?:19|20)\d{2}\s*년?|pp?\.|쪽\s*\d|출판부|출판사|박사학위|석사학위|학위논문|"
    r"서울\s*:|법인문화사|탐구당|동국대학교|서울대학교|연세대학교|고려대학교|"
    r"대학교출판|인문사|신서관|을유문화사|솔출판사"
)


def _is_fabricated_source(source: str) -> bool:
    """제안 문구에 현대 서지 표기가 있으면 지어낸 출처로 간주한다."""
    return bool(FABRICATED_SOURCE_RE.search(source or ""))


MAX_ITEM_CHARS = 500
MAX_HEADING_CHARS = 200


def _cap_edits(edits: list) -> list:
    return [
        edit.model_copy(
            update={
                "original": edit.original[:MAX_ITEM_CHARS],
                "revised": edit.revised[:MAX_ITEM_CHARS],
            }
        )
        for edit in edits[:MAX_EDITS]
    ]


def _cap_translations(items: list) -> list:
    return [
        item.model_copy(
            update={
                "original": item.original[:MAX_ITEM_CHARS],
                "translation": item.translation[:MAX_ITEM_CHARS],
            }
        )
        for item in items[:MAX_TRANSLATIONS]
    ]


def _cap_differences(items: list) -> list:
    return [
        item.model_copy(update={"difference": item.difference[:MAX_ITEM_CHARS]})
        for item in items[:MAX_DIFFERENCES]
    ]


def _cap_annotations(items: list) -> list:
    return [
        item.model_copy(update={"source": item.source[:MAX_ITEM_CHARS]})
        for item in items[:MAX_ANNOTATIONS]
    ]


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


@workflows.activity(name="pdf-load-pages", start_to_close_timeout=CONVERT_TIMEOUT)
async def load_pages(
    file_id: str,
    page_start: int | None,
    page_end: int | None,
    client: ToolCallClient = Depends(drive),
) -> tuple[list[tuple[int, str]], int]:
    """원본 PDF에서 지정한 물리 페이지 범위의 텍스트를 페이지별로 추출한다.

    파일 전체를 다운로드해 pypdf로 페이지를 읽으므로, 페이지 범위는
    PDF 뷰어의 물리 페이지 번호(1-based)와 정확히 일치한다. 책 속의
    인쇄 쪽번호가 권마다 다시 시작해도 영향을 받지 않는다.
    다운로드→추출을 한 액티비티에서 처리해 결과(범위 텍스트)만
    반환하므로 액티비티 결과가 base64 2MB 한도를 넘지 않는다.
    """
    import asyncio
    import base64

    result = await client.call_tool(
        tool_name="download_file_content",
        arguments={"file_id": file_id},
    )
    value = _tool_result_value(result)
    content = value.get("content") if isinstance(value, dict) else value
    if not content:
        raise workflows.WorkflowError(
            "PDF 다운로드에 실패했습니다.", non_retryable=True
        )
    pdf_bytes = base64.b64decode(content)

    def _extract() -> tuple[list[tuple[int, str]], int]:
        import io

        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(pdf_bytes))
        total = len(reader.pages)
        first = (page_start or 1) - 1
        last = page_end or total
        if first >= total:
            raise workflows.WorkflowError(
                f"요청한 시작 페이지({page_start})가 전체 페이지 수({total})를 넘습니다.",
                non_retryable=True,
            )
        extracted: list[tuple[int, str]] = []
        for index in range(first, min(last, total)):
            text = reader.pages[index].extract_text() or ""
            extracted.append((index + 1, text.strip()))
        return extracted, total

    return await asyncio.to_thread(_extract)


def _is_hanja_footnote(text: str) -> bool:
    """한문 위주의 각주 블록 판정 - 본문 교정 대상에서 제외한다.

    교수 지침: 한문 원문 각주를 한국어 번역문으로 바꾸는 교정을 제안하지 않는다.
    각주 줄('27) ...')이면서 한문 비중이 높으면 제외한다. 한국어 역자 주는 유지한다.
    """
    stripped = (text or "").strip()
    if not FOOTNOTE_LINE_RE.match(stripped):
        return False
    hanja = sum(1 for ch in stripped if "\u4e00" <= ch <= "\u9fff")
    return hanja >= max(4, len(stripped) * 0.3)


@workflows.activity(name="pdf-proofread-body", start_to_close_timeout=LLM_TIMEOUT)
async def proofread_body(blocks: list[BodyBlock]) -> list[TextEdit]:
    edits: list[TextEdit] = []
    proofread_blocks = [
        block for block in blocks if not _is_hanja_footnote(block.text)
    ]
    for chunk in _chunk(proofread_blocks, LLM_CHUNK_SIZE):
        listing = "\n".join(f"[{block.block_id}] {block.text}" for block in chunk)
        try:
            result = await _parse(_EditBatch, BODY_PROOFREAD_PROMPT + listing)
            edits.extend(result.edits)
        except Exception as error:
            _log_llm_failure("본문 교정", error)
    return _cap_edits(edits)


@workflows.activity(name="pdf-translate-hanja", start_to_close_timeout=LLM_TIMEOUT)
async def translate_hanja_blocks(blocks: list[BodyBlock]) -> list[HanjaTranslation]:
    # 검증기와 같은 기준: 한문 위주 각주만 번역 대상이다.
    targets = [block for block in blocks if _is_hanja_footnote(block.text)]

    async def _attempt_one(block: BodyBlock) -> HanjaTranslation | None:
        listing = f"[{block.block_id}] {block.text}"
        result = await _parse(_HanjaBatch, HANJA_TRANSLATE_PROMPT + listing)
        # 각주 번호는 원고 블록에서 확정한다(LLM 인용은 번호를 빼는 경우가 있다).
        footnote_number = _footnote_number(block.text)
        for translation in result.translations:
            matched = match_translation_to_block(translation, [block])
            if matched is not None:
                return translation.model_copy(
                    update={
                        "item_id": block.block_id,
                        "footnote_number": footnote_number,
                    }
                )
        return None

    translations: list[HanjaTranslation] = []
    # 한 각주당 호출 하나: 항목이 독립적으로 처리되어 부분 누락이
    # 재시도로 정확히 복구된다.
    for block in targets:
        got: HanjaTranslation | None = None
        for attempt in range(3):
            try:
                got = await _attempt_one(block)
                if got is not None:
                    break
            except Exception as error:
                _log_llm_failure(f"한문 번역({block.block_id}, {attempt + 1}차)", error)
        if got is None:
            logging.getLogger(__name__).warning(
                "한문 번역 최종 누락: %s (p%s)", block.block_id, block.page
            )
        else:
            translations.append(got)
    return _cap_translations(translations)


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
    return _cap_edits(edits)


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
    return _cap_differences(differences)


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
    # 각주 번호('27) ...')를 블록 텍스트에서 추출해 리포트 라벨로 쓴다.
    block_text_by_id = {block.block_id: block.text for block in blocks}
    for annotation in annotations:
        block_text = block_text_by_id.get(annotation.item_id, "")
        number_match = re.match(r"^\s*(\d{1,3})\s*\)", (block_text or "").strip())
        if number_match:
            annotation.footnote_number = number_match.group(1)
    # '원고에 없는 자료 금지'를 프롬프트만 믿지 말고 코드로 검사한다.
    kept = [a for a in annotations if not _is_fabricated_source(a.source)]
    dropped = len(annotations) - len(kept)
    if dropped:
        logging.getLogger(__name__).info(
            "지어낸 외부 서지 제안 %d건 차단", dropped
        )
    return _cap_annotations(kept)


@workflows.activity(name="pdf-build-result", start_to_close_timeout=CONVERT_TIMEOUT)
async def build_result_docx(
    source: DriveFileInfo,
    first_page: int,
    last_page: int,
    blocks: list[BodyBlock],
    total_pages: int,
    year_sections: list,
    stats: dict[str, int],
    client: ToolCallClient = Depends(drive),
) -> tuple[str, list[str]]:
    """검수 결과 한 파일(원본 양식 본문 + 년도별 리포트)을 만들어
    (base64, 품질 검증 오류 목록)을 반환한다. 오류 목록이 비어 있어야 업로드한다.

    다운로드→변환→교정을 한 액티비티 안에서 처리하고 작은 DOCX만
    반환한다(액티비티 결과가 base64 2MB 이벤트 한도를 넘지 않게).
    원본 PDF를 pdf2docx로 변환해 양식을 유지하고, 그 뒤에 년도별
    교정 내역/한문 번역 리포트를 붙인다. 변환에 실패하면 페이지
    구조를 재구성한 본문으로 대체한다.
    """
    import asyncio
    import base64

    # 년도 구간에서 평탄한 교정/번역 목록을 다시 만든다(중복 전달 제거)
    flat_edits = [
        _normalize_section_item(item, TextEdit)
        for section in year_sections
        for item in section["edits"]
    ]
    flat_translations = [
        _normalize_section_item(item, HanjaTranslation)
        for section in year_sections
        for item in section["translations"]
    ]
    pages_for_fallback = [
        (block.page, block.text) for block in blocks
    ]

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

        start = first_page - 1
        end = last_page or None
        docx_bytes = await asyncio.to_thread(
            convert_pdf_to_docx, pdf_bytes, start, end
        )
        docx_bytes = await asyncio.to_thread(unwrap_text_frames, docx_bytes)
        edited, _applied = await asyncio.to_thread(
            apply_edits_by_text, docx_bytes, flat_edits, flat_translations, blocks
        )
    except Exception as error:
        import logging

        logging.getLogger(__name__).warning(
            "원본 양식 변환 실패, 구조 재구성으로 대체: %s", error
        )
        edited, _applied = await asyncio.to_thread(
            build_structured_docx, pages_for_fallback, blocks, flat_edits, flat_translations
        )

    final = await asyncio.to_thread(insert_year_reports, edited, year_sections)

    # 품질 검증(무검증 비용): 통과해야만 업로드된다. 결함이 있으면
    # 실행이 실패하고 원인 목록이 에러 메시지에 담긴다.
    from worker.workflows.review_verifier import verify_review_docx

    flat_annotations = [
        _normalize_section_item(item, SourceAnnotation)
        for section in year_sections
        for item in section["annotations"]
    ]
    verification_errors = await asyncio.to_thread(
        verify_review_docx,
        final,
        blocks,
        flat_edits,
        flat_translations,
        year_sections,
        flat_annotations,
    )
    # 검증 실패 사유는 워크플로우 레벨에서 보고한다(플랫폼이 활동 내부
    # 예외 메시지를 삼키기 때문).
    return base64.b64encode(final).decode(), verification_errors


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

        pages, total_pages = await load_pages(
            source.file_id, data.page_start, data.page_end
        )
        blocks = blocks_from_pages(pages)

        async def _guarded(name: str, call):
            try:
                return await call
            except Exception as error:
                raise workflows.WorkflowError(
                    f"{name} 실패: {type(error).__name__}: {str(error)[:300]}",
                    non_retryable=True,
                ) from error

        body_edits = await _guarded("본문 교정", proofread_body(blocks))
        hanja_translations = await _guarded(
            "한문 번역", translate_hanja_blocks(blocks)
        )
        source_annotations = await _guarded(
            "주석 출처 판독", annotate_footnote_sources(blocks)
        )
        # 본문-주석 내용 차이는 워크플로우에서 수행하지 않는다.
        # 페이지 단위 LLM 비교가 각주 번호 미매핑 환각을 일으켰기 때문에,
        # 이 검사는 Work의 joseonsa-body-footnote-diff 스킬(각주 번호 1:1 매핑)이 담당한다.
        differences = []

        body_edits = reconcile_edits(blocks, body_edits)
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

        result_b64, verification_errors = await _guarded(
            "결과 문서 생성", build_result_docx(
            source,
            pages[0][0] if pages else 1,
            pages[-1][0] if pages else total_pages,
            blocks,
            total_pages,
            year_sections,
            stats,
        ))

        if verification_errors:
            raise workflows.WorkflowError(
                "출력 품질 검증 실패 (업로드 중단): "
                + " | ".join(verification_errors[:15]),
                non_retryable=True,
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
            section["year"]
            for section in year_sections
            if section["edits"]
            or section["translations"]
            or section["differences"]
            or section["annotations"]
        )
        summary = (
            f"본문 {len(blocks)}문단을 검수해{page_note} "
            f"{len(body_edits)}건을 교정하고 한문 번역 {len(hanja_translations)}건을 추가했으며, "
            f"본문-주석 내용 차이 {len(differences)}건과 주석 출처 보완 제안 "
            f"{len(source_annotations)}건을 확인했습니다. "
            f"년도 구간: {year_note}"
        )
        return PdfReviewOutput(
            source_title=source.title,
            result_file=result_file,
            stats=stats,
            summary=summary,
        )
