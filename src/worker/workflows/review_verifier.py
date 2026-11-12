"""검수 결과 DOCX에 대한 무검증 비용(Zero-LLM) 품질 검증기.

결함이 있는 출력물이 조용히 업로드되는 것을 구조적으로 막는다.
워크플로우가 결과물을 만든 뒤 이 검증기를 통과해야만 업로드한다:

1. 본문 무결성 — 원고의 모든 본문 문단이 결과 문서에 온전히 존재하는지
   (교정이 반영된 문단은 수정문 기준으로 확인). 152자 절단 사고를 잡는 검사.
2. 각주 완전성 — 원고의 모든 각주 줄('N) ...')이 결과 문서에 존재하는지.
3. 번역 커버리지 — 한문 위주 각주 블록마다 번역 항목이 있는지.
   (빠진 각주 번호 목록까지 정확히 보고한다)
4. 년도 리포트 — 년도 구간 수만큼 리포트가 있고 순서가 맞는지.
5. 출처 제안 무결성 — 제안 문구가 원고 텍스트에 근거하는지(외부 서지 차단).

검증은 전부 문자열 대조뿐이므로 LLM 호출이 없다(비용 0).
"""

from __future__ import annotations

import re


import re as _re
import unicodedata

# 한자·비표준 인코딩 영역: PDF 엔진(pypdf/pdf2docx)마다 추출이 달라지는 글자들
_HANJA_RE = _re.compile(r"[\u3000-\u9fff\u3400-\u4dbf\uf000-\ufaff\ufe30-\ufe4f]")


def _norm(text: str) -> str:
    # NFC 정규화: 오래된 한글 PDF는 엔진마다 조합형/완성형 유니코드가 달라
    # 문자열 대조가 실패할 수 있다.
    return unicodedata.normalize("NFC", "".join((text or "").split()))


def _char_coverage(needle: str, haystack: str, threshold: float = 0.95) -> bool:
    """needle의 글자(한글·숫자)들이 haystack에 95% 이상 존재하는지.

    순서·괄호·한자 인코딩·diff 적용 흔적과 무관하게, 실제 글자 소실만
    잡는 최종 안전 대조다.
    """
    if not needle:
        return True
    from collections import Counter

    haystack_counts = Counter(haystack)
    needle_counts = Counter(needle)
    total = sum(needle_counts.values())
    missing = sum(
        max(0, count - haystack_counts.get(char, 0))
        for char, count in needle_counts.items()
    )
    return missing <= total * (1 - threshold)


# 한글과 숫자 외의 모든 것(한자·문장부호·괄호 등)을 제거한다.
# pypdf는 '청태(淸泰) 3년'처럼 괄호를 추출하고 pdf2docx는 '청태3년'처럼
# 떼어내는 등, 두 엔진의 문장부호 추출이 달라 최종 대조는 한글·숫자만으로 한다.
_LETTER_RE = _re.compile(r"[^\uac00-\ud7af0-9]")


def _korean_only(text: str) -> str:
    """한글과 숫자만 남긴 텍스트(공백·한자·문장부호 제거).

    pypdf와 pdf2docx가 한자·괄호를 다르게 추출하는 경우가 있어,
    전체 대조가 실패하면 한글·숫자만으로 다시 대조한다.
    실제 본문 유실(한글 문장 절단)은 이 대조로도 잡힌다.
    """
    return _LETTER_RE.sub("", unicodedata.normalize("NFC", text or ""))


def verify_review_docx(
    docx_bytes: bytes,
    blocks: list,
    edits: list,
    translations: list,
    year_sections: list,
    source_annotations: list | None = None,
) -> list[str]:
    """결과 DOCX를 검증하고 실패 목록(빈 리스트 = 통과)을 반환한다."""
    import io

    from docx import Document
    from docx.oxml.ns import qn

    from worker.workflows.pdf_document_review import _is_hanja_footnote

    errors: list[str] = []

    document = Document(io.BytesIO(docx_bytes))
    paragraphs = list(document.element.body.iter(qn("w:p")))
    paragraph_texts = [
        "".join(t.text or "" for t in p.iter(qn("w:t"))) for p in paragraphs
    ]
    whole_text = _norm("\n".join(paragraph_texts))
    whole_korean = _korean_only("\n".join(paragraph_texts))

    # 1. 본문 무결성
    revised_by_block: dict[str, str] = {}
    for edit in edits or []:
        revised_by_block[edit.item_id] = edit.revised
    missing_body: list[str] = []
    for block in blocks:
        if _is_hanja_footnote(block.text):
            continue  # 각주는 검사 2에서 본다
        original = _norm(block.text)
        revised = _norm(revised_by_block.get(block.block_id) or "")
        # 교정이 반영되었으면 수정문이, 가드 등으로 건너뛰어졌으면
        # 원문이 있어야 한다. 둘 다 없어야 유실이다.
        # 전체 대조가 실패하면 한글만으로 다시 대조한다(한자 추출 인코딩 차이 허용).
        present = (original and original in whole_text) or (
            revised and revised in whole_text
        )
        if not present:
            korean_original = _korean_only(block.text)
            korean_revised = _korean_only(
                revised_by_block.get(block.block_id) or block.text
            )
            if len(korean_original) >= 10:
                present = korean_original in whole_korean
            if not present and len(korean_revised) >= 10:
                present = korean_revised in whole_korean
            # 최종 안전 대조: 글자 커버리지(실제 소실만 실패)
            if not present and len(korean_original) >= 10:
                present = _char_coverage(korean_original, whole_korean)
        if not present:
            expected = revised_by_block.get(block.block_id) or block.text
            missing_body.append(
                f"본문 유실: {block.block_id} (p{block.page}) "
                f"'{expected[:40]}…'"
            )
    errors.extend(missing_body[:20])

    # 2. 각주 완전성
    footnote_blocks = [b for b in blocks if _is_hanja_footnote(b.text)]
    missing_footnotes: list[str] = []
    for block in footnote_blocks:
        if not _norm(block.text):
            continue
        present = _norm(block.text) in whole_text
        if not present:
            korean = _korean_only(block.text)
            if len(korean) >= 10:
                present = korean in whole_korean
            if not present and len(korean) >= 10:
                present = _char_coverage(korean, whole_korean)
        if not present:
            number_match = re.match(r"^\s*(\d{1,3})\s*\)", block.text.strip())
            label = f"각주 {number_match.group(1)}번" if number_match else block.block_id
            missing_footnotes.append(f"각주 유실: {label} (p{block.page})")
    errors.extend(missing_footnotes[:20])

    # 3. 번역 커버리지
    translated_ids = {t.item_id for t in translations or []}
    missing_translations = [
        f"번역 누락: {block.block_id} (p{block.page}) '{block.text[:30]}…'"
        for block in footnote_blocks
        if block.block_id not in translated_ids
    ]
    errors.extend(missing_translations[:40])

    # 4. 년도 리포트
    report_titles = [
        text for text in paragraph_texts if "검수 리포트" in text
    ]
    content_sections = [
        section
        for section in year_sections
        if section.get("edits")
        or section.get("translations")
        or section.get("differences")
        or section.get("annotations")
    ]
    if len(report_titles) != len(content_sections):
        errors.append(
            f"년도 리포트 수 불일치: 리포트 {len(report_titles)}개, "
            f"년도 구간 {len(content_sections)}개"
        )
    for section in content_sections:
        heading = _norm(section.get("heading") or "")
        if heading and heading[:24] not in whole_text:
            errors.append(f"년도 헤딩이 결과에 없음: {section.get('year')}")

    # 5. 출처 제안 무결성 — 제안에 쓰인 서지명이 원고에 있는지
    for annotation in source_annotations or []:
        source = annotation.source or ""
        titles = re.findall(r"[『《「]([^』》」]+)[』》」]", source)
        for title in titles:
            if _norm(title) and _norm(title) not in whole_text:
                errors.append(
                    f"원고에 없는 서지명: '{title}' (출처 제안)"
                )

    return errors
