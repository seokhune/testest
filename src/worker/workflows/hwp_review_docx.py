"""DOCX parsing, rewriting, report building, and Hancom format conversion.

Helpers used by the HWP review workflow activities. All functions are plain
synchronous code: network I/O and subprocess calls stay in the activities that
call them.
"""

from __future__ import annotations

import base64
import copy
import io
import subprocess
import tempfile
from pathlib import Path

from lxml import etree

from worker.workflows.hwp_review_models import (
    BodyBlock,
    BodyFootnoteDifference,
    ExtractedDocument,
    FootnoteItem,
    HanjaTranslation,
    SourceAnnotation,
    TextEdit,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
FOOTNOTES_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)
SEPARATOR_FOOTNOTE_IDS = {"0", "-1", "none"}

DOCX_MIME_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

HANCOM_EXTENSIONS = (".hwp", ".hwpx")
HANCOM_INFILTERS = {"hwp": "Hwp2002_File", "hwpx": "Hwp2002_Reader"}
DOCX_EXTENSION = ".docx"
PDF_EXTENSION = ".pdf"
ACCEPTED_SOURCE_EXTENSIONS = HANCOM_EXTENSIONS + (DOCX_EXTENSION, PDF_EXTENSION)


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _element_text(element: etree._Element) -> str:
    return "".join(t.text or "" for t in element.iter(_w("t")))


def _page_break_count(element: etree._Element) -> int:
    count = 0
    for br in element.iter(_w("br")):
        if br.get(_w("type")) == "page":
            count += 1
    for _ in element.iter(_w("lastRenderedPageBreak")):
        count += 1
    return count


def _footnote_refs(element: etree._Element) -> set[str]:
    return {
        ref.get(_w("id"))
        for ref in element.iter(_w("footnoteReference"))
        if ref.get(_w("id")) is not None
    }


def contains_hanja(text: str) -> bool:
    return any(
        "\u3400" <= ch <= "\u4dbf"
        or "\u4e00" <= ch <= "\u9fff"
        or "\uf900" <= ch <= "\ufaff"
        for ch in text
    )


def _footnotes_part(document):
    for rel in document.part.rels.values():
        if rel.reltype == FOOTNOTES_RELTYPE:
            return rel.target_part
    return None


def _footnotes_root(document) -> etree._Element | None:
    part = _footnotes_part(document)
    if part is None:
        return None
    return etree.fromstring(part.blob)


def _body_paragraphs(document) -> list[etree._Element]:
    return [p for p in document.element.body.iter(_w("p")) if _element_text(p).strip()]


def _footnote_elements(root: etree._Element) -> dict[str, etree._Element]:
    footnotes: dict[str, etree._Element] = {}
    for footnote in root.findall(_w("footnote")):
        footnote_id = footnote.get(_w("id"))
        if footnote_id is None or footnote_id in SEPARATOR_FOOTNOTE_IDS:
            continue
        footnotes[footnote_id] = footnote
    return footnotes


def _first_paragraph(element: etree._Element) -> etree._Element | None:
    paragraphs = element.findall(_w("p"))
    return paragraphs[0] if paragraphs else None


def _replace_paragraph_text(paragraph: etree._Element, revised: str) -> None:
    rpr = None
    for run in paragraph.iter(_w("r")):
        candidate = run.find(_w("rPr"))
        if candidate is not None:
            rpr = copy.deepcopy(candidate)
            break
    for run in list(paragraph.iter(_w("r"))):
        run.getparent().remove(run)
    for wrapper in list(paragraph.iter(_w("hyperlink"))):
        wrapper.getparent().remove(wrapper)
    run = etree.SubElement(paragraph, _w("r"))
    if rpr is not None:
        run.append(rpr)
    t = etree.SubElement(run, _w("t"))
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = revised


def _append_run(paragraph: etree._Element, text: str) -> None:
    run = etree.SubElement(paragraph, _w("r"))
    t = etree.SubElement(run, _w("t"))
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text


def extract_document(
    docx_bytes: bytes,
    source_format: str,
    page_start: int | None = None,
    page_end: int | None = None,
    page_texts: list[str] | None = None,
) -> ExtractedDocument:
    from docx import Document

    document = Document(io.BytesIO(docx_bytes))
    body: list[BodyBlock] = []
    block_refs: dict[str, set[str]] = {}
    page = 1
    total_pages = 1
    nonempty_index = 0
    for paragraph in document.element.body.iter(_w("p")):
        text = _element_text(paragraph).strip()
        if text:
            block_id = f"p{nonempty_index}"
            body.append(BodyBlock(block_id=block_id, text=text, page=page))
            block_refs[block_id] = _footnote_refs(paragraph)
            nonempty_index += 1
        page += _page_break_count(paragraph)
        if page > total_pages:
            total_pages = page

    if page_texts is not None:
        pages = assign_pages_from_pdf([block.text for block in body], page_texts)
        for block, block_page in zip(body, pages, strict=True):
            block.page = block_page
        total_pages = max(len(page_texts), 1)

    footnotes: list[FootnoteItem] = []
    footnotes_root = _footnotes_root(document)
    if footnotes_root is not None:
        for footnote_id, footnote in _footnote_elements(footnotes_root).items():
            text = _element_text(footnote).strip()
            if text:
                footnotes.append(FootnoteItem(footnote_id=footnote_id, text=text))

    if page_start is not None or page_end is not None:
        start = page_start or 1
        end = page_end or total_pages
        body = [block for block in body if start <= block.page <= end]
        kept_refs = set().union(*(block_refs[block.block_id] for block in body)) if body else set()
        footnotes = [f for f in footnotes if f.footnote_id in kept_refs]

    return ExtractedDocument(
        source_format=source_format,
        docx_b64=base64.b64encode(docx_bytes).decode(),
        body=body,
        footnotes=footnotes,
        total_pages=total_pages,
        page_start=page_start,
        page_end=page_end,
    )


def apply_review_edits(
    docx_bytes: bytes,
    body_edits: list[TextEdit],
    footnote_edits: list[TextEdit],
    source_annotations: list[SourceAnnotation],
    hanja_translations: list[HanjaTranslation],
) -> bytes:
    from docx import Document

    document = Document(io.BytesIO(docx_bytes))
    body_by_id = {f"p{index}": p for index, p in enumerate(_body_paragraphs(document))}

    for edit in body_edits:
        paragraph = body_by_id.get(edit.item_id)
        if paragraph is not None:
            _replace_paragraph_text(paragraph, edit.revised)

    footnotes_root = _footnotes_root(document)
    if footnotes_root is not None:
        footnotes = _footnote_elements(footnotes_root)
        for edit in footnote_edits:
            footnote = footnotes.get(edit.item_id)
            if footnote is None:
                continue
            paragraph = _first_paragraph(footnote)
            if paragraph is None:
                paragraph = etree.SubElement(footnote, _w("p"))
            _replace_paragraph_text(paragraph, edit.revised)
        for annotation in source_annotations:
            footnote = footnotes.get(annotation.item_id)
            if footnote is None:
                continue
            paragraph = _first_paragraph(footnote)
            if paragraph is None:
                paragraph = etree.SubElement(footnote, _w("p"))
            _append_run(paragraph, f" [출처: {annotation.source}]")
        for translation in hanja_translations:
            footnote = footnotes.get(translation.item_id)
            if footnote is None:
                continue
            translated = etree.SubElement(footnote, _w("p"))
            _append_run(translated, f"번역: {translation.translation}")
        part = _footnotes_part(document)
        if part is not None:
            part._blob = etree.tostring(
                footnotes_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def build_report_docx(
    source_title: str,
    body_edits: list[TextEdit],
    footnote_edits: list[TextEdit],
    source_annotations: list[SourceAnnotation],
    hanja_translations: list[HanjaTranslation],
    differences: list[BodyFootnoteDifference],
) -> bytes:
    from docx import Document

    document = Document()
    document.add_heading("문서 검수 리포트", level=0)
    document.add_paragraph(f"원본 파일: {source_title}")

    document.add_heading("1. 본문 교정 (오탈자·윤문)", level=1)
    if not body_edits:
        document.add_paragraph("변경 사항 없음")
    for edit in body_edits:
        document.add_paragraph(f"[{edit.item_id}] ({edit.category}) {edit.reason}")
        document.add_paragraph(f"원문: {edit.original}")
        document.add_paragraph(f"수정: {edit.revised}")

    document.add_heading("2. 주석 교정 (오탈자·윤문·서지명 〈 〉)", level=1)
    if not footnote_edits:
        document.add_paragraph("변경 사항 없음")
    for edit in footnote_edits:
        document.add_paragraph(f"[주석 {edit.item_id}] ({edit.category}) {edit.reason}")
        document.add_paragraph(f"원문: {edit.original}")
        document.add_paragraph(f"수정: {edit.revised}")

    document.add_heading("3. 주석 출처 보완", level=1)
    if not source_annotations:
        document.add_paragraph("추가된 출처 없음")
    for annotation in source_annotations:
        document.add_paragraph(f"[주석 {annotation.item_id}] 출처: {annotation.source}")

    document.add_heading("4. 한문 주석 번역", level=1)
    if not hanja_translations:
        document.add_paragraph("한문 주석 없음")
    for translation in hanja_translations:
        document.add_paragraph(
            f"[주석 {translation.item_id}] 원문: {translation.original}"
        )
        document.add_paragraph(f"번역: {translation.translation}")

    document.add_heading("5. 본문-주석 차이점", level=1)
    if not differences:
        document.add_paragraph("차이점 없음")
    for difference in differences:
        location = difference.body_block_id or "-"
        footnote = difference.footnote_id or "-"
        document.add_paragraph(
            f"[본문 {location} / 주석 {footnote}] {difference.difference}"
        )

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _soffice_convert(
    source_bytes: bytes, source_name: str, target_format: str, infilter: str | None = None
) -> bytes:
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_path = tmp / source_name
        source_path.write_bytes(source_bytes)
        output_path = tmp / f"{source_path.stem}.{target_format}"
        attempts: list[list[str]] = []
        if infilter:
            attempts.append(["--infilter", infilter])
        attempts.append([])
        last_error = ""
        for extra_args in attempts:
            output_path.unlink(missing_ok=True)
            completed = subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--norestore",
                    *extra_args,
                    "--convert-to",
                    target_format,
                    "--outdir",
                    str(tmp),
                    str(source_path),
                ],
                capture_output=True,
                check=False,
                timeout=300,
            )
            if output_path.exists() and output_path.stat().st_size > 0:
                return output_path.read_bytes()
            last_error = (completed.stderr or b"").decode(errors="replace").strip()
    raise RuntimeError(
        f"LibreOffice {target_format} 변환에 실패했습니다."
        + (f" ({last_error})" if last_error else "")
    )


def convert_hancom_to_docx(source_bytes: bytes, source_ext: str) -> bytes:
    ext = source_ext.lower().lstrip(".")
    if ext not in HANCOM_INFILTERS:
        raise ValueError(f"지원하지 않는 한컴 문서 형식입니다: {source_ext}")
    return _soffice_convert(source_bytes, f"source.{ext}", "docx", HANCOM_INFILTERS[ext])


def convert_docx_to_pdf(docx_bytes: bytes) -> bytes:
    return _soffice_convert(docx_bytes, "source.docx", "pdf")


def convert_pdf_to_docx(pdf_bytes: bytes) -> bytes:
    """PDF를 양식을 보존하는 DOCX로 변환한다 (pdf2docx 기반)."""
    from pdf2docx import Converter

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_path = tmp / "source.pdf"
        source_path.write_bytes(pdf_bytes)
        output_path = tmp / "source.docx"
        converter = Converter(str(source_path))
        try:
            converter.convert(str(output_path))
        finally:
            converter.close()
        if output_path.exists() and output_path.stat().st_size > 0:
            return output_path.read_bytes()
    raise RuntimeError("PDF->DOCX 변환에 실패했습니다.")


def pdf_page_texts(pdf_bytes: bytes) -> list[str]:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(pdf_bytes))
    return [(page.extract_text() or "") for page in reader.pages]


def _normalize_for_match(text: str) -> str:
    return "".join(text.split())


def assign_pages_from_pdf(block_texts: list[str], page_texts: list[str]) -> list[int]:
    if not page_texts:
        return [1] * len(block_texts)
    boundaries: list[int] = []
    combined = ""
    for page_text in page_texts:
        boundaries.append(len(combined))
        combined += _normalize_for_match(page_text)

    def _page_at(offset: int) -> int:
        page = 1
        for index, boundary in enumerate(boundaries):
            if offset >= boundary:
                page = index + 1
        return page

    pages: list[int] = []
    cursor = 0
    last_page = 1
    for text in block_texts:
        needle = _normalize_for_match(text)
        if not needle:
            pages.append(last_page)
            continue
        found = -1
        for length in (min(len(needle), 60), min(len(needle), 20), min(len(needle), 8)):
            found = combined.find(needle[:length], cursor)
            if found != -1:
                cursor = found + length
                break
        if found == -1:
            found = combined.find(needle, cursor)
            if found != -1:
                cursor = found + len(needle)
        if found == -1:
            pages.append(last_page)
            continue
        last_page = _page_at(found)
        pages.append(last_page)
    return pages


def apply_edits_by_text(
    docx_bytes: bytes,
    body_edits: list[TextEdit],
    hanja_translations: list[HanjaTranslation],
) -> tuple[bytes, int]:
    """텍스트 매칭으로 문단을 찾아 교정/한문 번역을 그 자리에 반영한다.

    PDF에서 변환한 DOCX는 원본 문단 구조와 블록 순서가 어긋날 수 있어
    항목 ID 대신 원문 텍스트 일치로 문단을 찾는다. 반환값은 (결과 bytes,
    반영된 교정 수).
    """
    from docx import Document

    def _norm(text: str) -> str:
        return "".join(text.split())

    document = Document(io.BytesIO(docx_bytes))
    paragraphs = [p for p in document.element.body.iter(_w("p")) if _element_text(p).strip()]
    translated: set[int] = set()

    for translation in hanja_translations:
        needle = _norm(translation.original)
        for i, paragraph in enumerate(paragraphs):
            text = _norm(_element_text(paragraph))
            if i in translated:
                continue
            if text == needle or (needle and text.startswith(needle[:60])):
                translated_p = etree.Element(_w("p"))
                run = etree.SubElement(translated_p, _w("r"))
                t = etree.SubElement(run, _w("t"))
                t.set(f"{{{XML_NS}}}space", "preserve")
                t.text = f"번역: {translation.translation}"
                paragraph.addnext(translated_p)
                translated.add(i)
                break

    applied = 0
    used: set[int] = set()
    for edit in body_edits:
        needle = _norm(edit.original)
        match_index = None
        exact = {i: _norm(_element_text(p)) for i, p in enumerate(paragraphs)}
        for i, text in exact.items():
            if i in used:
                continue
            if text == needle or (needle and text.startswith(needle[:60])):
                match_index = i
                break
        if match_index is None:
            continue
        _replace_paragraph_text(paragraphs[match_index], edit.revised)
        used.add(match_index)
        applied += 1

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue(), applied


def build_structured_docx(
    pages: list[tuple[int, str]],
    blocks: list,
    body_edits: list[TextEdit],
    hanja_translations: list[HanjaTranslation],
) -> tuple[bytes, int]:
    """추출된 페이지 텍스트로 원본 구조를 재구성한 DOCX를 만든다.

    각 페이지는 원본 쪽번호 표식(- 8 - 형태)과 문단 순서를 유지한다.
    문단은 pages와 동일한 순서로 만들어지므로 blocks의 block_id로
    교정/번역을 정확히 그 자리에 반영한다.
    """
    import re
    import io

    from docx import Document
    from docx.enum.text import WD_BREAK

    edits_by_id = {edit.item_id: edit for edit in body_edits}
    hanja_by_id = {t.item_id: t for t in hanja_translations}
    applied = 0

    document = Document()
    paragraph_index = 0
    for page_index, (page_number, page_text) in enumerate(pages):
        if page_index > 0:
            document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        document.add_paragraph(f"- {page_number} -")
        for paragraph_text in re.split(r"\n\s*\n", page_text):
            cleaned = paragraph_text.strip()
            if not cleaned:
                continue
            block_id = f"p{paragraph_index}"
            edit = edits_by_id.get(block_id)
            document.add_paragraph(edit.revised if edit else cleaned)
            if edit:
                applied += 1
            translation = hanja_by_id.get(block_id)
            if translation:
                document.add_paragraph(f"번역: {translation.translation}")
            paragraph_index += 1

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue(), applied
