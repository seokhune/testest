"""페이지 범위(page_start/page_end) 슬라이싱 테스트."""

import io
import zipfile

from worker.workflows.hwp_review_docx import apply_review_edits, extract_document
from worker.workflows.hwp_review_models import TextEdit

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOCUMENT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"""


def _p(text: str, footnote_id: str | None = None) -> str:
    ref = (
        f'<w:r><w:footnoteReference w:id="{footnote_id}"/></w:r>' if footnote_id else ""
    )
    return f'<w:p xmlns:w="{W_NS}"><w:r><w:t>{text}</w:t></w:r>{ref}</w:p>'


def _page_break() -> str:
    return f'<w:p xmlns:w="{W_NS}"><w:r><w:br w:type="page"/></w:r></w:p>'


DOCUMENT_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}"><w:body>{_p("첫째 페이지 문단.")}{_page_break()}{_p("둘째 페이지 문단.")}{_page_break()}{_p("셋째 페이지 문단.", footnote_id="2")}<w:sectPr/></w:body></w:document>"""

FOOTNOTES_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W_NS}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>첫째 페이지 각주.</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>셋째 페이지 각주.</w:t></w:r></w:p></w:footnote>
</w:footnotes>"""


def make_paged_docx() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        archive.writestr("_rels/.rels", ROOT_RELS_XML)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)
        archive.writestr("word/document.xml", DOCUMENT_XML)
        archive.writestr("word/footnotes.xml", FOOTNOTES_XML)
    return buffer.getvalue()


def test_extract_assigns_page_numbers() -> None:
    document = extract_document(make_paged_docx(), source_format="docx")

    assert [block.page for block in document.body] == [1, 2, 3]
    assert document.total_pages == 3
    assert [f.footnote_id for f in document.footnotes] == ["1", "2"]
    assert document.page_start is None and document.page_end is None


def test_extract_page_range_slices_body_and_footnotes() -> None:
    document = extract_document(
        make_paged_docx(), source_format="docx", page_start=2, page_end=3
    )

    assert [block.block_id for block in document.body] == ["p1", "p2"]
    assert [block.page for block in document.body] == [2, 3]
    assert document.total_pages == 3
    assert document.page_start == 2 and document.page_end == 3
    assert [f.footnote_id for f in document.footnotes] == ["2"]


def test_apply_edits_uses_original_block_ids_after_slicing() -> None:
    edited = apply_review_edits(
        make_paged_docx(),
        [
            TextEdit(
                item_id="p2",
                original="셋째 페이지 문단.",
                revised="셋째 페이지 수정 문단.",
                category="typo",
                reason="오탈자 수정",
            )
        ],
        [],
        [],
        [],
    )

    document = extract_document(edited, source_format="docx")
    assert document.body[2].text == "셋째 페이지 수정 문단."
    assert document.body[2].page == 3
