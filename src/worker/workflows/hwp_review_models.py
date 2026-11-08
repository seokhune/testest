"""Pydantic models for the HWP document review workflow."""

import pydantic


class DriveFileInfo(pydantic.BaseModel):
    file_id: str
    title: str
    parent_id: str | None = None
    mime_type: str | None = None
    view_url: str | None = None
    file_size: int | None = None


class BodyBlock(pydantic.BaseModel):
    block_id: str
    text: str
    page: int = 1


class FootnoteItem(pydantic.BaseModel):
    footnote_id: str
    text: str


class ExtractedDocument(pydantic.BaseModel):
    source_format: str
    docx_b64: str
    body: list[BodyBlock]
    footnotes: list[FootnoteItem]
    total_pages: int = 1
    page_start: int | None = None
    page_end: int | None = None


class TextEdit(pydantic.BaseModel):

    @pydantic.field_validator("item_id", mode="before")
    @classmethod
    def _stringify_item_id(cls, value):
        if isinstance(value, (int, float)):
            return str(value)
        return value

    item_id: str
    original: str
    revised: str
    category: str
    reason: str


class SourceAnnotation(pydantic.BaseModel):
    item_id: str
    source: str
    footnote_number: str | None = None

    @pydantic.field_validator("item_id", mode="before")
    @classmethod
    def _stringify_item_id(cls, value):
        if isinstance(value, (int, float)):
            return str(value)
        return value


class HanjaTranslation(pydantic.BaseModel):

    @pydantic.field_validator("item_id", mode="before")
    @classmethod
    def _stringify_item_id(cls, value):
        if isinstance(value, (int, float)):
            return str(value)
        return value

    item_id: str
    original: str
    translation: str
    footnote_number: str | None = None


def _coerce_str(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return str(value)
    return value


class BodyFootnoteDifference(pydantic.BaseModel):
    body_block_id: str | None = None
    footnote_id: str | None = None
    difference: str

    @pydantic.field_validator("body_block_id", "footnote_id", mode="before")
    @classmethod
    def _stringify_ids(cls, value):
        return _coerce_str(value)


class HwpReviewInput(pydantic.BaseModel):
    file_name: str
    output_parent_id: str | None = None
    page_start: int | None = None
    page_end: int | None = None


class HwpReviewOutput(pydantic.BaseModel):
    source_title: str
    working_copy_title: str
    corrected_file: DriveFileInfo
    report_file: DriveFileInfo
    stats: dict[str, int]
    summary: str


class PdfReviewInput(pydantic.BaseModel):
    file_name: str
    output_parent_id: str | None = None
    page_start: int | None = None
    page_end: int | None = None


class PdfReviewOutput(pydantic.BaseModel):
    source_title: str
    result_file: DriveFileInfo
    stats: dict[str, int]
    summary: str
