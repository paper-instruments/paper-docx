"""Field authoring — formulas, never values.

A field is a formula; static text is a pasted value. This module authors
PAGE/NUMPAGES/DATE simple fields, REF/PAGEREF cross-references, and the TOC
complex field. Every inserted field carries placeholder result text and
sets the document's update-fields-on-open flag: **this package never
computes a field's value** — pagination and evaluation belong to a renderer
(Word, or headless LibreOffice in the harness).

The `in_field` guard recognizes everything authored here, so a span
landing inside one of our own fields refuses exactly like one landing in
Word's (self-consistency, tested).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional, Tuple

from docx._guard import check_install
from docx._transaction import rollback_on_error
from docx.errors import TargetNotFoundError
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import _refuse_if_protected

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document
    from docx.text.paragraph import Paragraph

check_install()

#: placeholder result texts — deliberately obviously-stale until a renderer
#: recomputes them on open (w:updateFields)
_PLACEHOLDERS = {
    "PAGE": "1",
    "NUMPAGES": "1",
    "DATE": "(date)",
}

_REFERENCE_KINDS = {
    "text": lambda name: f" REF {name} \\h ",
    "page": lambda name: f" PAGEREF {name} \\h ",
    "number": lambda name: f" REF {name} \\r \\h ",
}

#: CT_Settings children that FOLLOW w:updateFields in the schema sequence
#: (insertion point; see docx.oxml.settings._tag_seq index 76)
_UPDATE_FIELDS_SUCCESSORS = (
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
    "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema",
    "w:themeFontLang", "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions",
    "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary",
    "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol",
    "w:listSeparator",
)


def _document_of_paragraph(paragraph: "Paragraph") -> "Document":
    part = paragraph.part
    document = getattr(part, "document", None)  # header/footer parts lack it
    if document is not None:
        return document
    return part.package.main_document_part.document


def _set_update_fields_on_open(document: "Document") -> None:
    """`w:updateFields w:val="true"` — the renderer recomputes every field
    result the next time the document opens."""
    settings = document.settings.element
    existing = settings.find(qn("w:updateFields"))
    if existing is not None:
        existing.set(qn("w:val"), "true")
        return
    element = OxmlElement("w:updateFields")
    element.set(qn("w:val"), "true")
    settings.insert_element_before(
        element, *_UPDATE_FIELDS_SUCCESSORS
    )


def _simple_field(instr: str, placeholder: str) -> "_Element":
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    run.add_t(placeholder)
    field.append(run)
    return field


def add_page_number_field(paragraph: "Paragraph") -> None:
    """Append a PAGE field (typically to a footer paragraph)."""
    document = _document_of_paragraph(paragraph)
    _refuse_if_protected(document, "insert a field")
    with rollback_on_error(document):
        paragraph._p.append(_simple_field(" PAGE ", _PLACEHOLDERS["PAGE"]))
        _set_update_fields_on_open(document)


def add_page_count_field(paragraph: "Paragraph") -> None:
    """Append a NUMPAGES field (typically to a footer paragraph)."""
    document = _document_of_paragraph(paragraph)
    _refuse_if_protected(document, "insert a field")
    with rollback_on_error(document):
        paragraph._p.append(_simple_field(" NUMPAGES ", _PLACEHOLDERS["NUMPAGES"]))
        _set_update_fields_on_open(document)


def add_date_field(paragraph: "Paragraph", *, date_format: Optional[str] = None) -> None:
    """Append a DATE field; `date_format` is Word's \\@ picture (e.g.
    'MMMM d, yyyy'). The result stays a placeholder until a renderer opens
    the file — this package never computes dates into results."""
    document = _document_of_paragraph(paragraph)
    _refuse_if_protected(document, "insert a field")
    instr = " DATE "
    if date_format:
        escaped = date_format.replace('"', "")
        instr = f' DATE \\@ "{escaped}" '
    with rollback_on_error(document):
        paragraph._p.append(_simple_field(instr, _PLACEHOLDERS["DATE"]))
        _set_update_fields_on_open(document)


def add_reference_field(
    paragraph: "Paragraph", *, bookmark: str, kind: str = "text"
) -> None:
    """Append a cross-reference to `bookmark`: its text (`kind="text"`), its
    page (`"page"`), or its paragraph number (`"number"`)."""
    if kind not in _REFERENCE_KINDS:
        raise ValueError(
            f"kind must be one of {sorted(_REFERENCE_KINDS)}, got {kind!r}"
        )
    document = _document_of_paragraph(paragraph)
    _refuse_if_protected(document, "insert a field")
    from docx.bookmarks import list_bookmarks

    if not any(b.name == bookmark for b in list_bookmarks(document)):
        raise TargetNotFoundError(
            f"no bookmark named {bookmark!r} exists to cross-reference"
        )
    with rollback_on_error(document):
        paragraph._p.append(
            _simple_field(_REFERENCE_KINDS[kind](bookmark), "(reference)")
        )
        _set_update_fields_on_open(document)


def insert_toc_after(
    document: "Document", anchor, *, levels: Tuple[int, int] = (1, 3)
) -> None:
    """Insert a TOC complex field (begin / instrText / separate / placeholder
    / end) in a new paragraph after `anchor`, marked dirty so the renderer
    builds the real table on open. Heading `levels` maps to \\o "1-3"."""
    from docx.blocks import _insert_after, _resolve_anchor_paragraph

    low, high = levels
    if not (1 <= low <= high <= 9):
        raise ValueError(f"levels must satisfy 1 <= low <= high <= 9, got {levels!r}")
    story, anchor_p = _resolve_anchor_paragraph(document, anchor)
    if story != "word/document.xml":
        raise TargetNotFoundError(
            f"a TOC inserts into the main document body (anchor is in {story})"
        )
    from docx.blocks import _refuse_paragraph_in_open_field
    from docx.story import _story_elements

    root = next(r for s, r in _story_elements(document) if s == story)
    _refuse_paragraph_in_open_field(story, root, anchor_p, for_insertion=True)
    with rollback_on_error(document):
        paragraph = OxmlElement("w:p")

        def _fld_char(fld_type: str, *, dirty: bool = False) -> "_Element":
            run = OxmlElement("w:r")
            fld_char = OxmlElement("w:fldChar")
            fld_char.set(qn("w:fldCharType"), fld_type)
            if dirty:
                fld_char.set(qn("w:dirty"), "true")
            run.append(fld_char)
            return run

        instr_run = OxmlElement("w:r")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = f' TOC \\o "{low}-{high}" \\h \\z \\u '
        instr_run.append(instr)
        placeholder_run = OxmlElement("w:r")
        placeholder_run.add_t(
            "Table of contents placeholder — update fields to build it."
        )
        for piece in (
            _fld_char("begin", dirty=True),
            instr_run,
            _fld_char("separate"),
            placeholder_run,
            _fld_char("end"),
        ):
            paragraph.append(piece)
        _insert_after(anchor_p, [paragraph])
        _set_update_fields_on_open(document)
