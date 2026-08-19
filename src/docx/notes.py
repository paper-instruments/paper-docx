"""Create footnotes and endnotes, including the note body and the mark."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx._guard import check_install
from docx._ownership import require_span_owner
from docx._transaction import rollback_on_error
from docx.controls import (
    _control_type,
    _refuse_control_write_restrictions,
    _validate_span_surface_edit,
)
from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.parts.endnotes import EndnotesPart
from docx.parts.footnotes import FootnotesPart
from docx.protection import _refuse_if_protected
from docx.search import _validate_writable_text
from docx.story import _story_elements

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document
    from docx.search import Span

check_install()

_W_DECL = nsdecls("w")
_ID = qn("w:id")
_TYPE = qn("w:type")
_FOOTNOTE = qn("w:footnote")
_ENDNOTE = qn("w:endnote")
_P = qn("w:p")
_R = qn("w:r")
_FOOTNOTE_REFERENCE = qn("w:footnoteReference")
_ENDNOTE_REFERENCE = qn("w:endnoteReference")
_SDT = qn("w:sdt")
_SDT_PR = qn("w:sdtPr")

_FOOTNOTES_TEMPLATE = (
    f"<w:footnotes {_W_DECL}>"
    '<w:footnote w:type="separator" w:id="-1">'
    "<w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
    "</w:footnotes>"
)
_ENDNOTES_TEMPLATE = (
    f"<w:endnotes {_W_DECL}>"
    '<w:endnote w:type="separator" w:id="-1">'
    "<w:p><w:r><w:separator/></w:r></w:p></w:endnote>"
    '<w:endnote w:type="continuationSeparator" w:id="0">'
    "<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>"
    "</w:endnotes>"
)


def _ensure_notes_part(document: "Document", *, endnote: bool):
    reltype = RT.ENDNOTES if endnote else RT.FOOTNOTES
    try:
        return document.part.part_related_by(reltype)
    except KeyError:
        pass
    if endnote:
        part = EndnotesPart(
            PackURI("/word/endnotes.xml"),
            CT.WML_ENDNOTES,
            parse_xml(_ENDNOTES_TEMPLATE),
            document.part.package,
        )
    else:
        part = FootnotesPart(
            PackURI("/word/footnotes.xml"),
            CT.WML_FOOTNOTES,
            parse_xml(_FOOTNOTES_TEMPLATE),
            document.part.package,
        )
    document.part.relate_to(part, reltype)
    return part


def _next_note_id(root: "_Element", tag) -> int:
    used = []
    for note in root.findall(tag):
        if note.get(_TYPE) in ("separator", "continuationSeparator"):
            continue
        raw = note.get(_ID)
        if raw is None:
            continue
        used.append(int(raw))
    return (max(used) if used else 0) + 1


def _insert_note(
    document: "Document",
    span: "Span",
    text: str,
    *,
    endnote: bool,
) -> int:
    operation = "add an endnote" if endnote else "add a footnote"
    require_span_owner(document, span)
    _refuse_if_protected(document, operation)
    if not text:
        raise ValueError("note text must be non-empty")
    _validate_writable_text(text, argument="text")
    span._validate_fresh()  # noqa: SLF001
    main_story = next(
        story
        for story, root in _story_elements(document)
        if root is document.element
    )
    if span.story != main_story:
        raise UnsupportedStructureError(
            "footnotes and endnotes attach in the main document body"
            f" (span is in {span.story})"
        )
    if span.in_text_box:
        raise UnsupportedStructureError(
            "footnotes and endnotes cannot attach inside a text box"
        )
    if span.in_field:
        raise UnsupportedStructureError(
            "the span lies inside a field result; a note planted there"
            " is destroyed the next time the field updates"
        )
    _refuse_control_surface(span)
    with rollback_on_error(document, span):
        span._isolate_edge_runs()  # noqa: SLF001
        runs = []
        for atom in span._atoms:  # noqa: SLF001
            run = atom.run
            if run is not None and not any(existing is run for existing in runs):
                runs.append(run)
        if not runs:
            raise TargetNotFoundError("span has no runs to attach a note to")
        part = _ensure_notes_part(document, endnote=endnote)
        root = part._element  # noqa: SLF001
        tag = _ENDNOTE if endnote else _FOOTNOTE
        note_id = _next_note_id(root, tag)
        kind = "endnote" if endnote else "footnote"
        note = OxmlElement(f"w:{kind}")
        note.set(_ID, str(note_id))
        paragraph = OxmlElement("w:p")
        ref_run = OxmlElement("w:r")
        ref_run.append(OxmlElement(f"w:{kind}Ref"))
        text_run = OxmlElement("w:r")
        text_run.add_t(" " + text)
        paragraph.append(ref_run)
        paragraph.append(text_run)
        note.append(paragraph)
        root.append(note)
        mark = OxmlElement("w:r")
        r_pr = OxmlElement("w:rPr")
        align = OxmlElement("w:vertAlign")
        align.set(qn("w:val"), "superscript")
        r_pr.append(align)
        mark.append(r_pr)
        reference = OxmlElement(f"w:{kind}Reference")
        reference.set(_ID, str(note_id))
        mark.append(reference)
        _insert_after_span(runs, mark)
        return note_id


def _refuse_control_surface(span: "Span") -> None:
    controls = []
    for atom in span._atoms:  # noqa: SLF001
        current = atom.element.getparent()
        while current is not None:
            if current.tag == _SDT and not any(current is item for item in controls):
                controls.append(current)
            current = current.getparent()
    for control in controls:
        _validate_span_surface_edit(control)
        _refuse_control_write_restrictions(control)
        if _control_type(control.find(_SDT_PR)) == "text":
            raise UnsupportedStructureError(
                "span lies in a plain-text content control; a note mark would"
                " break the control's content model. Nothing was changed"
            )


def _is_note_mark(element: "_Element") -> bool:
    if element.tag != _R:
        return False
    return any(
        child.tag in (_FOOTNOTE_REFERENCE, _ENDNOTE_REFERENCE) for child in element
    )


def _insert_after_span(runs, mark: "_Element") -> None:
    anchor = runs[-1]
    sibling = anchor.getnext()
    while sibling is not None and _is_note_mark(sibling):
        anchor = sibling
        sibling = sibling.getnext()
    anchor.addnext(mark)


def add_footnote(document: "Document", span: "Span", text: str) -> int:
    """Insert a footnote mark after `span` and create the note body."""
    return _insert_note(document, span, text, endnote=False)


def add_endnote(document: "Document", span: "Span", text: str) -> int:
    """Insert an endnote mark after `span` and create the note body."""
    return _insert_note(document, span, text, endnote=True)
