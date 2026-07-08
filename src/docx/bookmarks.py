"""Bookmark enumeration, creation, and deletion (v0.11 Phase 6).

Bookmarks are the anchor infrastructure cross-references ride on. Creation
wraps an exact |Span| (the comment-range machinery generalized); ids are
globally unique across every story; deletion removes only the markers —
never the text — and refuses while a field instruction still references the
name (a dangling REF renders "Error!" in Word).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, List, Optional

from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import _refuse_if_protected
from docx.story import _story_elements

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document
    from docx.search import Span

_BOOKMARK_START = qn("w:bookmarkStart")
_BOOKMARK_END = qn("w:bookmarkEnd")
_ID = qn("w:id")
_NAME = qn("w:name")
_T = qn("w:t")
_INSTR_TEXT = qn("w:instrText")
_FLD_SIMPLE = qn("w:fldSimple")
_INSTR = qn("w:instr")

#: Word's bookmark-name rules: start with a letter, then letters/digits/
#: underscores, max 40 chars ("_"-prefixed names are Word-internal)
_NAME_RE = re.compile(r"^[^\W\d_][\w]{0,39}$", re.UNICODE)


@dataclass(frozen=True)
class BookmarkInfo:
    """One bookmark: its name, id, story, and the visible text it wraps."""

    name: str
    bookmark_id: int
    story: str
    text: str

    @property
    def is_point(self) -> bool:
        return self.text == ""

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "id": self.bookmark_id,
            "story": self.story,
            "text": self.text,
            "is_point": self.is_point,
        }


def list_bookmarks(document: "Document") -> "List[BookmarkInfo]":
    """Every bookmark across every story, in document order."""
    found: "List[BookmarkInfo]" = []
    for story, root in _story_elements(document):
        for start in root.iter(_BOOKMARK_START):
            name = start.get(_NAME) or ""
            raw_id = start.get(_ID)
            bookmark_id = int(raw_id) if raw_id is not None else -1
            found.append(
                BookmarkInfo(
                    name=name,
                    bookmark_id=bookmark_id,
                    story=story,
                    text=_bookmark_text(root, start, raw_id),
                )
            )
    return found


def _bookmark_text(root: "_Element", start: "_Element", raw_id: Optional[str]) -> str:
    if raw_id is None:
        return ""
    pieces: "List[str]" = []
    collecting = False
    for node in root.iter():
        if node is start:
            collecting = True
            continue
        if (
            collecting
            and node.tag == _BOOKMARK_END
            and node.get(_ID) == raw_id
        ):
            break
        if collecting and node.tag == _T:
            pieces.append(node.text or "")
    return "".join(pieces)


def _next_bookmark_id(document: "Document") -> int:
    highest = 0
    for _story, root in _story_elements(document):
        for node in root.iter(_BOOKMARK_START, _BOOKMARK_END):
            raw = node.get(_ID)
            if raw is not None:
                highest = max(highest, int(raw))
    return highest + 1


def create_bookmark(document: "Document", span: "Span", name: str) -> BookmarkInfo:
    """Bookmark exactly `span`'s text under `name` (unique, Word-legal).

    Boundary runs are split so the markers wrap precisely the span's
    characters — the same isolation the comment-range machinery uses.
    """
    _refuse_if_protected(document, "create a bookmark")
    if not _NAME_RE.match(name or ""):
        raise ValueError(
            f"bookmark name {name!r} is not Word-legal (start with a letter;"
            " letters, digits and underscores only; max 40 chars)"
        )
    if any(b.name == name for b in list_bookmarks(document)):
        raise UnsupportedStructureError(
            f"a bookmark named {name!r} already exists; bookmark names are"
            " document-unique"
        )
    span._validate_fresh()  # noqa: SLF001 - same-package machinery
    for atom in span._atoms:  # noqa: SLF001
        if atom.run is None:
            raise UnsupportedStructureError(
                "the span includes content outside ordinary runs; bookmark"
                " a plain text range"
            )
    span._isolate_edge_runs()  # noqa: SLF001
    bookmark_id = _next_bookmark_id(document)
    start_marker = OxmlElement("w:bookmarkStart")
    start_marker.set(_ID, str(bookmark_id))
    start_marker.set(_NAME, name)
    end_marker = OxmlElement("w:bookmarkEnd")
    end_marker.set(_ID, str(bookmark_id))
    first_run = span._atoms[0].run  # noqa: SLF001
    last_run = span._atoms[-1].run  # noqa: SLF001
    first_run.addprevious(start_marker)
    last_run.addnext(end_marker)
    return BookmarkInfo(
        name=name, bookmark_id=bookmark_id, story=span.story, text=span.text
    )


def delete_bookmark(document: "Document", name: str) -> None:
    """Remove `name`'s markers (the text stays). Refuses while any field
    instruction still references the name — a dangling REF/PAGEREF renders
    'Error! Reference source not found.' in Word."""
    _refuse_if_protected(document, "delete a bookmark")
    referencing = _field_references(document, name)
    if referencing:
        raise UnsupportedStructureError(
            f"bookmark {name!r} is referenced by {referencing} field"
            " instruction(s); update or remove those fields first"
        )
    for _story, root in _story_elements(document):
        for start in root.iter(_BOOKMARK_START):
            if start.get(_NAME) != name:
                continue
            raw_id = start.get(_ID)
            for end in root.iter(_BOOKMARK_END):
                if end.get(_ID) == raw_id:
                    end.getparent().remove(end)
                    break
            start.getparent().remove(start)
            return
    raise TargetNotFoundError(f"no bookmark named {name!r} exists")


def _field_references(document: "Document", name: str) -> int:
    pattern = re.compile(rf"\b(?:PAGE)?REF\s+{re.escape(name)}(?=\s|$)")
    count = 0
    for _story, root in _story_elements(document):
        for node in root.iter(_INSTR_TEXT):
            if node.text and pattern.search(node.text):
                count += 1
        for node in root.iter(_FLD_SIMPLE):
            instr = node.get(_INSTR)
            if instr and pattern.search(instr):
                count += 1
    return count
