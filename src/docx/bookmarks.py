"""Bookmark enumeration, creation, and deletion.

Bookmarks are the anchor infrastructure cross-references ride on. Creation
wraps an exact |Span| (the comment-range machinery generalized); ids are
globally unique across every story; deletion removes only the markers —
never the text — and refuses while a field instruction still references the
name (a dangling REF renders "Error!" in Word).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, List, Optional, Tuple

from docx._ownership import require_span_owner
from docx._textatoms import DEL_TEXT, INSTR_TEXT, is_direct_run_child, project_run_child
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


def _paragraph_of(node: "_Element") -> "Optional[_Element]":
    current = node.getparent()
    while current is not None and current.tag != qn("w:p"):
        current = current.getparent()
    return current


def _bookmark_text(root: "_Element", start: "_Element", raw_id: Optional[str]) -> str:
    if raw_id is None:
        return ""
    pieces: "List[str]" = []
    collecting = False
    last_paragraph: "Optional[_Element]" = None
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
        if (
            collecting
            and is_direct_run_child(node)
            and node.tag not in (DEL_TEXT, INSTR_TEXT)
        ):
            projection = project_run_child(node)
            if projection.barrier or not projection.text:
                continue
            paragraph = _paragraph_of(node)
            if last_paragraph is not None and paragraph is not last_paragraph:
                pieces.append("\n")  # boundary, matching Span.text semantics
            last_paragraph = paragraph
            pieces.append(projection.text)
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
    require_span_owner(document, span)
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
    if span.in_field:
        raise UnsupportedStructureError(
            "the span lies inside a field result; a bookmark planted there"
            " is destroyed the next time the field updates — bookmark the"
            " field's source text instead"
        )
    for atom in span._atoms:  # noqa: SLF001
        if atom.run is None:
            raise UnsupportedStructureError(
                "the span includes content outside ordinary runs; bookmark"
                " a plain text range"
            )
    # Allocation parses every existing marker id, so it must happen before
    # edge isolation splits any source runs.
    bookmark_id = _next_bookmark_id(document)
    start_marker = OxmlElement("w:bookmarkStart")
    start_marker.set(_ID, str(bookmark_id))
    start_marker.set(_NAME, name)
    end_marker = OxmlElement("w:bookmarkEnd")
    end_marker.set(_ID, str(bookmark_id))
    span._isolate_edge_runs()  # noqa: SLF001
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
    pairs = []
    for _story, root in _story_elements(document):
        starts_by_id = {}
        ends_by_id = {}
        for tag, destination in (
            (_BOOKMARK_START, starts_by_id),
            (_BOOKMARK_END, ends_by_id),
        ):
            for marker in root.iter(tag):
                raw_id = marker.get(_ID)
                try:
                    marker_id = int(raw_id) if raw_id is not None else -1
                except ValueError:
                    marker_id = -1
                if marker_id < 0 or marker_id in destination:
                    raise UnsupportedStructureError(
                        "bookmark markers have missing, non-numeric, or duplicate"
                        f" w:id {raw_id!r}; nothing was changed"
                    )
                destination[marker_id] = marker
        if set(starts_by_id) != set(ends_by_id):
            raise UnsupportedStructureError(
                "bookmark start/end markers are not one-to-one; nothing was changed"
            )
        pairs.extend(
            (start, ends_by_id[marker_id])
            for marker_id, start in starts_by_id.items()
            if start.get(_NAME) == name
        )
    if not pairs:
        raise TargetNotFoundError(f"no bookmark named {name!r} exists")

    # -- validated; mutate --
    for start, end in pairs:
        end.getparent().remove(end)
        start.getparent().remove(start)


_FLD_CHAR = qn("w:fldChar")
_FLD_CHAR_TYPE = qn("w:fldCharType")


def _iter_field_instructions(root: "_Element"):
    """(concatenated-instruction, [w:instrText nodes]) per complex field —
    Word freely SPLITS one instruction across several runs, so any scan
    matching single nodes silently misses references.
    """
    stack: "List[Tuple[List[_Element], bool]]" = []
    for node in root.iter():
        if node.tag == _FLD_CHAR:
            fld_type = node.get(_FLD_CHAR_TYPE)
            if fld_type == "begin":
                stack.append(([], True))
            elif fld_type == "separate" and stack:
                nodes, _collecting = stack[-1]
                stack[-1] = (nodes, False)
            elif fld_type == "end" and stack:
                nodes, _collecting = stack.pop()
                yield "".join(n.text or "" for n in nodes), nodes
        elif node.tag == _INSTR_TEXT and stack:
            for nodes, collecting in reversed(stack):
                if collecting:
                    nodes.append(node)
                    break


def _reference_pattern(name: str):
    escaped = re.escape(name)
    return re.compile(
        rf'\b(?:PAGEREF|NOTEREF|REF)\s+(?:"{escaped}"|{escaped})(?=\s|$)',
        re.IGNORECASE,
    )


def _field_references(document: "Document", name: str) -> int:
    pattern = _reference_pattern(name)
    count = 0
    for _story, root in _story_elements(document):
        for instruction, _nodes in _iter_field_instructions(root):
            if pattern.search(instruction):
                count += 1
        for node in root.iter(_FLD_SIMPLE):
            instr = node.get(_INSTR)
            if instr and pattern.search(instr):
                count += 1
        for link in root.iter(qn("w:hyperlink")):
            anchor = link.get(qn("w:anchor"))
            if anchor is not None and anchor.casefold() == name.casefold():
                count += 1  # internal links target bookmarks too
    return count
