"""Visibility-complete story traversal and inspection (paper-docx, opt-in).

Standard python-docx traversal (`Document.paragraphs`, `.tables`,
`iter_inner_content()`) is blind to text inside tracked insertions and
deletions, content controls, text boxes, and to entire story parts
(footnotes, endnotes). This module is the *new, explicitly named* perception
layer (CONVENTIONS §1.1): existing traversal semantics are untouched.

Traversal rules (pinned by API-PROPOSAL.md §4):

* Every story part is walked: body, headers, footers, footnotes, endnotes,
  comments. Separator/continuation-separator footnotes and endnotes are
  plumbing, not content, and are skipped.
* Paragraphs inside table cells are not emitted as separate blocks — their
  text belongs to the table block. Paragraphs inside content controls and
  text boxes ARE emitted, flagged.
* `mc:AlternateContent` contributes only its first `mc:Choice`; fallbacks
  (which duplicate content, e.g. VML copies of text boxes) are skipped, so
  one visible text box yields its text exactly once.
* Empty paragraphs are emitted — block indices must be stable, and an empty
  paragraph is a real edit target.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from typing import TYPE_CHECKING, Dict, Iterator, List, Optional, Tuple

from docx._normalize import normalize_text
from docx.oxml.ns import qn

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

VIEWS = ("current", "original", "all")

_P = qn("w:p")
_TBL = qn("w:tbl")
_SDT = qn("w:sdt")
_SDT_CONTENT = qn("w:sdtContent")
_INS = qn("w:ins")
_DEL = qn("w:del")
_T = qn("w:t")
_DEL_TEXT = qn("w:delText")
_TAB = qn("w:tab")
_BR = qn("w:br")
_CR = qn("w:cr")
_TXBX = qn("w:txbxContent")
_P_STYLE_XPATH = "./w:pPr/w:pStyle/@w:val"
_FOOTNOTE_TYPE = qn("w:type")

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_ALTERNATE = f"{{{_MC_NS}}}AlternateContent"
_MC_CHOICE = f"{{{_MC_NS}}}Choice"

def _story_sort_key(name: str) -> Tuple[int, str]:
    """Traversal order: body, headers, footers, footnotes, endnotes, comments."""
    if name == "word/document.xml":
        return (0, name)
    if name.startswith("word/header"):
        return (1, name)
    if name.startswith("word/footer"):
        return (2, name)
    if name == "word/footnotes.xml":
        return (3, name)
    if name == "word/endnotes.xml":
        return (4, name)
    return (5, name)  # word/comments.xml


@dataclass(frozen=True)
class Anchor:
    """Stable block address: story part + index + content hash (§2, pinned).

    The hash (first 8 hex chars of SHA-256 over the block's normalized text)
    is what detects staleness — a raw index alone is forbidden as a public
    anchor because it goes stale across edits.
    """

    story: str
    index: int
    content_hash: str

    def to_dict(self) -> dict:
        return {"story": self.story, "index": self.index, "content_hash": self.content_hash}


@dataclass(frozen=True)
class TableShape:
    rows: int
    columns: int
    has_merges: bool
    has_nested_table: bool

    def to_dict(self) -> dict:
        return {
            "rows": self.rows,
            "columns": self.columns,
            "has_merges": self.has_merges,
            "has_nested_table": self.has_nested_table,
        }


@dataclass(frozen=True)
class Block:
    """One block-level item (paragraph or table) somewhere in the document."""

    story: str
    kind: str  # "paragraph" | "table"
    index: int
    anchor: Anchor
    text: str
    style_id: Optional[str]
    in_insert: bool
    in_delete: bool
    in_content_control: bool
    in_text_box: bool
    table: Optional[TableShape]

    def to_dict(self) -> dict:
        return {
            "story": self.story,
            "kind": self.kind,
            "index": self.index,
            "anchor": self.anchor.to_dict(),
            "text": self.text,
            "style_id": self.style_id,
            "in_insert": self.in_insert,
            "in_delete": self.in_delete,
            "in_content_control": self.in_content_control,
            "in_text_box": self.in_text_box,
            "table": self.table.to_dict() if self.table else None,
        }


@dataclass(frozen=True)
class Outline:
    """Inspection snapshot of a document: every block in every story part."""

    story_parts: Tuple[str, ...]
    blocks: Tuple[Block, ...]
    blind_region_counts: Dict[str, int]

    def to_dict(self) -> dict:
        return {
            "schema": "paper_outline",
            "version": 1,
            "story_parts": list(self.story_parts),
            "blind_region_counts": dict(sorted(self.blind_region_counts.items())),
            "blocks": [block.to_dict() for block in self.blocks],
        }


def content_hash(text: str) -> str:
    """First 8 hex chars of SHA-256 over the block's normalized text."""
    return hashlib.sha256(normalize_text(text).encode("utf-8")).hexdigest()[:8]


def _story_elements(document: "Document") -> "List[Tuple[str, _Element]]":
    """(story-name, root element) for every story part, deterministic order."""
    package = document.part.package
    assert package is not None
    found: "List[Tuple[str, _Element]]" = []
    for part in package.iter_parts():
        name = str(part.partname).lstrip("/")
        element = getattr(part, "_element", None)
        if element is None:
            continue
        if name == "word/document.xml" or (
            name.startswith(("word/header", "word/footer")) and name.endswith(".xml")
        ) or name in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
            found.append((name, element))
    return sorted(found, key=lambda item: _story_sort_key(item[0]))


def story_parts(document: "Document") -> Tuple[str, ...]:
    """Every story part present in `document`, in traversal order."""
    return tuple(name for name, _ in _story_elements(document))


def _first_choice_children(element: "_Element") -> "List[_Element]":
    """`element`'s children with mc:AlternateContent collapsed to its first
    mc:Choice (fallback content is a duplicate and must not be traversed)."""
    result: "List[_Element]" = []
    for child in element:
        if child.tag == _MC_ALTERNATE:
            for alt_child in child:
                if alt_child.tag == _MC_CHOICE:
                    result.extend(_first_choice_children(alt_child))
                    break
        else:
            result.append(child)
    return result


class _TextVisitor:
    """Accumulates view-filtered text and region flags over a subtree."""

    def __init__(self, view: str) -> None:
        self.view = view
        self.pieces: List[str] = []
        self.in_insert = False
        self.in_delete = False
        self.in_content_control = False
        self.in_text_box = False

    def visit(self, element: "_Element", *, in_ins: bool, in_del: bool,
              in_sdt: bool, in_txbx: bool, skip_text_boxes: bool) -> None:
        tag = element.tag
        if tag == _TXBX and skip_text_boxes:
            return
        if tag == _INS:
            in_ins = True
        elif tag == _DEL:
            in_del = True
        elif tag == _SDT:
            in_sdt = True
        elif tag == _TXBX:
            in_txbx = True

        if tag == _T:
            if self.view == "original" and in_ins:
                return
            self._emit(element.text or "", in_ins, in_del, in_sdt, in_txbx)
            return
        if tag == _DEL_TEXT:
            if self.view == "current":
                return
            if self.view == "original" and in_ins:
                # a deletion nested inside a pending insertion never existed
                # in the original document
                return
            self._emit(element.text or "", in_ins, True, in_sdt, in_txbx)
            return
        if tag == _TAB:
            self._emit("\t", in_ins, in_del, in_sdt, in_txbx)
            return
        if tag in (_BR, _CR):
            self._emit("\n", in_ins, in_del, in_sdt, in_txbx)
            return
        for child in _first_choice_children(element):
            self.visit(child, in_ins=in_ins, in_del=in_del, in_sdt=in_sdt,
                       in_txbx=in_txbx, skip_text_boxes=skip_text_boxes)

    def _emit(self, text: str, in_ins: bool, in_del: bool, in_sdt: bool,
              in_txbx: bool) -> None:
        self.pieces.append(text)
        self.in_insert = self.in_insert or in_ins
        self.in_delete = self.in_delete or in_del
        self.in_content_control = self.in_content_control or in_sdt
        self.in_text_box = self.in_text_box or in_txbx

    @property
    def text(self) -> str:
        return "".join(self.pieces)


def _subtree_text(element: "_Element", view: str, *, skip_text_boxes: bool,
                  in_sdt: bool = False, in_txbx: bool = False) -> _TextVisitor:
    visitor = _TextVisitor(view)
    for child in _first_choice_children(element):
        visitor.visit(child, in_ins=False, in_del=False, in_sdt=in_sdt,
                      in_txbx=in_txbx, skip_text_boxes=skip_text_boxes)
    return visitor


def _table_shape(table: "_Element") -> TableShape:
    rows = table.findall(qn("w:tr"))
    columns = max((len(row.findall(qn("w:tc"))) for row in rows), default=0)
    has_merges = bool(
        table.findall(f".//{qn('w:vMerge')}") or table.findall(f".//{qn('w:gridSpan')}")
    )
    has_nested = any(t is not table for t in table.iter(_TBL))
    return TableShape(
        rows=len(rows), columns=columns, has_merges=has_merges, has_nested_table=has_nested
    )


def _text_box_contents(paragraph: "_Element") -> "List[_Element]":
    """w:txbxContent elements reachable from `paragraph`, fallbacks excluded."""
    found: "List[_Element]" = []

    def walk(element: "_Element") -> None:
        for child in _first_choice_children(element):
            if child.tag == _TXBX:
                found.append(child)
            else:
                walk(child)

    walk(paragraph)
    return found


def _block_containers(story: str, root: "_Element") -> "Iterator[Tuple[_Element, bool]]":
    """(container, in_content_control) holders of block-level content."""
    tag = root.tag
    if tag == qn("w:document"):
        body = root.find(qn("w:body"))
        if body is not None:
            yield body, False
    elif tag in (qn("w:hdr"), qn("w:ftr")):
        yield root, False
    elif tag in (qn("w:footnotes"), qn("w:endnotes")):
        for note in root:
            if note.tag not in (qn("w:footnote"), qn("w:endnote")):
                continue
            note_type = note.get(_FOOTNOTE_TYPE)
            if note_type in ("separator", "continuationSeparator"):
                continue  # plumbing, not content
            yield note, False
    elif tag == qn("w:comments"):
        for comment in root:
            if comment.tag == qn("w:comment"):
                yield comment, False
    else:  # pragma: no cover - unknown story roots are a programming error
        raise ValueError(f"unrecognized story root {tag!r} in {story}")


def _walk_container(
    container: "_Element",
    counter: List[int],
    *,
    in_sdt: bool,
    in_txbx: bool,
) -> "Iterator[Tuple[str, int, _Element, bool, bool]]":
    """(kind, block-index, element, in_sdt, in_txbx) for each block, in order.

    THE single definition of block identity and indexing — `iter_blocks` and
    `docx.search` both ride on it, so a span's block anchor can never disagree
    with the outline's.
    """
    for child in _first_choice_children(container):
        if child.tag == _P:
            index = counter[0]
            counter[0] += 1
            yield ("paragraph", index, child, in_sdt, in_txbx)
            for txbx in _text_box_contents(child):
                yield from _walk_container(txbx, counter, in_sdt=in_sdt, in_txbx=True)
        elif child.tag == _TBL:
            index = counter[0]
            counter[0] += 1
            yield ("table", index, child, in_sdt, in_txbx)
        elif child.tag == _SDT:
            content = child.find(_SDT_CONTENT)
            if content is not None:
                yield from _walk_container(content, counter, in_sdt=True, in_txbx=in_txbx)


def _iter_block_elements(
    story: str, root: "_Element"
) -> "Iterator[Tuple[str, int, _Element, bool, bool]]":
    counter = [0]
    for container, in_sdt in _block_containers(story, root):
        yield from _walk_container(container, counter, in_sdt=in_sdt, in_txbx=False)


def _build_block(
    story: str,
    kind: str,
    index: int,
    element: "_Element",
    view: str,
    *,
    in_sdt: bool,
    in_txbx: bool,
) -> Block:
    if kind == "table":
        visitor = _subtree_text(element, view, skip_text_boxes=False,
                                in_sdt=in_sdt, in_txbx=in_txbx)
        text = _table_text(element, view)
        style_id = None
        table = _table_shape(element)
    else:
        visitor = _subtree_text(element, view, skip_text_boxes=True,
                                in_sdt=in_sdt, in_txbx=in_txbx)
        text = visitor.text
        style_values = element.xpath(_P_STYLE_XPATH)
        style_id = str(style_values[0]) if style_values else None
        table = None
    return Block(
        story=story,
        kind=kind,
        index=index,
        anchor=Anchor(story=story, index=index, content_hash=content_hash(text)),
        text=text,
        style_id=style_id,
        in_insert=visitor.in_insert,
        in_delete=visitor.in_delete,
        in_content_control=in_sdt or visitor.in_content_control,
        in_text_box=in_txbx or visitor.in_text_box,
        table=table,
    )


def _table_text(table: "_Element", view: str) -> str:
    """Cell texts in row-major order, newline-joined (the table block owns
    all text inside it, nested content included)."""
    pieces: List[str] = []
    for row in table.findall(qn("w:tr")):
        for cell in row.findall(qn("w:tc")):
            visitor = _subtree_text(cell, view, skip_text_boxes=False)
            pieces.append(visitor.text)
    return "\n".join(pieces)


def _count_blind_regions(root: "_Element") -> Dict[str, int]:
    """Occurrences of each blind-region element in traversal space (i.e.,
    with mc:Fallback duplicates excluded)."""
    counts = {
        "tracked_insertions": 0,
        "tracked_deletions": 0,
        "content_controls": 0,
        "text_boxes": 0,
    }
    tag_keys = {
        _INS: "tracked_insertions",
        _DEL: "tracked_deletions",
        _SDT: "content_controls",
        _TXBX: "text_boxes",
    }

    def walk(element: "_Element") -> None:
        for child in _first_choice_children(element):
            key = tag_keys.get(child.tag)
            if key is not None:
                counts[key] += 1
            walk(child)

    walk(root)
    return counts


def iter_blocks(document: "Document", *, view: str = "current") -> Iterator[Block]:
    """Every block in every story part of `document`, in document order.

    `view` selects the text layer: "current" (insertions in, deletions out —
    the document if all changes were accepted), "original" (deletions in,
    insertions out), or "all" (everything).
    """
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}, got {view!r}")
    for story, root in _story_elements(document):
        for kind, index, element, in_sdt, in_txbx in _iter_block_elements(story, root):
            yield _build_block(
                story, kind, index, element, view, in_sdt=in_sdt, in_txbx=in_txbx
            )


def outline(document: "Document", *, view: str = "current") -> Outline:
    """Inspection snapshot: story parts, all blocks, blind-region counts.

    Deterministic: the same document yields byte-identical `to_dict()` output
    on every call (CONVENTIONS §4, inspection determinism).
    """
    blocks = tuple(iter_blocks(document, view=view))
    totals = {
        "tracked_insertions": 0,
        "tracked_deletions": 0,
        "content_controls": 0,
        "text_boxes": 0,
    }
    for _, root in _story_elements(document):
        for key, value in _count_blind_regions(root).items():
            totals[key] += value
    return Outline(
        story_parts=story_parts(document),
        blocks=blocks,
        blind_region_counts=totals,
    )
