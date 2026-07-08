"""Anchor-relative block operations and clause-level redlines (paper-docx).

Operations here compose whole paragraphs: insert a section after an anchor,
tracked-delete a paragraph range, tracked-replace a range. The load-bearing
safety rule (kept from the battle-tested reference): every selected paragraph
must share one parent element — selections that would span story regions,
table boundaries, or content-control boundaries are refused loudly.

Improvements over the reference helpers (documented in API-PROPOSAL.md §7):

* Tracked deletion moves the paragraph's own runs into `w:del`, retagging
  `w:t` -> `w:delText` in place, so every run keeps its `rPr` and a later
  reject restores the original exactly (the reference collapsed formatting
  into one run).
* Tracked insert/delete also stamp the paragraph MARK (`w:pPr/w:rPr/w:ins`
  or `/w:del`), so accepting a deletion removes the paragraph instead of
  leaving an empty husk, and rejecting an insertion removes the inserted
  paragraph entirely.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import TYPE_CHECKING, List, Optional, Sequence, Tuple, Union

from docx import _clock
from docx.errors import (
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.revision import CT_RunTrackChange
from docx.search import Span, find_one
from docx.story import Anchor, Block, _iter_block_elements, _story_elements

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

AnchorLike = Union[str, Block, Span, Anchor]

_P = qn("w:p")
_PPR = qn("w:pPr")
_RPR = qn("w:rPr")
_R = qn("w:r")
_T = qn("w:t")
_TAB = qn("w:tab")
_BR = qn("w:br")
_CR = qn("w:cr")
_DEL_TEXT = qn("w:delText")

#: run children a tracked paragraph delete knows how to handle safely
_SAFE_RUN_CHILDREN = frozenset({_RPR, _T, _TAB, _BR, _CR})


@dataclass(frozen=True)
class BlockEditResult:
    """Outcome of a block-level edit."""

    story: str
    inserted_blocks: int
    deleted_blocks: int
    deleted_text: Tuple[str, ...]
    revision_ids: Tuple[int, ...]

    def to_dict(self) -> dict:
        return {
            "schema": "paper_block_edit",
            "version": 1,
            "story": self.story,
            "inserted_blocks": self.inserted_blocks,
            "deleted_blocks": self.deleted_blocks,
            "deleted_text": list(self.deleted_text),
            "revision_ids": list(self.revision_ids),
        }


# ---------------------------------------------------------------------------
# anchor resolution
# ---------------------------------------------------------------------------


def _resolve_anchor_paragraph(
    document: "Document", anchor: AnchorLike
) -> "Tuple[str, _Element]":
    """(story, paragraph element) for `anchor`, staleness-verified.

    Strings are found via `find_one` (ambiguity refuses); Block/Anchor values
    are re-verified by content hash against the current-view text of the
    block at their recorded position.
    """
    if isinstance(anchor, str):
        span = find_one(document, anchor)
        paragraph = span._atoms[0].paragraph  # noqa: SLF001 - same-package access
        if paragraph is None:
            raise TargetNotFoundError(f"anchor text {anchor!r} is not inside a paragraph")
        return span.story, paragraph
    if isinstance(anchor, Span):
        anchor._validate_fresh()  # noqa: SLF001 - same-package access
        paragraph = anchor._atoms[0].paragraph  # noqa: SLF001
        if paragraph is None:
            raise TargetNotFoundError("span anchor is not inside a paragraph")
        return anchor.story, paragraph
    block_anchor = anchor.anchor if isinstance(anchor, Block) else anchor
    for story, root in _story_elements(document):
        if story != block_anchor.story:
            continue
        for kind, index, element, _sdt, _txbx in _iter_block_elements(story, root):
            if index != block_anchor.index:
                continue
            if kind != "paragraph":
                raise UnsupportedStructureError(
                    "anchor addresses a table block; block operations anchor"
                    " on paragraphs"
                )
            from docx.story import _build_block

            block = _build_block(
                story, kind, index, element, "current", in_sdt=_sdt, in_txbx=_txbx
            )
            if block.anchor.content_hash != block_anchor.content_hash:
                raise TargetNotFoundError(
                    f"anchor is stale: block {block_anchor.index} in"
                    f" {block_anchor.story} no longer carries the anchored content"
                )
            return story, element
        raise TargetNotFoundError(
            f"anchor index {block_anchor.index} does not exist in {block_anchor.story}"
        )
    raise TargetNotFoundError(f"story part {block_anchor.story!r} not found")


def _validated_style_id(
    document: "Document", style_id: Optional[str], *, argument: str
) -> Optional[str]:
    if style_id is None:
        return None
    defined = {style.style_id for style in document.styles}
    if style_id not in defined:
        raise TargetNotFoundError(
            f"{argument} {style_id!r} is not defined in this document's styles"
        )
    return style_id


# ---------------------------------------------------------------------------
# paragraph construction / revision stamping (all through the oxml layer)
# ---------------------------------------------------------------------------


def _new_paragraph(text: str, style_id: Optional[str]) -> "_Element":
    paragraph = OxmlElement("w:p")
    if style_id:
        paragraph.style = style_id  # CT_P descriptor path, ordering handled
    run = OxmlElement("w:r")
    run.add_t(text)
    paragraph.append(run)
    return paragraph


def _stamp_paragraph_mark(
    paragraph: "_Element", tag: str, revision_id: int, author: str, stamp: dt.datetime
) -> None:
    """Mark the paragraph MARK inserted/deleted via `w:pPr/w:rPr/<tag>`."""
    p_pr = paragraph.get_or_add_pPr()
    r_pr = p_pr.find(_RPR)
    if r_pr is None:
        r_pr = OxmlElement("w:rPr")
        # w:rPr's successors in the pPr child sequence (CT_PPr._tag_seq)
        p_pr.insert_element_before(r_pr, "w:sectPr", "w:pPrChange")
    # CT_ParaRPr's schema puts w:ins/w:del FIRST, before any run properties
    r_pr.insert(0, CT_RunTrackChange.new(tag, revision_id, author, stamp))


def _wrap_paragraph_content_as_insertion(
    paragraph: "_Element", revision_id: int, author: str, stamp: dt.datetime
) -> None:
    ins = CT_RunTrackChange.new("w:ins", revision_id, author, stamp)
    for child in list(paragraph):
        if child.tag != _PPR:
            paragraph.remove(child)
            ins.append(child)
    paragraph.append(ins)


def _next_revision_id(document: "Document") -> int:
    from docx.search import _next_revision_id as shared

    return shared(document)


def _validate_deletable_paragraph(paragraph: "_Element") -> None:
    for child in paragraph:
        if child.tag == _PPR:
            continue
        if child.tag != _R:
            raise UnsupportedStructureError(
                "tracked paragraph delete supports plain runs only; found"
                f" {child.tag.rsplit('}', 1)[-1]!r} content (hyperlinks,"
                " controls and existing revisions are refused, not corrupted)"
            )
        for run_child in child:
            if run_child.tag not in _SAFE_RUN_CHILDREN:
                raise UnsupportedStructureError(
                    "tracked paragraph delete supports text-only runs; found"
                    f" {run_child.tag.rsplit('}', 1)[-1]!r} inside a run"
                )


def _paragraph_visible_text(paragraph: "_Element") -> str:
    pieces: List[str] = []
    for node in paragraph.iter():
        if node.tag == _T:
            pieces.append(node.text or "")
        elif node.tag == _TAB:
            pieces.append("\t")
        elif node.tag in (_BR, _CR):
            pieces.append("\n")
    return "".join(pieces)


def _mark_paragraph_deleted(
    paragraph: "_Element", revision_id: int, author: str, stamp: dt.datetime
) -> str:
    """Move the paragraph's runs into `w:del`, preserving each run's `rPr`."""
    text = _paragraph_visible_text(paragraph)
    deletion = CT_RunTrackChange.new("w:del", revision_id, author, stamp)
    for child in list(paragraph):
        if child.tag == _PPR:
            continue
        paragraph.remove(child)
        for t_elm in child.iter(_T):
            t_elm.tag = _DEL_TEXT
        deletion.append(child)
    paragraph.append(deletion)
    _stamp_paragraph_mark(paragraph, "w:del", revision_id, author, stamp)
    return text


# ---------------------------------------------------------------------------
# paragraph-range selection (same-parent rule)
# ---------------------------------------------------------------------------


def _select_paragraph_range(
    document: "Document",
    start_anchor: AnchorLike,
    end_anchor: Optional[AnchorLike],
    count: int,
) -> "Tuple[str, List[_Element]]":
    if count < 1:
        raise ValueError("count must be >= 1")
    story, start_p = _resolve_anchor_paragraph(document, start_anchor)
    # ranges are counted among the start paragraph's SIBLINGS: nested
    # paragraphs (table cells, text boxes) never silently join a range, and
    # the same-parent safety rule holds by construction
    parent = start_p.getparent()
    siblings = [child for child in parent if child.tag == _P]
    start_index = next(i for i, p in enumerate(siblings) if p is start_p)
    if end_anchor is not None:
        end_story, end_p = _resolve_anchor_paragraph(document, end_anchor)
        if end_story != story:
            raise BoundaryViolationError(
                "start and end anchors live in different story parts"
            )
        if end_p.getparent() is not parent:
            raise BoundaryViolationError(
                "start and end anchors do not share one parent; edits spanning"
                " story regions, table boundaries, or content-control"
                " boundaries are refused"
            )
        end_index = next(i for i, p in enumerate(siblings) if p is end_p)
        if end_index < start_index:
            raise TargetNotFoundError("end anchor appears before start anchor")
    else:
        end_index = start_index + count - 1
        if end_index >= len(siblings):
            raise TargetNotFoundError(
                f"count={count} extends past the last paragraph of {story}"
            )
    return story, siblings[start_index : end_index + 1]


def _insert_after(anchor: "_Element", nodes: "Sequence[_Element]") -> None:
    point = anchor
    for node in nodes:
        point.addnext(node)
        point = node


# ---------------------------------------------------------------------------
# public operations
# ---------------------------------------------------------------------------


def insert_section_after(
    document: "Document",
    anchor: AnchorLike,
    *,
    heading: str,
    paragraphs: Sequence[str],
    heading_style: str = "Heading2",
    body_style: Optional[str] = None,
    tracked: bool = False,
    author: Optional[str] = None,
    date: Optional[dt.datetime] = None,
) -> BlockEditResult:
    """Insert a heading plus body paragraphs after `anchor`.

    With `tracked=True` every inserted paragraph is a real Word insertion
    (content wrapped in `w:ins`, paragraph mark stamped) attributed to
    `author` (required) and `date` (default: the injectable clock).
    """
    if tracked and not author:
        raise ValueError("author is required when tracked=True")
    heading_style_id = _validated_style_id(document, heading_style, argument="heading_style")
    body_style_id = _validated_style_id(document, body_style, argument="body_style")
    story, anchor_p = _resolve_anchor_paragraph(document, anchor)

    # -- validated; build and mutate --
    stamp = date if date is not None else _clock.now()
    nodes = [_new_paragraph(heading, heading_style_id)]
    nodes.extend(_new_paragraph(text, body_style_id) for text in paragraphs)
    revision_ids: List[int] = []
    if tracked:
        next_id = _next_revision_id(document)
        for node in nodes:
            _wrap_paragraph_content_as_insertion(node, next_id, author, stamp)
            _stamp_paragraph_mark(node, "w:ins", next_id, author, stamp)
            revision_ids.append(next_id)
            next_id += 1
    _insert_after(anchor_p, nodes)
    return BlockEditResult(
        story=story,
        inserted_blocks=len(nodes),
        deleted_blocks=0,
        deleted_text=(),
        revision_ids=tuple(revision_ids),
    )


def tracked_delete_paragraphs(
    document: "Document",
    start_anchor: AnchorLike,
    *,
    end_anchor: Optional[AnchorLike] = None,
    count: int = 1,
    author: str,
    date: Optional[dt.datetime] = None,
) -> BlockEditResult:
    """Mark a paragraph range deleted with real tracked changes.

    Each paragraph's runs move into `w:del` with their formatting intact and
    the paragraph mark is stamped deleted, so accept removes the paragraphs
    and reject restores them exactly.
    """
    if not author:
        raise ValueError("author is required")
    story, selected = _select_paragraph_range(document, start_anchor, end_anchor, count)
    for paragraph in selected:
        _validate_deletable_paragraph(paragraph)

    # -- validated; mutate --
    stamp = date if date is not None else _clock.now()
    next_id = _next_revision_id(document)
    deleted_text: List[str] = []
    revision_ids: List[int] = []
    for paragraph in selected:
        deleted_text.append(_mark_paragraph_deleted(paragraph, next_id, author, stamp))
        revision_ids.append(next_id)
        next_id += 1
    return BlockEditResult(
        story=story,
        inserted_blocks=0,
        deleted_blocks=len(selected),
        deleted_text=tuple(deleted_text),
        revision_ids=tuple(revision_ids),
    )


def tracked_replace_paragraphs(
    document: "Document",
    start_anchor: AnchorLike,
    replacement_paragraphs: Sequence[str],
    *,
    end_anchor: Optional[AnchorLike] = None,
    count: int = 1,
    body_style: Optional[str] = None,
    author: str,
    date: Optional[dt.datetime] = None,
) -> BlockEditResult:
    """Tracked-delete a paragraph range and tracked-insert replacements after it."""
    if not author:
        raise ValueError("author is required")
    body_style_id = _validated_style_id(document, body_style, argument="body_style")
    story, selected = _select_paragraph_range(document, start_anchor, end_anchor, count)
    for paragraph in selected:
        _validate_deletable_paragraph(paragraph)

    # -- validated; mutate --
    stamp = date if date is not None else _clock.now()
    next_id = _next_revision_id(document)
    deleted_text: List[str] = []
    revision_ids: List[int] = []
    for paragraph in selected:
        deleted_text.append(_mark_paragraph_deleted(paragraph, next_id, author, stamp))
        revision_ids.append(next_id)
        next_id += 1
    inserted = [_new_paragraph(text, body_style_id) for text in replacement_paragraphs]
    for node in inserted:
        _wrap_paragraph_content_as_insertion(node, next_id, author, stamp)
        _stamp_paragraph_mark(node, "w:ins", next_id, author, stamp)
        revision_ids.append(next_id)
        next_id += 1
    _insert_after(selected[-1], inserted)
    return BlockEditResult(
        story=story,
        inserted_blocks=len(inserted),
        deleted_blocks=len(selected),
        deleted_text=tuple(deleted_text),
        revision_ids=tuple(revision_ids),
    )
