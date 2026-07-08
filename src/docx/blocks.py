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
from docx.protection import _refuse_if_protected
from docx.oxml.revision import CT_RunTrackChange
from docx.search import Span, _validate_writable_text, find_one
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
_SAFE_RUN_CHILDREN = frozenset(
    {_RPR, _T, _TAB, _BR, _CR, qn("w:commentReference"), qn("w:lastRenderedPageBreak")}
)

#: markup Word scatters through virtually every saved document (v0.1 S2):
#: spell/grammar flags, point bookmarks, comment anchors. Tracked block ops
#: treat these as transparent — preserved in place (proofErr dropped: it is
#: transient checker state Word regenerates) instead of refusing the edit.
_TRANSPARENT_PARAGRAPH_CHILDREN = frozenset(
    qn(tag)
    for tag in (
        "w:proofErr",
        "w:bookmarkStart",
        "w:bookmarkEnd",
        "w:commentRangeStart",
        "w:commentRangeEnd",
    )
)
_PROOF_ERR = qn("w:proofErr")


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

    Every block operation resolves its anchor here, so this is also the
    protection choke point (v0.11 Phase 3).
    """
    _refuse_if_protected(document, "insert or remove paragraphs")
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


def _refuse_cell_anchor(paragraph: "_Element") -> None:
    """Section/block insertion targets body-level flow; an anchor resolving
    inside a table cell would silently build the section INSIDE the cell."""
    parent = paragraph.getparent()
    if parent is not None and parent.tag == qn("w:tc"):
        raise UnsupportedStructureError(
            "anchor resolves inside a table cell; block insertion targets"
            " body-level paragraphs — anchor on a paragraph outside the table"
        )


def _field_open_flags(story: str, root: "_Element", paragraph: "_Element"):
    """(open_before, open_after) complex-field state around `paragraph`'s block.

    A multi-paragraph field (every Word TOC) keeps begin..end open across
    blocks; block operations inside that region would write content Word
    erases on the next field update (v0.1 H4 at block level).
    """
    from docx.story import _count_fldchar_delta

    depth = 0
    for _kind, _index, element, _sdt, _txbx in _iter_block_elements(story, root):
        contains = element is paragraph or any(
            node is paragraph for node in element.iter(_P)
        )
        delta = _count_fldchar_delta(element)
        if contains:
            return depth > 0, (depth + delta) > 0
        depth = max(0, depth + delta)
    return False, False


def _refuse_paragraph_in_open_field(
    story: str, root: "_Element", paragraph: "_Element", *, for_insertion: bool
) -> None:
    open_before, open_after = _field_open_flags(story, root, paragraph)
    blocked = open_after if for_insertion else (open_before or open_after)
    if blocked:
        raise UnsupportedStructureError(
            "target lies inside a field result that spans paragraphs (a TOC or"
            " similar); Word regenerates field results on update, so content"
            " written there would silently vanish"
        )


def _named_bookmarks_in(paragraph: "_Element"):
    """Names of non-point, non-_GoBack bookmarks with a marker in `paragraph`."""
    starts = {}
    named = []
    stream = list(paragraph.iter())
    for position, node in enumerate(stream):
        if node.tag == qn("w:bookmarkStart"):
            name = node.get(qn("w:name")) or ""
            starts[node.get(qn("w:id"))] = (position, name)
            if name != "_GoBack":
                named.append((node.get(qn("w:id")), name))
        elif node.tag == qn("w:bookmarkEnd"):
            entry = starts.pop(node.get(qn("w:id")), None)
            if entry is None:
                continue
            start_pos, name = entry
            if name == "_GoBack":
                continue
            has_text = any(
                inner.tag == _T and (inner.text or "")
                for inner in stream[start_pos + 1 : position]
            )
            if not has_text:
                named = [(i, n) for i, n in named if i != node.get(qn("w:id"))]
    return [name for _, name in named]


def _validate_deletable_paragraph(paragraph: "_Element") -> None:
    hollowable = _named_bookmarks_in(paragraph)
    if hollowable:
        raise UnsupportedStructureError(
            f"paragraph carries named bookmark(s) {hollowable} — the targets"
            " of REF/PAGEREF/TOC references; deleting it would hollow them."
            " Remove the bookmark deliberately first (only point bookmarks"
            " like _GoBack are transparent)"
        )
    for child in paragraph:
        if child.tag == _PPR or child.tag in _TRANSPARENT_PARAGRAPH_CHILDREN:
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
    """Move the paragraph's runs into `w:del`, preserving each run's `rPr`.

    Transparent markup stays OUTSIDE the deletion: bookmarks and comment
    anchors keep their positions at paragraph level (Word does the same), and
    `w:proofErr` flags are dropped — they are transient spell/grammar state
    Word regenerates on open, and text marked deleted has no checker state.
    """
    text = _paragraph_visible_text(paragraph)
    deletion = CT_RunTrackChange.new("w:del", revision_id, author, stamp)
    placement = None
    for position, child in enumerate(list(paragraph)):
        if child.tag == _PPR:
            continue
        if child.tag == _PROOF_ERR:
            paragraph.remove(child)
            continue
        if child.tag in _TRANSPARENT_PARAGRAPH_CHILDREN:
            continue  # bookmarks / comment anchors keep their places
        if placement is None:
            placement = position
        paragraph.remove(child)
        for t_elm in child.iter(_T):
            t_elm.tag = _DEL_TEXT
        deletion.append(child)
    # the deletion sits WHERE the content was, so comment range marks around
    # it keep wrapping it and reject restores the original order exactly
    if placement is None:
        paragraph.append(deletion)
    else:
        placement = min(placement, len(paragraph))
        paragraph.insert(placement, deletion)
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
    root = dict(_story_elements(document))[story]
    _refuse_paragraph_in_open_field(story, root, start_p, for_insertion=False)
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
    _validate_writable_text(heading, argument="heading")
    for index, text in enumerate(paragraphs):
        _validate_writable_text(text, argument=f"paragraphs[{index}]")
    heading_style_id = _validated_style_id(document, heading_style, argument="heading_style")
    body_style_id = _validated_style_id(document, body_style, argument="body_style")
    story, anchor_p = _resolve_anchor_paragraph(document, anchor)
    _refuse_cell_anchor(anchor_p)
    root = dict(_story_elements(document))[story]
    _refuse_paragraph_in_open_field(story, root, anchor_p, for_insertion=True)

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
    for index, text in enumerate(replacement_paragraphs):
        _validate_writable_text(text, argument=f"replacement_paragraphs[{index}]")
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


# ---------------------------------------------------------------------------
# v0.1 V3 — rich block insertion (a small TYPED block vocabulary, not
# arbitrary richness)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class TextRun:
    """One run of a rich paragraph."""

    text: str
    bold: bool = False
    italic: bool = False


@dataclass(frozen=True)
class RichParagraph:
    """A paragraph built from styled runs."""

    runs: Sequence[TextRun]
    style: Optional[str] = None  # style ID, validated against the document


@dataclass(frozen=True)
class ListBlock:
    """A real bullet/decimal list (backed by a real numbering definition)."""

    items: Sequence[str]
    kind: str = "bullet"  # "bullet" | "decimal"
    level: int = 0


@dataclass(frozen=True)
class TableBlock:
    """A simple rectangular table of plain-text cells."""

    rows: Sequence[Sequence[str]]


def _validate_rich_blocks(document: "Document", blocks, *, tracked: bool) -> None:
    if not blocks:
        raise ValueError("blocks must not be empty")
    for position, block in enumerate(blocks):
        label = f"blocks[{position}]"
        if isinstance(block, RichParagraph):
            if not block.runs:
                raise ValueError(f"{label}: RichParagraph needs at least one run")
            for run in block.runs:
                _validate_writable_text(run.text, argument=f"{label} run text")
            _validated_style_id(document, block.style, argument=f"{label}.style")
        elif isinstance(block, ListBlock):
            if block.kind not in ("bullet", "decimal"):
                raise ValueError(f"{label}: kind must be 'bullet' or 'decimal'")
            if not 0 <= block.level <= 2:
                raise ValueError(f"{label}: level must be 0..2")
            if not block.items:
                raise ValueError(f"{label}: ListBlock needs at least one item")
            for item in block.items:
                _validate_writable_text(item, argument=f"{label} item")
        elif isinstance(block, TableBlock):
            if tracked:
                raise UnsupportedStructureError(
                    "tracked table insertion is not supported in v0.1 (Word"
                    " marks table revisions row-by-row; that vocabulary is a"
                    " later phase) — insert the table untracked or as text"
                )
            if not block.rows or not all(block.rows):
                raise ValueError(f"{label}: TableBlock rows must be non-empty")
            width = len(block.rows[0])
            if any(len(r) != width for r in block.rows):
                raise ValueError(f"{label}: TableBlock rows must be rectangular")
            for row in block.rows:
                for cell_text in row:
                    _validate_writable_text(cell_text, argument=f"{label} cell")
        else:
            raise TypeError(
                f"{label}: expected RichParagraph, ListBlock, or TableBlock,"
                f" got {type(block).__name__}"
            )


def _new_rich_paragraph(rich: RichParagraph, style_id: Optional[str]) -> "_Element":
    paragraph = OxmlElement("w:p")
    if style_id:
        paragraph.style = style_id
    for run_spec in rich.runs:
        run = OxmlElement("w:r")
        if run_spec.bold or run_spec.italic:
            r_pr = OxmlElement("w:rPr")
            if run_spec.bold:
                r_pr.append(OxmlElement("w:b"))
            if run_spec.italic:
                r_pr.append(OxmlElement("w:i"))
            run.append(r_pr)
        run.add_t(run_spec.text)
        paragraph.append(run)
    return paragraph


def _new_list_paragraphs(document: "Document", block: ListBlock) -> "List[_Element]":
    from docx.numbering import ensure_bullet_definition, ensure_decimal_definition
    from docx.oxml.ns import nsdecls
    from docx.oxml.parser import parse_xml

    if block.kind == "bullet":
        num_id = ensure_bullet_definition(document)
    else:
        num_id = ensure_decimal_definition(document)
    paragraphs = []
    for item in block.items:
        paragraph = _new_paragraph(item, None)
        p_pr = paragraph.get_or_add_pPr()
        p_pr.append(
            parse_xml(
                f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{block.level}"/>'
                f'<w:numId w:val="{num_id}"/></w:numPr>'
            )
        )
        paragraphs.append(paragraph)
    return paragraphs


def _new_table(document: "Document", block: TableBlock) -> "_Element":
    from docx.oxml.table import CT_Tbl
    from docx.shared import Emu, Inches

    try:
        section = document.sections[-1]
        width = Emu(
            int(section.page_width) - int(section.left_margin) - int(section.right_margin)
        )
    except (IndexError, TypeError):
        width = Inches(6)
    tbl = CT_Tbl.new_tbl(len(block.rows), len(block.rows[0]), width)
    for tr, row_values in zip(tbl.tr_lst, block.rows):
        for tc, cell_text in zip(tr.tc_lst, row_values):
            if cell_text:
                paragraph = tc.find(qn("w:p"))
                run = OxmlElement("w:r")
                run.add_t(cell_text)
                paragraph.append(run)
    return tbl


def insert_blocks_after(
    document: "Document",
    anchor: AnchorLike,
    *,
    blocks: "Sequence[object]",
    tracked: bool = False,
    author: Optional[str] = None,
    date: Optional[dt.datetime] = None,
) -> BlockEditResult:
    """Insert a typed block list (v0.1 V3) after `anchor`.

    `blocks` mixes |RichParagraph| (styled runs), |ListBlock| (REAL bullet or
    decimal lists — the numbering definition is created on demand), and
    |TableBlock| (simple rectangular tables). This is deliberately a small
    vocabulary: rich enough for a report section, small enough to stay safe.
    Tracked mode covers paragraphs and lists; tracked TABLE insertion refuses
    (Word's row-revision vocabulary is a later phase).
    """
    if tracked and not author:
        raise ValueError("author is required when tracked=True")
    _validate_rich_blocks(document, blocks, tracked=tracked)
    story, anchor_p = _resolve_anchor_paragraph(document, anchor)
    _refuse_cell_anchor(anchor_p)
    root = dict(_story_elements(document))[story]
    _refuse_paragraph_in_open_field(story, root, anchor_p, for_insertion=True)

    # -- validated; build and mutate --
    stamp = date if date is not None else _clock.now()
    nodes: "List[_Element]" = []
    paragraph_nodes: "List[_Element]" = []
    for block in blocks:
        if isinstance(block, RichParagraph):
            style_id = _validated_style_id(
                document, block.style, argument="RichParagraph.style"
            )
            node = _new_rich_paragraph(block, style_id)
            nodes.append(node)
            paragraph_nodes.append(node)
        elif isinstance(block, ListBlock):
            for node in _new_list_paragraphs(document, block):
                nodes.append(node)
                paragraph_nodes.append(node)
        else:
            nodes.append(_new_table(document, block))
    # Word fuses adjacent sibling tables and refuses a cell/body ending in
    # w:tbl: pad tables with an empty paragraph where needed
    padded: "List[_Element]" = []
    for position, node in enumerate(nodes):
        padded.append(node)
        if node.tag != qn("w:tbl"):
            continue
        next_in_batch = nodes[position + 1] if position + 1 < len(nodes) else None
        if next_in_batch is not None and next_in_batch.tag == qn("w:tbl"):
            padded.append(OxmlElement("w:p"))
        elif next_in_batch is None:
            following = anchor_p.getnext()
            if following is None or following.tag == qn("w:tbl"):
                padded.append(OxmlElement("w:p"))
    nodes = padded
    revision_ids: "List[int]" = []
    if tracked:
        next_id = _next_revision_id(document)
        for node in paragraph_nodes:
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
