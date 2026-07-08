"""Revision enumeration and resolution (paper-docx, `Document.revisions`).

Enumerates every tracked change across every story part and resolves them:
accept applies the change (insertions unwrap, deletions disappear), reject
undoes it (insertions disappear, deleted text returns as live `w:t`).

Paragraph-mark revisions (`w:pPr/w:rPr/w:ins|w:del`, as emitted by
`docx.blocks`) are handled with the content: accepting a mark-deleted
paragraph that has no visible content left removes the paragraph; rejecting a
mark-inserted paragraph removes it entirely.

Resolution mutates the in-memory document; save with
`docx.package.patch_save` to keep the changed-part budget narrow.
"""

from __future__ import annotations

import copy
import datetime as dt
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, List, Optional, Sequence, Tuple

from docx.errors import UnsupportedStructureError
from docx.oxml.ns import qn
from docx.protection import _refuse_if_protected
from docx.story import (
    Anchor,
    _first_choice_children,
    _iter_block_elements,
    _story_elements,
    content_hash,
)

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

_INS = qn("w:ins")
_DEL = qn("w:del")
_MOVE_FROM = qn("w:moveFrom")
_MOVE_TO = qn("w:moveTo")
_T = qn("w:t")
_DEL_TEXT = qn("w:delText")
_INSTR_TEXT = qn("w:instrText")
_DEL_INSTR_TEXT = qn("w:delInstrText")
_R = qn("w:r")
_RPR = qn("w:rPr")
_PPR = qn("w:pPr")
_P = qn("w:p")
_TR = qn("w:tr")
_TRPR = qn("w:trPr")
_TBL = qn("w:tbl")
_SECT_PR = qn("w:sectPr")
_RPR_CHANGE = qn("w:rPrChange")
_PPR_CHANGE = qn("w:pPrChange")
_AUTHOR = qn("w:author")
_DATE = qn("w:date")

#: tag -> revision_type for everything Document.revisions enumerates.
#: `w:ins`/`w:del` refine to "row_insertion"/"row_deletion" when the node is
#: a `w:trPr` row marker (see `_revision_type_of`) — classifying those as
#: plain insertion/deletion would resolve just the MARKER and leave ghost
#: rows behind, reporting false state.
_REVISION_TYPES = {
    _INS: "insertion",
    _DEL: "deletion",
    _MOVE_FROM: "move_from",
    _MOVE_TO: "move_to",
    _RPR_CHANGE: "format_change",
    _PPR_CHANGE: "format_change",
}
for _change_tag in ("w:tblPrChange", "w:tblPrExChange", "w:tblGridChange",
                    "w:trPrChange", "w:tcPrChange"):
    _REVISION_TYPES[qn(_change_tag)] = "table_property_change"
for _change_tag in ("w:cellIns", "w:cellDel", "w:cellMerge"):
    _REVISION_TYPES[qn(_change_tag)] = "cell_revision"
_REVISION_TYPES[qn("w:sectPrChange")] = "section_property_change"
_REVISION_TYPES[qn("w:numberingChange")] = "numbering_change"
for _change_tag in ("w:customXmlInsRangeStart", "w:customXmlDelRangeStart",
                    "w:customXmlMoveFromRangeStart", "w:customXmlMoveToRangeStart"):
    _REVISION_TYPES[qn(_change_tag)] = "custom_xml_revision"
del _change_tag


def _revision_type_of(node: "_Element") -> str:
    parent = node.getparent()
    if parent is not None and parent.tag == _TRPR:
        return "row_insertion" if node.tag == _INS else "row_deletion"
    return _REVISION_TYPES[node.tag]


#: the only revision types accept()/reject() know how to resolve correctly.
#: move_from/move_to resolve as PAIRED UNITS: accepting or rejecting either
#: site of a move resolves both (never one side alone).
RESOLVABLE_TYPES = frozenset(
    {
        "insertion",
        "deletion",
        "format_change",
        "row_insertion",
        "row_deletion",
        "move_from",
        "move_to",
    }
)


def _node_text(node: "_Element") -> str:
    pieces: List[str] = []
    for child in node.iter():
        if child.tag in (_T, _DEL_TEXT):
            pieces.append(child.text or "")
    return "".join(pieces)


def _applies_to_text(node: "_Element") -> str:
    """The text a property-change revision APPLIES to (its own subtree holds
    only stored properties, so `_node_text` is empty and the revision would
    be unaddressable): the containing run's text for a run change, the
    containing paragraph's for a paragraph or paragraph-mark change."""
    parent = node.getparent()
    if parent is None:
        return ""
    if parent.tag == _RPR:
        holder = parent.getparent()  # w:r, or w:pPr for a paragraph mark
        if holder is not None and holder.tag == _PPR:
            holder = holder.getparent()
        return _node_text(holder) if holder is not None else ""
    if parent.tag == _PPR:  # w:pPrChange
        paragraph = parent.getparent()
        return _node_text(paragraph) if paragraph is not None else ""
    return ""


def _is_paragraph_mark_revision(node: "_Element") -> bool:
    parent = node.getparent()
    return parent is not None and parent.tag == _RPR and parent.getparent() is not None \
        and parent.getparent().tag == _PPR


def _parse_date(value: Optional[str]) -> Optional[dt.datetime]:
    if not value:
        return None
    try:
        return dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


@dataclass(frozen=True)
class Revision:
    """One tracked change, addressable — and resolvable when supported.

    Resolvable `revision_type`s: "insertion", "deletion", "format_change"
    (`w:rPrChange`/`w:pPrChange`, run or paragraph mark), "row_insertion",
    "row_deletion" (`w:trPr` row markers), and — as paired units —
    "move_from"/"move_to". The exotic remainder ("table_property_change",
    "cell_revision", "section_property_change", "numbering_change",
    "custom_xml_revision") is enumerated and counted but resolution is
    refused by name — claiming to resolve them would report false state.
    """

    revision_type: str
    author: str
    date: Optional[dt.datetime]
    text: str
    story: str
    anchor: Anchor
    is_paragraph_mark: bool
    _element: "_Element"
    _document: "Optional[Document]" = None

    @property
    def is_resolvable(self) -> bool:
        return self.revision_type in RESOLVABLE_TYPES

    def _refuse_unresolvable(self, verb: str) -> None:
        if not self.is_resolvable:
            raise UnsupportedStructureError(
                f"cannot {verb} a {self.revision_type!r} revision: this"
                " revision type is enumerated but not resolvable by"
                " paper-docx (resolve it in Word, or a later paper-docx"
                " version)"
            )

    def accept(self) -> None:
        """Apply this change to the document. Tracked moves resolve as a
        PAIR: accepting either site accepts both (v0.11 Phase 2)."""
        self._refuse_unresolvable("accept")
        if self._document is not None:
            _refuse_if_protected(self._document, "resolve a revision")
        _resolve_one(self._element, accept=True, document=self._document)

    def reject(self) -> None:
        """Undo this change, restoring the pre-change content. Tracked moves
        resolve as a PAIR: rejecting either site rejects both."""
        self._refuse_unresolvable("reject")
        if self._document is not None:
            _refuse_if_protected(self._document, "resolve a revision")
        _resolve_one(self._element, accept=False, document=self._document)

    def to_dict(self) -> dict:
        return {
            "revision_type": self.revision_type,
            "author": self.author,
            "date": self.date.isoformat() if self.date else None,
            "text": self.text,
            "story": self.story,
            "anchor": self.anchor.to_dict(),
            "is_paragraph_mark": self.is_paragraph_mark,
        }


class Revisions(Sequence[Revision]):
    """All tracked changes in a document, across every story part.

    A fresh snapshot is enumerated on each `Document.revisions` access;
    resolving revisions invalidates previously-held |Revision| objects.
    """

    def __init__(self, document: "Document") -> None:
        self._document = document
        self._items = tuple(_enumerate_revisions(document))

    def __len__(self) -> int:
        return len(self._items)

    def __getitem__(self, key):  # noqa: ANN001 - Sequence protocol
        return self._items[key]

    def __iter__(self) -> Iterator[Revision]:
        return iter(self._items)

    def accept_all(self, *, author: Optional[str] = None) -> int:
        """Apply every selected revision (optionally only `author`'s).

        Validates the WHOLE selected set first: if it contains revision types
        this package cannot resolve (moves, formatting changes), the call
        refuses atomically — it never half-resolves and reports success while
        Word still shows pending changes. Returns the resolved count; check
        `remaining_unsupported()` before inferring "the document is clean".
        """
        return self._resolve_all(accept=True, author=author)

    def reject_all(self, *, author: Optional[str] = None) -> int:
        """Undo every selected revision (optionally only `author`'s).

        Same selected-set validation and census semantics as `accept_all`.
        """
        return self._resolve_all(accept=False, author=author)

    def remaining_unsupported(self) -> dict:
        """{revision_type: count} for enumerated-but-unresolvable revisions."""
        census: dict = {}
        for revision in self._items:
            if not revision.is_resolvable:
                census[revision.revision_type] = census.get(revision.revision_type, 0) + 1
        return dict(sorted(census.items()))

    def _resolve_all(self, *, accept: bool, author: Optional[str]) -> int:
        _refuse_if_protected(self._document, "resolve revisions")
        selected = [
            revision
            for revision in self._items
            if author is None or revision.author == author
        ]
        unresolvable = sorted(
            {r.revision_type for r in selected if not r.is_resolvable}
        )
        if unresolvable:
            raise UnsupportedStructureError(
                f"selected revisions include {unresolvable} which this package"
                " can enumerate but not resolve; nothing was changed. Filter"
                " by author, resolve individual insertions/deletions, or"
                " resolve the rest in Word"
            )
        _validate_moves(selected, self._document)
        resolved = 0
        # content revisions first, then paragraph marks (mark resolution can
        # remove whole paragraphs and must see post-content state)
        ordered = sorted(selected, key=lambda item: item.is_paragraph_mark)
        for revision in ordered:
            _resolve_one(  # noqa: SLF001
                revision._element, accept=accept, document=self._document
            )
            resolved += 1
        self._items = tuple(_enumerate_revisions(self._document))
        return resolved

    def to_dict(self) -> dict:
        return {
            "schema": "paper_revisions",
            # v2: move/format_change types + census (v0.1 H1-H3)
            # v3: row_insertion/row_deletion + named exotic types; format
            #     changes and row revisions resolvable (v0.11 Phase 1)
            "version": 3,
            "revisions": [revision.to_dict() for revision in self._items],
            "remaining_unsupported": self.remaining_unsupported(),
        }


# ---------------------------------------------------------------------------
# move units (Phase 2): w:moveFrom/w:moveTo paired by range-marker w:name
# ---------------------------------------------------------------------------

_MOVE_FROM_RANGE_START = qn("w:moveFromRangeStart")
_MOVE_FROM_RANGE_END = qn("w:moveFromRangeEnd")
_MOVE_TO_RANGE_START = qn("w:moveToRangeStart")
_MOVE_TO_RANGE_END = qn("w:moveToRangeEnd")
_ID = qn("w:id")
_NAME = qn("w:name")


class _MoveSite:
    """One side of a tracked move: its range brackets, run wrappers
    (`w:moveFrom`/`w:moveTo`) and paragraph-mark stamps, in one story."""

    def __init__(self, story: str, name: str, start: "_Element") -> None:
        self.story = story
        self.name = name
        self.start = start
        self.end: "Optional[_Element]" = None
        self.wrappers: "List[_Element]" = []
        self.mark_stamps: "List[_Element]" = []

    @property
    def elements(self) -> "List[_Element]":
        markers = [self.start] + ([self.end] if self.end is not None else [])
        return markers + self.wrappers + self.mark_stamps


class _MoveUnit:
    """A complete tracked move: source and destination site, same name."""

    def __init__(self, from_site: _MoveSite, to_site: _MoveSite) -> None:
        self.name = from_site.name
        self.from_site = from_site
        self.to_site = to_site

    @property
    def elements(self) -> "List[_Element]":
        return self.from_site.elements + self.to_site.elements


def _scan_story_moves(
    story: str,
    root: "_Element",
    sites: "List[_MoveSite]",
    orphans: "List[Tuple[_Element, str]]",
) -> None:
    """Collect move sites in document order; anything unpaired is an orphan."""
    open_site: dict = {"from": None, "to": None}
    directions = {
        _MOVE_FROM_RANGE_START: ("from", "start"),
        _MOVE_TO_RANGE_START: ("to", "start"),
        _MOVE_FROM_RANGE_END: ("from", "end"),
        _MOVE_TO_RANGE_END: ("to", "end"),
        _MOVE_FROM: ("from", "content"),
        _MOVE_TO: ("to", "content"),
    }

    def walk(element: "_Element") -> None:
        for child in _first_choice_children(element):
            role = directions.get(child.tag)
            if role is not None:
                direction, kind = role
                site = open_site[direction]
                if kind == "start":
                    name = child.get(_NAME)
                    if site is not None:
                        orphans.append((child, "nested move ranges"))
                    elif not name:
                        orphans.append((child, "move range without a w:name"))
                    else:
                        open_site[direction] = _MoveSite(story, name, child)
                elif kind == "end":
                    if site is not None and child.get(_ID) == site.start.get(_ID):
                        site.end = child
                        sites.append(site)
                        open_site[direction] = None
                    else:
                        orphans.append((child, "move range end without its start"))
                else:  # content: a wrapper or a paragraph-mark stamp
                    if site is None:
                        orphans.append(
                            (child, "move content outside any move range")
                        )
                    elif _is_paragraph_mark_revision(child):
                        site.mark_stamps.append(child)
                    else:
                        site.wrappers.append(child)
            walk(child)

    walk(root)
    for direction in ("from", "to"):
        if open_site[direction] is not None:
            orphans.append(
                (open_site[direction].start, "move range start without its end")
            )


def _move_units(
    document: "Document",
) -> "Tuple[List[_MoveUnit], List[Tuple[_Element, str]]]":
    """All well-formed move units, plus every orphaned move element."""
    sites: "List[_MoveSite]" = []
    orphans: "List[Tuple[_Element, str]]" = []
    for story, root in _story_elements(document):
        _scan_story_moves(story, root, sites, orphans)
    by_name: dict = {}
    for site in sites:
        direction = "from" if site.start.tag == _MOVE_FROM_RANGE_START else "to"
        by_name.setdefault(site.name, {"from": [], "to": []})[direction].append(site)
    units: "List[_MoveUnit]" = []
    for name, pair in sorted(by_name.items()):
        froms, tos = pair["from"], pair["to"]
        if len(froms) != 1 or len(tos) != 1:
            reason = (
                f"move {name!r} has {len(froms)} source and {len(tos)}"
                " destination range(s); expected exactly one of each"
            )
            for site in froms + tos:
                orphans.extend((element, reason) for element in site.elements)
            continue
        if froms[0].story != tos[0].story:
            reason = (
                f"move {name!r} crosses stories ({froms[0].story} ->"
                f" {tos[0].story}); cross-story moves are not resolvable"
            )
            for site in froms + tos:
                orphans.extend((element, reason) for element in site.elements)
            continue
        units.append(_MoveUnit(froms[0], tos[0]))
    return units, orphans


def _resolve_move(
    node: "_Element", *, accept: bool, document: "Optional[Document]"
) -> None:
    if document is None:
        raise UnsupportedStructureError(
            "cannot resolve a tracked move without its document context"
        )
    units, orphans = _move_units(document)
    for unit in units:
        if any(element is node for element in unit.elements):
            _resolve_move_unit(unit, accept=accept, document=document)
            return
    for element, reason in orphans:
        if element is node:
            raise UnsupportedStructureError(
                f"cannot resolve this tracked move: {reason}; nothing was"
                " changed. Resolve it in Word instead"
            )
    raise UnsupportedStructureError(
        "cannot resolve this tracked move: its range markers were not found;"
        " nothing was changed. Resolve it in Word instead"
    )


def _resolve_move_unit(
    unit: _MoveUnit, *, accept: bool, document: "Optional[Document]"
) -> None:
    """Apply or undo a move as ONE unit (never one site alone).

    Accept: destination content becomes plain, source range disappears.
    Reject: source content becomes plain again, destination disappears.
    Mark stamps ride the existing paragraph-mark machinery (`w:moveFrom` is
    del-like, `w:moveTo` ins-like), AFTER range markers are removed so an
    emptied source/destination paragraph is recognized as empty.
    """
    keep_site = unit.to_site if accept else unit.from_site
    drop_site = unit.from_site if accept else unit.to_site
    for wrapper in keep_site.wrappers:
        if wrapper.getparent() is not None:
            _unwrap(wrapper)
    orphaned_comments: "List[int]" = []
    for wrapper in drop_site.wrappers:
        if wrapper.getparent() is not None:
            orphaned_comments.extend(_comment_ids_inside(wrapper))
            wrapper.getparent().remove(wrapper)
    for site in (unit.from_site, unit.to_site):
        for marker in (site.start, site.end):
            if marker is not None and marker.getparent() is not None:
                marker.getparent().remove(marker)
    for site in (unit.from_site, unit.to_site):
        for stamp in site.mark_stamps:
            if stamp.getparent() is not None:
                _resolve_paragraph_mark(stamp, accept=accept)
    _cleanup_comment_anchors(document, orphaned_comments)


def _validate_moves(selected: "List[Revision]", document: "Document") -> None:
    """Refuse BEFORE mutating when any selected move is orphaned, duplicated
    or cross-story (refusal atomicity for the batch)."""
    move_nodes = [
        revision._element  # noqa: SLF001
        for revision in selected
        if revision.revision_type in ("move_from", "move_to")
    ]
    if not move_nodes:
        return
    units, orphans = _move_units(document)
    unit_elements = {id(e) for unit in units for e in unit.elements}
    orphan_reasons = {id(element): reason for element, reason in orphans}
    for node in move_nodes:
        if id(node) in unit_elements:
            continue
        reason = orphan_reasons.get(
            id(node), "its range markers were not found"
        )
        raise UnsupportedStructureError(
            f"selected revisions include an unresolvable tracked move:"
            f" {reason}; nothing was changed. Resolve it in Word instead"
        )


def _iter_revision_nodes(
    element: "_Element", *, skip_text_boxes: bool
) -> Iterator["_Element"]:
    """Revision nodes (ins/del/moves/format changes) in traversal space:
    mc:Fallback duplicates excluded, and text-box content excluded for
    paragraph blocks (those revisions belong to the text box's own blocks)."""
    for child in _first_choice_children(element):
        if skip_text_boxes and child.tag == qn("w:txbxContent"):
            continue
        if child.tag in _REVISION_TYPES:
            yield child
        yield from _iter_revision_nodes(child, skip_text_boxes=skip_text_boxes)


def _enumerate_revisions(document: "Document") -> Iterator[Revision]:
    from docx.story import _build_block

    for story, root in _story_elements(document):
        for kind, index, element, in_sdt, in_txbx in _iter_block_elements(story, root):
            skip_boxes = kind == "paragraph"
            block_anchor = None
            for node in _iter_revision_nodes(element, skip_text_boxes=skip_boxes):
                if block_anchor is None:
                    # the anchor is the containing BLOCK's (so it verifies
                    # against outline blocks and is usable as an AnchorLike)
                    block_anchor = _build_block(
                        story, kind, index, element, "current",
                        in_sdt=in_sdt, in_txbx=in_txbx,
                    ).anchor
                revision_type = _revision_type_of(node)
                text = _node_text(node)
                if not text and node.tag in (_RPR_CHANGE, _PPR_CHANGE):
                    text = _applies_to_text(node)
                yield Revision(
                    revision_type=revision_type,
                    author=node.get(_AUTHOR) or "",
                    date=_parse_date(node.get(_DATE)),
                    text=text,
                    story=story,
                    anchor=block_anchor,
                    is_paragraph_mark=_is_paragraph_mark_revision(node),
                    _element=node,
                    _document=document,
                )


#: every tag that constitutes revision markup, for the post-resolution rescan
#: (enumerated types plus the range brackets that resolution must sweep up)
_MARKUP_SCAN_TAGS = tuple(_REVISION_TYPES) + tuple(
    qn(tag)
    for tag in (
        "w:moveFromRangeStart", "w:moveFromRangeEnd",
        "w:moveToRangeStart", "w:moveToRangeEnd",
        "w:customXmlInsRangeEnd", "w:customXmlDelRangeEnd",
        "w:customXmlMoveFromRangeEnd", "w:customXmlMoveToRangeEnd",
    )
)


def _iter_markup_nodes(element: "_Element") -> Iterator["_Element"]:
    for child in _first_choice_children(element):
        if child.tag in _MARKUP_SCAN_TAGS:
            yield child
        yield from _iter_markup_nodes(child)


def _remaining_markup(document: "Document") -> dict:
    """{local-tag-name: count} of ALL revision markup left in traversal space.

    The invariant oracle: after a successful `accept_all()`/`reject_all()`
    (and, post-Phase-2, move resolution) this is empty — "resolved" while
    markup remains anywhere would be false state.
    """
    counts: dict = {}
    for _story, root in _story_elements(document):
        for node in _iter_markup_nodes(root):
            name = node.tag.rsplit("}", 1)[-1]
            counts[name] = counts.get(name, 0) + 1
    return dict(sorted(counts.items()))


def _comment_ids_inside(node: "_Element"):
    return [
        int(ref.get(qn("w:id")))
        for ref in node.iter(qn("w:commentReference"))
        if ref.get(qn("w:id"))
    ]


def _cleanup_comment_anchors(document: "Optional[Document]", comment_ids) -> None:
    """A resolution removed the run holding a comment's reference mark: also
    remove the now-orphaned range markers and the comment itself, exactly as
    Word does — half-deleted comments are silent corruption."""
    if document is None or not comment_ids:
        return
    wanted = set(comment_ids)
    for _story, root in _story_elements(document):
        for marker in list(
            root.iter(qn("w:commentRangeStart"), qn("w:commentRangeEnd"))
        ):
            raw = marker.get(qn("w:id"))
            if raw and int(raw) in wanted:
                marker.getparent().remove(marker)
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        comments_part = document.part.part_related_by(RT.COMMENTS)
    except KeyError:
        return
    comments_root = comments_part._element  # noqa: SLF001
    for comment in list(comments_root):
        raw = comment.get(qn("w:id"))
        if raw and int(raw) in wanted:
            comments_root.remove(comment)


def _resolve_one(
    node: "_Element", *, accept: bool, document: "Optional[Document]" = None
) -> None:
    if node.getparent() is None:
        return  # already resolved via an enclosing operation
    revision_type = _revision_type_of(node)
    if revision_type in ("move_from", "move_to"):
        _resolve_move(node, accept=accept, document=document)
        return
    if revision_type == "format_change":
        _resolve_format_change(node, accept=accept)
        return
    if revision_type in ("row_insertion", "row_deletion"):
        _resolve_row_revision(node, accept=accept, document=document)
        return
    if _is_paragraph_mark_revision(node):
        _resolve_paragraph_mark(node, accept=accept)
        return
    removes_content = (node.tag == _INS and not accept) or (
        node.tag == _DEL and accept
    )
    orphaned = _comment_ids_inside(node) if removes_content else []
    if node.tag == _INS:
        _resolve_insertion(node, accept=accept)
    else:
        _resolve_deletion(node, accept=accept)
    _cleanup_comment_anchors(document, orphaned)


def _unwrap(node: "_Element") -> None:
    parent = node.getparent()
    for child in list(node):
        node.addprevious(child)
    parent.remove(node)


def _resolve_insertion(node: "_Element", *, accept: bool) -> None:
    if accept:
        _unwrap(node)
    else:
        node.getparent().remove(node)


def _resolve_deletion(node: "_Element", *, accept: bool) -> None:
    if accept:
        node.getparent().remove(node)
    else:
        for text_elm in node.iter(_DEL_TEXT):
            text_elm.tag = _T
        for text_elm in node.iter(_DEL_INSTR_TEXT):
            text_elm.tag = _INSTR_TEXT
        _unwrap(node)


#: children of a paragraph-mark `w:rPr` (CT_ParaRPr) that are revision
#: bookkeeping, not formatting — a format-change reject must preserve them.
_PARA_RPR_STAMPS = (_INS, _DEL, _MOVE_FROM, _MOVE_TO)


def _resolve_format_change(node: "_Element", *, accept: bool) -> None:
    """Resolve `w:rPrChange`/`w:pPrChange` (run, paragraph or paragraph mark).

    Accept keeps the current properties and drops the stored previous ones;
    reject swaps the stored previous properties back in. On a paragraph-mark
    `w:rPr` the revision stamps (`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`)
    are bookkeeping, not formatting, and survive a reject; on a `w:pPr` the
    mark's `w:rPr` and any `w:sectPr` likewise stay.
    """
    parent = node.getparent()
    if accept:
        parent.remove(node)
        if len(parent) == 0:
            parent.getparent().remove(parent)
        return
    if node.tag == _RPR_CHANGE:
        stored = node.find(_RPR)
        kept = [child for child in parent if child.tag in _PARA_RPR_STAMPS]
    else:  # w:pPrChange stores a CT_PPrBase (never rPr/sectPr)
        stored = node.find(_PPR)
        kept = [child for child in parent if child.tag in (_RPR, _SECT_PR)]
    restored = [copy.deepcopy(child) for child in stored] if stored is not None else []
    # CT_ParaRPr puts revision stamps first; CT_PPr puts rPr/sectPr last
    new_children = kept + restored if node.tag == _RPR_CHANGE else restored + kept
    for child in list(parent):
        parent.remove(child)
    for child in new_children:
        parent.append(child)
    if len(parent) == 0:
        parent.getparent().remove(parent)


def _resolve_row_revision(
    node: "_Element", *, accept: bool, document: "Optional[Document]" = None
) -> None:
    """Resolve a `w:trPr` row marker (`w:ins` = row inserted with tracking,
    `w:del` = row deleted with tracking).

    Keeping the row = remove just the marker. Removing the row = remove the
    whole `w:tr` (the cell-level content revisions inside it are subsumed,
    exactly as Word's Accept/Reject All treats them). Removing a table's
    LAST row removes the table itself — Word's semantic for a fully
    tracked-deleted table; a zero-row `w:tbl` would be invalid XML. When the
    table was its container's only block, an empty paragraph takes its place
    (a `w:tc`/body must keep one block).
    """
    tr_pr = node.getparent()
    row = tr_pr.getparent()
    keeps_row = (node.tag == _INS and accept) or (node.tag == _DEL and not accept)
    if keeps_row:
        tr_pr.remove(node)
        if len(tr_pr) == 0:
            row.remove(tr_pr)
        return
    table = row.getparent()
    if sum(1 for child in table if child.tag == _TR) == 1:
        orphaned = _comment_ids_inside(table)
        parent = table.getparent()
        siblings = [
            child for child in parent if child.tag in (_P, _TBL) and child is not table
        ]
        if not siblings:
            from docx.oxml.parser import OxmlElement

            table.addprevious(OxmlElement("w:p"))
        parent.remove(table)
        _cleanup_comment_anchors(document, orphaned)
        return
    orphaned = _comment_ids_inside(row)
    table.remove(row)
    _cleanup_comment_anchors(document, orphaned)


def _paragraph_of_mark(node: "_Element") -> "Optional[_Element]":
    r_pr = node.getparent()
    p_pr = r_pr.getparent() if r_pr is not None else None
    paragraph = p_pr.getparent() if p_pr is not None else None
    return paragraph if paragraph is not None and paragraph.tag == _P else None


def _paragraph_has_content(paragraph: "_Element") -> bool:
    return any(child.tag != _PPR for child in paragraph)


def _is_last_block_in_container(paragraph: "_Element") -> bool:
    parent = paragraph.getparent()
    if parent is None:
        return True
    blocks = [c for c in parent if c.tag in (_P, qn("w:tbl")) and c is not paragraph]
    return not blocks


def _next_paragraph_sibling(paragraph: "_Element") -> "Optional[_Element]":
    node = paragraph.getnext()
    while node is not None:
        if node.tag == _P:
            return node
        if node.tag == qn("w:tbl"):
            return None  # merging across a table is not a paragraph join
        node = node.getnext()
    return None


def _resolve_paragraph_mark(node: "_Element", *, accept: bool) -> None:
    """Resolve a paragraph-MARK revision (`w:pPr/w:rPr/w:ins|w:del`).

    Removing the mark applies/undoes the paragraph-break change:

    * mark removed + paragraph empty -> the paragraph disappears, unless it
      is the container's last block (a `w:tc`/body must keep one block —
      removing it would emit schema-invalid XML).
    * mark-DELETED paragraph that still has content (Word's "deleted
      pilcrow") -> accepting merges its content into the following
      paragraph, exactly as Word's Accept All does. The symmetric reject of
      a mark-INSERTED split merges the same way.
    """
    paragraph = _paragraph_of_mark(node)
    # w:moveFrom stamps are del-like (the mark moved AWAY from here),
    # w:moveTo stamps ins-like — same break-change algebra
    is_deletion = node.tag in (_DEL, _MOVE_FROM)
    applying_break_change = (accept and is_deletion) or (not accept and not is_deletion)

    r_pr = node.getparent()
    r_pr.remove(node)
    if len(r_pr) == 0:
        r_pr.getparent().remove(r_pr)
    if paragraph is None or not applying_break_change:
        return
    if _paragraph_has_content(paragraph):
        following = _next_paragraph_sibling(paragraph)
        if following is None:
            return  # nothing to join with; the content stays as a paragraph
        insert_at = 1 if following.find(_PPR) is not None else 0
        for child in reversed([c for c in paragraph if c.tag != _PPR]):
            following.insert(insert_at, child)
        paragraph.getparent().remove(paragraph)
        return
    if not _is_last_block_in_container(paragraph) and paragraph.getparent() is not None:
        paragraph.getparent().remove(paragraph)
