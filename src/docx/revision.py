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
from dataclasses import dataclass, field, replace
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
_COMMENT = qn("w:comment")
_COMMENT_RANGE_START = qn("w:commentRangeStart")
_COMMENT_RANGE_END = qn("w:commentRangeEnd")
_COMMENT_REFERENCE = qn("w:commentReference")
_FOOTNOTE = qn("w:footnote")
_ENDNOTE = qn("w:endnote")
_FOOTNOTE_REFERENCE = qn("w:footnoteReference")
_ENDNOTE_REFERENCE = qn("w:endnoteReference")
_NOTE_TYPE = qn("w:type")
_W14_PARA_ID = qn("w14:paraId")
_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
_W15_COMMENT_EX = f"{{{_W15_NS}}}commentEx"
_W15_PARA_ID = f"{{{_W15_NS}}}paraId"
_REL_NS_PREFIX = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
)
_OFFICE_REL_ID = "{urn:schemas-microsoft-com:office:office}relid"

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
    _snapshot_signature: "Tuple[_Element, ...]" = field(
        default=(), repr=False, compare=False
    )

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

    def _refuse_if_stale(self, verb: str) -> None:
        if self._document is None:
            if self._element.getparent() is None:
                raise UnsupportedStructureError(
                    f"cannot {verb} a stale Revision; reacquire it from"
                    " document.revisions. Nothing was changed"
                )
            return
        if (
            self._snapshot_signature != _revision_signature(self._document)
            or not _is_in_story(self._element, self._document)
        ):
            raise UnsupportedStructureError(
                f"cannot {verb} a stale Revision; reacquire it from"
                " document.revisions. Nothing was changed"
            )

    def accept(self) -> None:
        """Apply this change to the document. Tracked moves resolve as a
        PAIR: accepting either site accepts both."""
        self._refuse_unresolvable("accept")
        self._refuse_if_stale("accept")
        if self._document is not None:
            _refuse_if_protected(self._document, "resolve a revision")
            _preflight_resolution(
                self._document,
                [self],
                accept=True,
                author=None,
                require_clean=False,
            )
        _resolve_one(self._element, accept=True, document=self._document)

    def reject(self) -> None:
        """Undo this change, restoring the pre-change content. Tracked moves
        resolve as a PAIR: rejecting either site rejects both."""
        self._refuse_unresolvable("reject")
        self._refuse_if_stale("reject")
        if self._document is not None:
            _refuse_if_protected(self._document, "resolve a revision")
            _preflight_resolution(
                self._document,
                [self],
                accept=False,
                author=None,
                require_clean=False,
            )
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
    mutation through another snapshot invalidates this snapshot and its
    previously-held |Revision| objects. Stale resolution attempts refuse.
    """

    def __init__(self, document: "Document") -> None:
        self._document = document
        self._items, self._snapshot_signature = _revision_snapshot(document)

    def __len__(self) -> int:
        return len(self._items)

    def __getitem__(self, key):  # noqa: ANN001 - Sequence protocol
        return self._items[key]

    def __iter__(self) -> Iterator[Revision]:
        return iter(self._items)

    def accept_all(self, *, author: Optional[str] = None) -> int:
        """Apply every selected revision (optionally only `author`'s).

        Validates the WHOLE selected set first: if it contains revision types
        this package cannot resolve, the call refuses atomically — it never
        half-resolves and reports success while
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
        if self._snapshot_signature != _revision_signature(self._document):
            raise UnsupportedStructureError(
                "cannot resolve a stale Revisions snapshot; reacquire"
                " document.revisions. Nothing was changed"
            )
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
        _preflight_resolution(
            self._document,
            selected,
            accept=accept,
            author=author,
            require_clean=author is None,
        )
        # moves first (their range brackets must still be intact — a row
        # removal in the same batch could take a move site with it), then
        # other content, then paragraph marks (mark resolution can remove
        # whole paragraphs and must see post-content state)
        def _order(item: Revision) -> int:
            if item.revision_type in ("move_from", "move_to"):
                return 0
            return 2 if item.is_paragraph_mark else 1

        ordered = sorted(selected, key=_order)
        handled: set[int] = set()
        for revision in ordered:
            element_id = id(revision._element)  # noqa: SLF001
            if element_id in handled:
                continue
            if not _is_in_story(revision._element, self._document):  # noqa: SLF001
                # An earlier move/row/enclosing-wrapper operation resolved
                # this selected node as part of its compound unit.
                handled.add(element_id)
                continue
            _resolve_one(  # noqa: SLF001
                revision._element, accept=accept, document=self._document
            )
            handled.add(element_id)
        self._items, self._snapshot_signature = _revision_snapshot(self._document)
        return len(selected)

    def to_dict(self) -> dict:
        return {
            "schema": "paper_revisions",
            # v2: move/format_change types + census
            # v3: row_insertion/row_deletion + named exotic types; format
            #     changes and row revisions resolvable
            "version": 3,
            "revisions": [revision.to_dict() for revision in self._items],
            "remaining_unsupported": self.remaining_unsupported(),
        }


# ---------------------------------------------------------------------------
# move units: w:moveFrom/w:moveTo paired by range-marker w:name
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
                        # overlapping/nested ranges: wrapper attribution is
                        # ambiguous — orphan BOTH ranges, never guess
                        orphans.append((child, "overlapping move ranges"))
                        orphans.extend(
                            (element, "overlapping move ranges")
                            for element in site.elements
                        )
                        open_site[direction] = None
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
    for wrapper in drop_site.wrappers:
        if wrapper.getparent() is not None:
            _discard_element(wrapper, document)
    for site in (unit.from_site, unit.to_site):
        for marker in (site.start, site.end):
            if marker is not None and marker.getparent() is not None:
                marker.getparent().remove(marker)
    for site in (unit.from_site, unit.to_site):
        for stamp in site.mark_stamps:
            if stamp.getparent() is not None:
                _resolve_paragraph_mark(stamp, accept=accept)


def _validate_moves(
    selected: "List[Revision]",
    document: "Document",
    *,
    author: "Optional[str]",
) -> None:
    """Refuse BEFORE mutating when any selected move is orphaned, duplicated
    or cross-story, or an author filter would resolve another author's side
    of the same compound move (refusal atomicity for the batch)."""
    move_nodes = [
        revision._element  # noqa: SLF001
        for revision in selected
        if revision.revision_type in ("move_from", "move_to")
    ]
    if not move_nodes:
        return
    units, orphans = _move_units(document)
    units_by_element = {
        id(element): unit for unit in units for element in unit.elements
    }
    orphan_reasons = {id(element): reason for element, reason in orphans}
    touched_units: dict[int, _MoveUnit] = {}
    for node in move_nodes:
        unit = units_by_element.get(id(node))
        if unit is not None:
            touched_units[id(unit)] = unit
            continue
        reason = orphan_reasons.get(
            id(node), "its range markers were not found"
        )
        raise UnsupportedStructureError(
            f"selected revisions include an unresolvable tracked move:"
            f" {reason}; nothing was changed. Resolve it in Word instead"
        )
    if author is None:
        return
    for unit in touched_units.values():
        revision_elements = (
            unit.from_site.wrappers
            + unit.from_site.mark_stamps
            + unit.to_site.wrappers
            + unit.to_site.mark_stamps
        )
        unit_authors = {element.get(_AUTHOR) or "" for element in revision_elements}
        unit_authors.update(
            value
            for element in (
                unit.from_site.start,
                unit.from_site.end,
                unit.to_site.start,
                unit.to_site.end,
            )
            if element is not None
            if (value := element.get(_AUTHOR)) is not None
        )
        if unit_authors != {author}:
            raise UnsupportedStructureError(
                f"cannot resolve move {unit.name!r} with author={author!r}:"
                f" its compound move unit's authors differ"
                f" ({sorted(unit_authors)!r}); nothing was changed"
            )


def _preflight_resolution(
    document: "Document",
    selected: "List[Revision]",
    *,
    accept: bool,
    author: "Optional[str]",
    require_clean: bool,
) -> None:
    """Validate every dependency a resolution can touch before mutation."""
    _validate_comment_marker_ids(document)
    _validate_moves(selected, document, author=author)
    if require_clean:
        # An unfiltered resolution certifies "clean afterwards". Refuse
        # upfront anything that would falsify that claim.
        _refuse_unaccounted_markup(document)
    _validate_paragraph_mark_joins(selected, document, accept=accept)
    discard_roots = _resolution_discard_roots(selected, document, accept=accept)
    _validate_filtered_destructive_closure(
        discard_roots, selected, document, author=author
    )
    _validate_cleanup_dependencies(document, discard_roots)


def _resolution_discard_roots(
    selected: "List[Revision]",
    document: "Document",
    *,
    accept: bool,
) -> "List[_Element]":
    """Subtrees that `_resolve_one()` can detach for this selected set."""
    units, _orphans = _move_units(document)
    units_by_element = {
        id(element): unit for unit in units for element in unit.elements
    }
    roots: "List[_Element]" = []
    seen: "set[int]" = set()

    def add(root: "Optional[_Element]") -> None:
        if root is not None and id(root) not in seen:
            roots.append(root)
            seen.add(id(root))

    for revision in selected:
        node = revision._element  # noqa: SLF001
        if revision.is_paragraph_mark and _paragraph_mark_applies(
            node, accept=accept
        ):
            paragraph = _paragraph_of_mark(node)
            if paragraph is not None and _paragraph_has_content(paragraph):
                if _next_paragraph_sibling(paragraph) is not None:
                    add(paragraph.find(_PPR))
            elif (
                paragraph is not None
                and not _is_last_block_in_container(paragraph)
            ):
                add(paragraph)

        if revision.revision_type in ("move_from", "move_to"):
            unit = units_by_element.get(id(node))
            if unit is None:
                continue  # `_validate_moves()` already raised for this case.
            drop_site = unit.from_site if accept else unit.to_site
            for wrapper in drop_site.wrappers:
                add(wrapper)
            continue

        if revision.revision_type in ("row_insertion", "row_deletion"):
            keeps_row = (
                node.tag == _INS and accept
            ) or (node.tag == _DEL and not accept)
            if keeps_row:
                continue
            tr_pr = node.getparent()
            row = tr_pr.getparent() if tr_pr is not None else None
            table = row.getparent() if row is not None else None
            if row is None or row.tag != _TR or table is None or table.tag != _TBL:
                raise UnsupportedStructureError(
                    "malformed tracked-row revision cannot be resolved;"
                    " nothing was changed"
                )
            row_count = sum(1 for child in table if child.tag == _TR)
            add(table if row_count == 1 else row)
            continue

        removes_content = (node.tag == _INS and not accept) or (
            node.tag == _DEL and accept
        )
        if removes_content and not revision.is_paragraph_mark:
            add(node)

    return roots


def _validate_filtered_destructive_closure(
    discard_roots: "List[_Element]",
    selected: "List[Revision]",
    document: "Document",
    *,
    author: "Optional[str]",
) -> None:
    """Refuse a filter that would erase unselected revision markup."""
    if author is None:
        return
    selected_ids = {id(revision._element) for revision in selected}  # noqa: SLF001
    accounted_ids = set(selected_ids)
    selected_move_ids = {
        id(revision._element)  # noqa: SLF001
        for revision in selected
        if revision.revision_type in ("move_from", "move_to")
    }
    if selected_move_ids:
        units, _orphans = _move_units(document)
        for unit in units:
            if any(id(element) in selected_move_ids for element in unit.elements):
                accounted_ids.update(id(element) for element in unit.elements)

    conflicts: "set[str]" = set()
    for root in discard_roots:
        for node in root.iter():
            if node.tag not in _MARKUP_SCAN_TAGS or id(node) in accounted_ids:
                continue
            node_author = node.get(_AUTHOR) or ""
            conflicts.add(node_author or "<missing author>")
    if conflicts:
        raise UnsupportedStructureError(
            f"cannot resolve revisions with author={author!r}: a selected"
            " destructive revision would also consume unselected revision"
            f" markup (authors {sorted(conflicts)!r}); nothing was changed"
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
        # BODY-level w:sectPr (the final section) lives outside every block;
        # a tracked section-property change there must still be enumerated —
        # invisible-to-census markup would let accept_all report a clean
        # document while w:sectPrChange remains. The synthetic index -1
        # anchor marks "story level, not a block".
        for sect_pr in root.iter(_SECT_PR):
            parent = sect_pr.getparent()
            if parent is not None and parent.tag == _PPR:
                continue  # paragraph-level section break: reached via its block
            for node in _iter_revision_nodes(sect_pr, skip_text_boxes=False):
                yield Revision(
                    revision_type=_revision_type_of(node),
                    author=node.get(_AUTHOR) or "",
                    date=_parse_date(node.get(_DATE)),
                    text="",
                    story=story,
                    anchor=Anchor(story=story, index=-1, content_hash=content_hash("")),
                    is_paragraph_mark=False,
                    _element=node,
                    _document=document,
                )


def _revision_signature(document: "Document") -> "Tuple[_Element, ...]":
    """Identity signature for the document's currently enumerable revisions."""
    return tuple(revision._element for revision in _enumerate_revisions(document))


def _revision_snapshot(
    document: "Document",
) -> "Tuple[Tuple[Revision, ...], Tuple[_Element, ...]]":
    revisions = tuple(_enumerate_revisions(document))
    signature = tuple(revision._element for revision in revisions)
    return (
        tuple(
            replace(revision, _snapshot_signature=signature)
            for revision in revisions
        ),
        signature,
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


def _remaining_markup(document: "Document") -> dict:
    """{local-tag-name: count} of ALL revision markup left ANYWHERE — the
    full serialized space, including mc:Fallback branches (which traversal
    skips but a saved file still carries). The invariant oracle: after a
    successful `accept_all()`/`reject_all()` this is empty — "resolved"
    while markup remains anywhere would be false state.
    """
    counts: dict = {}
    for _story, root in _story_elements(document):
        for node in root.iter(*_MARKUP_SCAN_TAGS):
            name = node.tag.rsplit("}", 1)[-1]
            counts[name] = counts.get(name, 0) + 1
    return dict(sorted(counts.items()))


_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _fallback_markup(document: "Document") -> dict:
    """Revision markup hiding inside mc:Fallback branches — invisible to
    traversal (first-Choice-only) but alive in the saved file."""
    counts: dict = {}
    for _story, root in _story_elements(document):
        for fallback in root.iter(_MC_FALLBACK):
            for node in fallback.iter(*_MARKUP_SCAN_TAGS):
                name = node.tag.rsplit("}", 1)[-1]
                counts[name] = counts.get(name, 0) + 1
    return dict(sorted(counts.items()))


def _refuse_unaccounted_markup(document: "Document") -> None:
    """Refuse an unfiltered resolution that could not end clean: fallback
    markup, orphaned move markup, or markup enumeration cannot reach would
    all survive a 'successful' accept_all — reporting resolved-and-clean
    while markup remains is false state."""
    fallback = _fallback_markup(document)
    if fallback:
        raise UnsupportedStructureError(
            f"revision markup lives inside mc:AlternateContent fallback"
            f" branches ({fallback}); this package resolves the primary"
            " content only, so the saved file would still carry it. Nothing"
            " was changed — resolve this document in Word"
        )
    units, orphans = _move_units(document)
    if orphans:
        reasons = sorted({reason for _element, reason in orphans})
        raise UnsupportedStructureError(
            "the document carries move markup this package cannot pair"
            f" ({'; '.join(reasons)}); nothing was changed. Resolve it in"
            " Word instead"
        )
    marker_only = sorted(
        unit.name
        for unit in units
        if not (
            unit.from_site.wrappers
            or unit.from_site.mark_stamps
            or unit.to_site.wrappers
            or unit.to_site.mark_stamps
        )
    )
    if marker_only:
        raise UnsupportedStructureError(
            "the document carries complete marker-only move units with no"
            f" resolvable move content ({marker_only!r}); nothing was"
            " changed. Resolve it in Word instead"
        )
    accounted = {id(r._element) for r in _enumerate_revisions(document)}  # noqa: SLF001
    for unit in units:
        accounted.update(id(element) for element in unit.elements)
    unaccounted: dict = {}
    for _story, root in _story_elements(document):
        for node in root.iter(*_MARKUP_SCAN_TAGS):
            if id(node) not in accounted:
                name = node.tag.rsplit("}", 1)[-1]
                unaccounted[name] = unaccounted.get(name, 0) + 1
    if unaccounted:
        raise UnsupportedStructureError(
            f"the document carries revision markup this package cannot"
            f" enumerate ({dict(sorted(unaccounted.items()))}); nothing was"
            " changed. Resolve it in Word instead"
        )


def _is_in_story(node: "_Element", document: "Document") -> bool:
    root_ids = {id(root) for _story, root in _story_elements(document)}
    top = node
    while top.getparent() is not None:
        top = top.getparent()
    return id(top) in root_ids


def _comment_id(node: "_Element") -> int:
    raw = node.get(_ID)
    try:
        if raw is None:
            raise ValueError
        return int(raw)
    except (TypeError, ValueError):
        name = node.tag.rsplit("}", 1)[-1]
        raise UnsupportedStructureError(
            f"malformed {name} comment marker w:id={raw!r}; nothing was"
            " changed"
        ) from None


def _related_xml_root(
    document: "Document", reltype: str, label: str
) -> "Optional[_Element]":
    """Return one related live-XML root or refuse an ambiguous/opaque part."""
    try:
        part = document.part.part_related_by(reltype)
    except KeyError:
        return None
    except ValueError:
        raise UnsupportedStructureError(
            f"multiple {label} relationships make revision cleanup"
            " ambiguous; nothing was changed"
        ) from None
    root = getattr(part, "_element", None)
    if root is None:
        raise UnsupportedStructureError(
            f"{label} part is not loaded as live XML; nothing was changed"
        )
    return root


def _validate_comment_marker_ids(document: "Document") -> None:
    """Validate comment identifiers and cleanup parts before mutation."""
    for _story, root in _story_elements(document):
        for marker in root.iter(
            _COMMENT_RANGE_START, _COMMENT_RANGE_END, _COMMENT_REFERENCE
        ):
            _comment_id(marker)
        for comment in root.iter(_COMMENT):
            _comment_id(comment)

    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    comments_root = _related_xml_root(document, RT.COMMENTS, "comments")
    if comments_root is not None:
        for comment in comments_root.iter(_COMMENT):
            _comment_id(comment)

    # Cleanup also edits commentsExtended. Ensure it is live XML before a
    # destructive operation can discover that too late for atomic refusal.
    from docx.commentops import COMMENTS_EXTENDED_RELATIONSHIP_TYPE

    _related_xml_root(
        document,
        COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
        "commentsExtended",
    )


def _comment_ids_inside(node: "_Element") -> "List[int]":
    return [_comment_id(ref) for ref in node.iter(_COMMENT_REFERENCE)]


def _comment_range_ids_inside(node: "_Element") -> "set[int]":
    return {
        _comment_id(marker)
        for marker in node.iter(_COMMENT_RANGE_START, _COMMENT_RANGE_END)
    }


def _part_containing(node: "_Element", document: "Document"):
    top = node
    while top.getparent() is not None:
        top = top.getparent()
    package = document.part.package
    assert package is not None
    for part in package.iter_parts():
        if getattr(part, "_element", None) is top:
            return part
    return None


def _is_relationship_attribute(name: str) -> bool:
    return name.startswith(_REL_NS_PREFIX) or name == _OFFICE_REL_ID


def _relationship_ids_inside(node: "_Element", part) -> "set[str]":
    available = set(part.rels)
    return {
        value
        for descendant in node.iter()
        for name, value in descendant.attrib.items()
        if _is_relationship_attribute(name) and value in available
    }


def _drop_unreferenced_relationships(part, candidates: "set[str]") -> None:
    if not candidates:
        return
    root = getattr(part, "_element", None)
    if root is None:
        return
    referenced = {
        value
        for descendant in root.iter()
        for name, value in descendant.attrib.items()
        if _is_relationship_attribute(name) and value in candidates
    }
    for r_id in sorted(candidates - referenced):
        if r_id in part.rels:
            part.drop_rel(r_id)


def _note_id(node: "_Element") -> int:
    raw = node.get(_ID)
    try:
        if raw is None:
            raise ValueError
        return int(raw)
    except (TypeError, ValueError):
        name = node.tag.rsplit("}", 1)[-1]
        raise UnsupportedStructureError(
            f"malformed {name} w:id={raw!r}; nothing was changed"
        ) from None


def _note_ids_inside(node: "_Element", reference_tag: str) -> "set[int]":
    return {_note_id(reference) for reference in node.iter(reference_tag)}


def _note_reference_remains(
    document: "Document", reference_tag: str, note_id: int
) -> bool:
    return any(
        _note_id(reference) == note_id
        for _story, root in _story_elements(document)
        for reference in root.iter(reference_tag)
    )


def _validate_note_cleanup_graph(document: "Document") -> None:
    """Validate both note relationship graphs and all decimal identifiers."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    note_kinds = (
        (_FOOTNOTE_REFERENCE, _FOOTNOTE, RT.FOOTNOTES, "footnotes"),
        (_ENDNOTE_REFERENCE, _ENDNOTE, RT.ENDNOTES, "endnotes"),
    )
    roots = [root for _story, root in _story_elements(document)]
    comments_root = _related_xml_root(document, RT.COMMENTS, "comments")
    if comments_root is not None and all(
        comments_root is not existing for existing in roots
    ):
        roots.append(comments_root)
    for _reference_tag, body_tag, reltype, label in note_kinds:
        root = _related_xml_root(document, reltype, label)
        if root is None:
            continue
        if all(root is not existing for existing in roots):
            roots.append(root)
        for note in root.iter(body_tag):
            _note_id(note)
    for root in roots:
        for marker in root.iter(
            _COMMENT_RANGE_START, _COMMENT_RANGE_END, _COMMENT_REFERENCE
        ):
            _comment_id(marker)
        for comment in root.iter(_COMMENT):
            _comment_id(comment)
        for reference_tag, _body_tag, _reltype, _label in note_kinds:
            for reference in root.iter(reference_tag):
                _note_id(reference)


def _validate_cleanup_dependencies(
    document: "Document", discard_roots: "List[_Element]"
) -> None:
    """Preflight every side-part/resource cleanup a discard can trigger."""
    has_note_references = False
    comment_ids: "set[int]" = set()
    comment_range_ids: "set[int]" = set()
    for root in discard_roots:
        comment_ids.update(_comment_ids_inside(root))
        comment_range_ids.update(_comment_range_ids_inside(root))
        for reference_tag in (_FOOTNOTE_REFERENCE, _ENDNOTE_REFERENCE):
            if _note_ids_inside(root, reference_tag):
                has_note_references = True
        part = _part_containing(root, document)
        has_relationships = any(
            _is_relationship_attribute(name)
            for descendant in root.iter()
            for name in descendant.attrib
        )
        if has_relationships and part is None:
            raise UnsupportedStructureError(
                "cannot identify the package part containing relationship-"
                "bearing revision content; nothing was changed"
            )
    partial_comment_ids = sorted(comment_range_ids - comment_ids)
    if partial_comment_ids:
        raise UnsupportedStructureError(
            "revision cleanup would remove comment range markers while"
            f" leaving their reference marks live (comment ids"
            f" {partial_comment_ids}); nothing was changed"
        )
    if comment_ids:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT

        comments_root = _related_xml_root(document, RT.COMMENTS, "comments")
        _validate_comment_cleanup_graph(
            document,
            comment_ids,
            discard_roots,
            comments_root,
        )
        if comments_root is not None and any(
            True
            for reference_tag in (_FOOTNOTE_REFERENCE, _ENDNOTE_REFERENCE)
            for _reference in comments_root.iter(reference_tag)
        ):
            has_note_references = True
    if has_note_references:
        _validate_note_cleanup_graph(document)


def _validate_comment_cleanup_graph(
    document: "Document",
    comment_ids: "set[int]",
    discard_roots: "List[_Element]",
    comments_root: "Optional[_Element]",
) -> None:
    """Require one unambiguous reference and body for each removed comment."""
    def is_discarded(node: "_Element") -> bool:
        while node is not None:
            if any(node is root for root in discard_roots):
                return True
            node = node.getparent()
        return False

    story_roots = [root for _story, root in _story_elements(document)]
    for comment_id in sorted(comment_ids):
        references = [
            reference
            for root in story_roots
            for reference in root.iter(_COMMENT_REFERENCE)
            if _comment_id(reference) == comment_id
        ]
        if len(references) != 1 or not is_discarded(references[0]):
            raise UnsupportedStructureError(
                f"comment {comment_id} has {len(references)} reference marks;"
                " revision cleanup requires exactly one reference inside the"
                " discarded content. Nothing was changed"
            )

        bodies = (
            []
            if comments_root is None
            else [
                comment
                for comment in comments_root.iter(_COMMENT)
                if _comment_id(comment) == comment_id
            ]
        )
        if len(bodies) != 1:
            raise UnsupportedStructureError(
                f"comment {comment_id} has {len(bodies)} comment bodies;"
                " revision cleanup requires exactly one. Nothing was changed"
            )

        starts = [
            marker
            for root in story_roots
            for marker in root.iter(_COMMENT_RANGE_START)
            if _comment_id(marker) == comment_id
        ]
        ends = [
            marker
            for root in story_roots
            for marker in root.iter(_COMMENT_RANGE_END)
            if _comment_id(marker) == comment_id
        ]
        if len(starts) != len(ends) or len(starts) > 1:
            raise UnsupportedStructureError(
                f"comment {comment_id} has an ambiguous range-marker graph;"
                " nothing was changed"
            )


def _cleanup_note_bodies(
    document: "Document",
    footnote_ids: "set[int]",
    endnote_ids: "set[int]",
) -> None:
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    note_kinds = (
        (_FOOTNOTE_REFERENCE, _FOOTNOTE, RT.FOOTNOTES, footnote_ids),
        (_ENDNOTE_REFERENCE, _ENDNOTE, RT.ENDNOTES, endnote_ids),
    )
    for reference_tag, body_tag, reltype, note_ids in note_kinds:
        removable = {
            note_id
            for note_id in note_ids
            if not _note_reference_remains(document, reference_tag, note_id)
        }
        if not removable:
            continue
        try:
            notes_part = document.part.part_related_by(reltype)
        except KeyError:
            continue
        notes_root = getattr(notes_part, "_element", None)
        if notes_root is None:
            raise UnsupportedStructureError(
                "note part is not loaded as live XML; cleanup cannot"
                " continue without leaving orphaned note bodies"
            )
        for note in list(notes_root):
            if (
                note.tag == body_tag
                and note.get(_NOTE_TYPE) is None
                and _note_id(note) in removable
            ):
                _discard_element(note, document)


def _cleanup_comments_extended(
    document: "Document", paragraph_ids: "set[str]"
) -> None:
    if not paragraph_ids:
        return
    from docx.commentops import COMMENTS_EXTENDED_RELATIONSHIP_TYPE

    try:
        extended = document.part.part_related_by(
            COMMENTS_EXTENDED_RELATIONSHIP_TYPE
        )
    except KeyError:
        return
    root = getattr(extended, "_element", None)
    if root is None:  # prevalidated for public resolution paths
        raise UnsupportedStructureError(
            "commentsExtended is not loaded as live XML"
        )
    for entry in list(root):
        if (
            entry.tag == _W15_COMMENT_EX
            and entry.get(_W15_PARA_ID) in paragraph_ids
        ):
            root.remove(entry)


def _cleanup_comment_anchors(document: "Optional[Document]", comment_ids) -> None:
    """A resolution removed the run holding a comment's reference mark: also
    remove the now-orphaned range markers and the comment itself, exactly as
    Word does — half-deleted comments are silent corruption."""
    if document is None or not comment_ids:
        return
    wanted = set(comment_ids)
    for _story, root in _story_elements(document):
        for marker in list(root.iter(_COMMENT_RANGE_START, _COMMENT_RANGE_END)):
            if _comment_id(marker) in wanted:
                marker.getparent().remove(marker)
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        comments_part = document.part.part_related_by(RT.COMMENTS)
    except KeyError:
        return
    comments_root = comments_part._element  # noqa: SLF001
    paragraph_ids: "set[str]" = set()
    for comment in list(comments_root):
        if _comment_id(comment) not in wanted:
            continue
        paragraphs = comment.findall(_P)
        if paragraphs:
            para_id = paragraphs[-1].get(_W14_PARA_ID)
            if para_id is not None:
                paragraph_ids.add(para_id)
        _discard_element(comment, document)
    _cleanup_comments_extended(document, paragraph_ids)


def _discard_element(
    node: "_Element", document: "Optional[Document]"
) -> None:
    """Detach a subtree and prune resources that only that subtree used."""
    parent = node.getparent()
    if parent is None:
        raise UnsupportedStructureError(
            "cannot discard a detached revision subtree; nothing was changed"
        )
    comment_ids = _comment_ids_inside(node)
    footnote_ids = _note_ids_inside(node, _FOOTNOTE_REFERENCE)
    endnote_ids = _note_ids_inside(node, _ENDNOTE_REFERENCE)
    part = _part_containing(node, document) if document is not None else None
    relationship_ids = (
        _relationship_ids_inside(node, part) if part is not None else set()
    )

    parent.remove(node)

    if part is not None:
        _drop_unreferenced_relationships(part, relationship_ids)
    if document is not None:
        _cleanup_note_bodies(document, footnote_ids, endnote_ids)
        _cleanup_comment_anchors(document, comment_ids)


def _resolve_one(
    node: "_Element", *, accept: bool, document: "Optional[Document]" = None
) -> None:
    if node.getparent() is None:
        raise UnsupportedStructureError(
            "cannot resolve a stale Revision; reacquire it from"
            " document.revisions. Nothing was changed"
        )
    if document is not None and not _is_in_story(node, document):
        raise UnsupportedStructureError(
            "cannot resolve a stale Revision; reacquire it from"
            " document.revisions. Nothing was changed"
        )
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
    if removes_content:
        _discard_element(node, document)
        return
    if node.tag == _INS:
        _resolve_insertion(node, accept=accept)
    else:
        _resolve_deletion(node, accept=accept)


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
        parent = table.getparent()
        siblings = [
            child for child in parent if child.tag in (_P, _TBL) and child is not table
        ]
        if not siblings:
            from docx.oxml.parser import OxmlElement

            table.addprevious(OxmlElement("w:p"))
        _discard_element(table, document)
        return
    _discard_element(row, document)


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


_TRANSPARENT_BLOCK_MARKERS = frozenset(
    qn(tag)
    for tag in (
        "w:bookmarkStart",
        "w:bookmarkEnd",
        "w:commentRangeStart",
        "w:commentRangeEnd",
        "w:permStart",
        "w:permEnd",
        "w:proofErr",
        "w:moveFromRangeStart",
        "w:moveFromRangeEnd",
        "w:moveToRangeStart",
        "w:moveToRangeEnd",
        "w:customXmlInsRangeStart",
        "w:customXmlInsRangeEnd",
        "w:customXmlDelRangeStart",
        "w:customXmlDelRangeEnd",
        "w:customXmlMoveFromRangeStart",
        "w:customXmlMoveFromRangeEnd",
        "w:customXmlMoveToRangeStart",
        "w:customXmlMoveToRangeEnd",
    )
)


def _next_paragraph_sibling(paragraph: "_Element") -> "Optional[_Element]":
    node = paragraph.getnext()
    while node is not None:
        if node.tag == _P:
            return node
        if node.tag in (_TBL, qn("w:sdt"), _SECT_PR):
            # merging across a table or INTO a block-level content control
            # is not a paragraph join — hopping content past it would
            # silently reorder document text
            return None
        if node.tag in _TRANSPARENT_BLOCK_MARKERS:
            node = node.getnext()
            continue
        name = node.tag.rsplit("}", 1)[-1]
        raise UnsupportedStructureError(
            f"cannot resolve a paragraph-mark revision across the {name!r}"
            " block; its content boundary is not a safe paragraph join."
            " Nothing was changed"
        )
    return None


def _paragraph_mark_applies(node: "_Element", *, accept: bool) -> bool:
    is_deletion = node.tag in (_DEL, _MOVE_FROM)
    return (accept and is_deletion) or (not accept and not is_deletion)


def _preflight_paragraph_mark(node: "_Element", *, accept: bool) -> None:
    paragraph = _paragraph_of_mark(node)
    if (
        paragraph is not None
        and _paragraph_mark_applies(node, accept=accept)
        and _paragraph_has_content(paragraph)
    ):
        _next_paragraph_sibling(paragraph)


def _validate_paragraph_mark_joins(
    selected: "List[Revision]", document: "Document", *, accept: bool
) -> None:
    """Preflight every paragraph join, including compound move stamps."""
    marks = [
        revision._element  # noqa: SLF001
        for revision in selected
        if revision.is_paragraph_mark
    ]
    move_nodes = {
        id(revision._element)  # noqa: SLF001
        for revision in selected
        if revision.revision_type in ("move_from", "move_to")
    }
    if move_nodes:
        units, _orphans = _move_units(document)
        for unit in units:
            if any(id(element) in move_nodes for element in unit.elements):
                marks.extend(unit.from_site.mark_stamps)
                marks.extend(unit.to_site.mark_stamps)
    seen: "set[int]" = set()
    for mark in marks:
        if id(mark) in seen:
            continue
        seen.add(id(mark))
        _preflight_paragraph_mark(mark, accept=accept)


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
    applying_break_change = _paragraph_mark_applies(node, accept=accept)
    following = None
    if (
        paragraph is not None
        and applying_break_change
        and _paragraph_has_content(paragraph)
    ):
        # Resolve the join target before touching the revision stamp. This is
        # the last-line atomicity guard for direct/internal callers.
        following = _next_paragraph_sibling(paragraph)

    r_pr = node.getparent()
    r_pr.remove(node)
    if len(r_pr) == 0:
        r_pr.getparent().remove(r_pr)
    if paragraph is None or not applying_break_change:
        return
    if _paragraph_has_content(paragraph):
        if following is None:
            return  # nothing to join with; the content stays as a paragraph
        insert_at = 1 if following.find(_PPR) is not None else 0
        for child in reversed([c for c in paragraph if c.tag != _PPR]):
            following.insert(insert_at, child)
        paragraph.getparent().remove(paragraph)
        return
    if not _is_last_block_in_container(paragraph) and paragraph.getparent() is not None:
        paragraph.getparent().remove(paragraph)
