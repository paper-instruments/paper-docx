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

import datetime as dt
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, List, Optional, Sequence, Tuple

from docx.errors import UnsupportedStructureError
from docx.oxml.ns import qn
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
_R = qn("w:r")
_RPR = qn("w:rPr")
_PPR = qn("w:pPr")
_P = qn("w:p")
_AUTHOR = qn("w:author")
_DATE = qn("w:date")

#: tag -> revision_type for everything Document.revisions enumerates
_REVISION_TYPES = {
    _INS: "insertion",
    _DEL: "deletion",
    _MOVE_FROM: "move_from",
    _MOVE_TO: "move_to",
}
for _change_tag in (
    "w:rPrChange", "w:pPrChange", "w:tblPrChange", "w:tcPrChange",
    "w:trPrChange", "w:sectPrChange", "w:numberingChange",
    "w:cellIns", "w:cellDel", "w:cellMerge",
):
    _REVISION_TYPES[qn(_change_tag)] = "format_change"
del _change_tag

#: the only revision types accept()/reject() know how to resolve correctly
RESOLVABLE_TYPES = frozenset({"insertion", "deletion"})


def _node_text(node: "_Element") -> str:
    pieces: List[str] = []
    for child in node.iter():
        if child.tag in (_T, _DEL_TEXT):
            pieces.append(child.text or "")
    return "".join(pieces)


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

    `revision_type` is one of "insertion" | "deletion" (resolvable),
    "move_from" | "move_to" | "format_change" (enumerated and counted, but
    resolution is refused — v0.1 knows how to SEE these, not how to apply
    them; claiming otherwise would report false state).
    """

    revision_type: str
    author: str
    date: Optional[dt.datetime]
    text: str
    story: str
    anchor: Anchor
    is_paragraph_mark: bool
    _element: "_Element"

    @property
    def is_resolvable(self) -> bool:
        return self.revision_type in RESOLVABLE_TYPES

    def _refuse_unresolvable(self, verb: str) -> None:
        if not self.is_resolvable:
            raise UnsupportedStructureError(
                f"cannot {verb} a {self.revision_type!r} revision: tracked"
                " moves and formatting changes are enumerated but not yet"
                " resolvable (resolve them in Word, or a later paper-docx"
                " version)"
            )

    def accept(self) -> None:
        """Apply this change to the document."""
        self._refuse_unresolvable("accept")
        _resolve_one(self._element, accept=True)

    def reject(self) -> None:
        """Undo this change, restoring the pre-change content."""
        self._refuse_unresolvable("reject")
        _resolve_one(self._element, accept=False)

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
        resolved = 0
        # content revisions first, then paragraph marks (mark resolution can
        # remove whole paragraphs and must see post-content state)
        ordered = sorted(selected, key=lambda item: item.is_paragraph_mark)
        for revision in ordered:
            _resolve_one(revision._element, accept=accept)  # noqa: SLF001
            resolved += 1
        self._items = tuple(_enumerate_revisions(self._document))
        return resolved

    def to_dict(self) -> dict:
        return {
            "schema": "paper_revisions",
            "version": 2,  # v2: move/format_change types + census (v0.1 H1-H3)
            "revisions": [revision.to_dict() for revision in self._items],
            "remaining_unsupported": self.remaining_unsupported(),
        }


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
                yield Revision(
                    revision_type=_REVISION_TYPES[node.tag],
                    author=node.get(_AUTHOR) or "",
                    date=_parse_date(node.get(_DATE)),
                    text=_node_text(node),
                    story=story,
                    anchor=block_anchor,
                    is_paragraph_mark=_is_paragraph_mark_revision(node),
                    _element=node,
                )


def _resolve_one(node: "_Element", *, accept: bool) -> None:
    if node.getparent() is None:
        return  # already resolved via an enclosing operation
    if _is_paragraph_mark_revision(node):
        _resolve_paragraph_mark(node, accept=accept)
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
        _unwrap(node)


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
    is_deletion = node.tag == _DEL
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
