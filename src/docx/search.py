"""Normalized text search over visibility-complete traversal.

`find_text` matches the way people (and models) actually quote documents:
smart quotes, dashes, exotic spaces and case differences are normalized away
(`normalize_text`), and matches assemble across Word's fragmented runs and
across paragraph boundaries. The returned |Span| is the pivotal object of the
editing surface: it maps a visible-text interval back to the concrete
`w:t` text atoms that hold it, carries a stable block anchor, and is the
receiver of the safe replace operations (`Span.replace`).

Search space and block identity are shared with `docx.story` (one walker
defines both), so a span's anchor always agrees with the outline.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterator, List, Optional, Sequence, Tuple

from docx import _clock, _textatoms
from docx._guard import check_install
from docx._normalize import normalize_text  # noqa: F401 - public re-export
from docx._transaction import rollback_on_error
from docx.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import OP_COMMENT, _refuse_if_protected
from docx.story import (
    VIEWS,
    Anchor,
    _first_choice_children,
    _iter_block_elements,
    _story_elements,
    content_hash,
)

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

check_install()

_T = _textatoms.T
_DEL_TEXT = _textatoms.DEL_TEXT
_INSTR_TEXT = _textatoms.INSTR_TEXT
_TEXT_TAGS = _textatoms.TEXT_TAGS
_R = _textatoms.R
_INS = qn("w:ins")
_DEL = qn("w:del")
_MOVE_FROM = qn("w:moveFrom")
_MOVE_TO = qn("w:moveTo")
_SDT = qn("w:sdt")
_HYPERLINK = qn("w:hyperlink")
_TXBX = qn("w:txbxContent")
_RPR = qn("w:rPr")
_XML_SPACE = qn("xml:space")
_FLD_SIMPLE = qn("w:fldSimple")
_FLD_CHAR = qn("w:fldChar")
_FLD_CHAR_TYPE = qn("w:fldCharType")


@dataclass
class _Atom:
    """One text node in search space, with its structural context.

    `fixed_text` marks a synthetic atom such as a tab, break, or no-break
    hyphen. ``barrier`` marks visible run XML with no safe text model; it
    contributes an object-replacement sentinel solely to split search text.
    Neither kind can be edited directly through a span.
    """

    element: "_Element"
    tag: str
    story: str
    block_index: int
    paragraph: "Optional[_Element]"
    run: "Optional[_Element]"
    sdt: "Optional[_Element]"  # innermost content control, if any
    in_insert: bool
    in_delete: bool
    in_text_box: bool
    hyperlink: "Optional[_Element]" = None  # innermost w:hyperlink, if any
    in_field: bool = False
    fixed_text: Optional[str] = None
    barrier: bool = False

    @property
    def in_hyperlink(self) -> bool:
        return self.hyperlink is not None

    @property
    def text(self) -> str:
        if self.fixed_text is not None:
            return self.fixed_text
        return self.element.text or ""

    @property
    def is_synthetic(self) -> bool:
        return self.fixed_text is not None


@dataclass(frozen=True)
class _FreshnessCensus:
    """One story census shared by a synchronous replacement batch."""

    atoms: "Tuple[_Atom, ...]"
    by_element: "dict[int, _Atom]"
    positions: "dict[int, int]"


_CONTEXT_SCOPE_TAGS = frozenset(
    (_INS, _DEL, _MOVE_FROM, _MOVE_TO, _SDT, _HYPERLINK, _FLD_SIMPLE, _TXBX)
)


def _atom_context_signature(atom: _Atom) -> tuple:
    """Snapshot the edit-sensitive structural context of one atom.

    Element identity catches reparenting into a different run, paragraph, or
    edit-sensitive scope without coupling freshness to absolute sibling
    indexes. Attributes catch retargeted hyperlinks and re-authored revisions;
    SDT properties catch lock/placeholder/type changes. Complex-field
    membership is carried separately because its begin/end markers are
    siblings rather than ancestors of the result text.
    """
    from lxml import etree

    scopes = []
    current = atom.element.getparent()
    while current is not None:
        if current.tag in _CONTEXT_SCOPE_TAGS:
            extra = b""
            if current.tag == _SDT:
                sdt_pr = current.find(qn("w:sdtPr"))
                if sdt_pr is not None:
                    extra = etree.tostring(sdt_pr, with_tail=False)
            scopes.append(
                (
                    current,
                    current.tag,
                    tuple(sorted(current.attrib.items())),
                    extra,
                )
            )
        current = current.getparent()
    scopes.reverse()
    return (
        atom.story,
        atom.tag,
        atom.paragraph,
        atom.run,
        atom.in_insert,
        atom.in_delete,
        atom.in_text_box,
        atom.in_field,
        atom.fixed_text,
        atom.barrier,
        tuple(scopes),
    )


def _collect_block_atoms(
    story: str,
    block_index: int,
    element: "_Element",
    *,
    skip_text_boxes: bool,
    in_txbx: bool,
    field_depth: "Optional[List[int]]" = None,
) -> "Iterator[_Atom]":
    """Text atoms under one block element, in document order.

    Complex fields (`w:fldChar` begin…separate…end run sequences) are tracked
    with a depth counter as the walk passes them in document order: every
    atom between begin and end — instruction AND cached result — carries
    `in_field=True`, as do `w:fldSimple` descendants. Pass a shared
    `field_depth` cell when iterating consecutive blocks of one story:
    multi-paragraph fields (every Word TOC) keep their begin…end state OPEN
    across block boundaries. Tracked moves flag their atoms deletion-like
    (`w:moveFrom`) / insertion-like (`w:moveTo`).
    """
    if field_depth is None:
        field_depth = [0]

    def walk(node, paragraph, run, sdt, in_ins, in_del, hlink, in_box, in_fld):
        for child in _first_choice_children(node):
            tag = child.tag
            if tag == _TXBX and skip_text_boxes:
                continue
            if tag == _FLD_CHAR:
                fld_type = child.get(_FLD_CHAR_TYPE)
                if fld_type == "begin":
                    field_depth[0] += 1
                elif fld_type == "end" and field_depth[0] > 0:
                    field_depth[0] -= 1
                continue
            if _textatoms.is_direct_run_child(child):
                projection = _textatoms.project_run_child(child)
                if not projection.text and not projection.barrier:
                    continue
                yield _Atom(
                    element=child,
                    tag=tag,
                    story=story,
                    block_index=block_index,
                    paragraph=paragraph,
                    run=run,
                    sdt=sdt,
                    in_insert=in_ins,
                    in_delete=in_del,
                    in_text_box=in_box,
                    hyperlink=hlink,
                    in_field=in_fld or field_depth[0] > 0,
                    fixed_text=(
                        None if tag in _TEXT_TAGS else projection.text
                    ),
                    barrier=projection.barrier,
                )
                continue
            yield from walk(
                child,
                child if tag == qn("w:p") else paragraph,
                child if tag == _R else run,
                child if tag == _SDT else sdt,
                in_ins or tag in (_INS, _MOVE_TO),
                in_del or tag in (_DEL, _MOVE_FROM),
                child if tag == _HYPERLINK else hlink,
                in_box or tag == _TXBX,
                in_fld or tag == _FLD_SIMPLE,
            )

    paragraph = element if element.tag == qn("w:p") else None
    yield from walk(element, paragraph, None, None, False, False, None, in_txbx, False)


def _story_atoms(document: "Document", story_name: str, root: "_Element") -> "List[_Atom]":
    atoms: "List[_Atom]" = []
    # ONE field-depth cell for the whole story: a field opened in one block
    # stays open into the next (TOC shape) — resetting per block would be a
    # false-state hole
    field_depth = [0]
    for kind, index, element, _in_sdt, in_txbx in _iter_block_elements(story_name, root):
        skip_boxes = kind == "paragraph"  # text-box paragraphs are their own blocks
        atoms.extend(
            _collect_block_atoms(
                story_name, index, element, skip_text_boxes=skip_boxes,
                in_txbx=in_txbx, field_depth=field_depth,
            )
        )
    return atoms


def _include_atom(atom: _Atom, view: str) -> bool:
    if atom.is_synthetic:
        if view == "current":
            return not atom.in_delete
        if view == "original":
            return not atom.in_insert
        return True
    if view == "current":
        # in_delete covers moveFrom sources (live w:t that vanishes on accept)
        return atom.tag == _T and not atom.in_delete
    if view == "original":
        # nothing inside a pending insertion (incl. moveTo destinations)
        # existed in the original — including deletions nested within it
        return not atom.in_insert and atom.tag in (_T, _DEL_TEXT)
    return True  # "all"


def _assemble(atoms: "Sequence[_Atom]") -> Tuple[str, List[Tuple[int, int]]]:
    """(text, per-char (atom-index, offset) map); paragraph boundaries appear
    as a single "\\n" whose map entry has offset -1."""
    pieces: List[str] = []
    mapping: List[Tuple[int, int]] = []
    previous: "Optional[_Atom]" = None
    for atom_idx, atom in enumerate(atoms):
        if previous is not None and atom.paragraph is not previous.paragraph:
            pieces.append("\n")
            mapping.append((atom_idx, -1))
        for offset in range(len(atom.text)):
            mapping.append((atom_idx, offset))
        pieces.append(atom.text)
        previous = atom
    return "".join(pieces), mapping


def _normalized_with_map(value: str) -> Tuple[str, List[int]]:
    """Normalized text plus, per normalized char, the source index in `value`.

    Exact port of the reference algorithm: translate, collapse whitespace
    runs, casefold — keeping a char-level map back to the raw string.
    """
    from docx._normalize import NORMALIZE_CHARS

    chars: List[str] = []
    mapping: List[int] = []
    for index, char in enumerate(value):
        translated = char.translate(NORMALIZE_CHARS)
        if translated == "":
            continue
        for out in translated:
            if out.isspace():
                out = " "
            # casefold can EXPAND one char to several ('ß' -> 'ss'); every
            # expanded char must map back to the same source index or the
            # map desynchronizes and matching crashes past the expansion
            for folded in out.casefold():
                chars.append(folded)
                mapping.append(index)
    compact: List[str] = []
    compact_map: List[int] = []
    previous_space = False
    for char, original in zip(chars, mapping):
        if char == " ":
            if previous_space:
                continue
            previous_space = True
        else:
            previous_space = False
        compact.append(char)
        compact_map.append(original)
    return "".join(compact), compact_map


@dataclass(frozen=True)
class ReplaceResult:
    """Outcome of a `Span.replace` call.

    `deleted_text` and `inserted_text` cover the whole span for an untracked replace but only
    the affix-trimmed middle for a tracked one, so comparing them against `Span.text` is
    wrong for tracked edits.
    """

    story: str
    deleted_text: str
    inserted_text: str
    tracked: bool
    revision_ids: Tuple[int, ...]
    preserved_structure: bool = False
    preserved_revision_ids: Tuple[int, ...] = ()

    def to_dict(self) -> dict:
        return {
            "schema": "paper_replace",
            "version": 1,
            "story": self.story,
            "deleted_text": self.deleted_text,
            "inserted_text": self.inserted_text,
            "tracked": self.tracked,
            "revision_ids": list(self.revision_ids),
            "preserved_structure": self.preserved_structure,
            "preserved_revision_ids": list(self.preserved_revision_ids),
        }


@dataclass(frozen=True)
class _TextAssignment:  # pyright: ignore[reportUnusedClass]
    """One preflighted text-only mutation for a replacement plan."""

    element: "_Element"
    before: str
    after: str


@dataclass
class Span:
    """A visible-text interval mapped back to its concrete text atoms.

    Spans hold live references into the document tree and go stale when the
    underlying text changes; every operation revalidates first and raises
    |TargetNotFoundError| on staleness.
    """

    text: str
    story: str
    anchor: Anchor
    in_insert: bool
    in_delete: bool
    in_content_control: bool
    in_text_box: bool
    in_field: bool
    crosses_paragraphs: bool
    _document: "Document" = field(repr=False)
    _atoms: "List[_Atom]" = field(repr=False)
    _start_offset: int = field(repr=False)  # into first atom's text
    _end_offset: int = field(repr=False)  # exclusive, into last atom's text
    _norm_start: int = field(repr=False)  # position in the story's normalized text
    _consumed: bool = field(default=False, repr=False)  # set by tracked replace
    _context_signatures: "Tuple[tuple, ...]" = field(init=False, repr=False)
    _atom_sequence: "Tuple[_Element, ...]" = field(init=False, repr=False)
    _sequence_view: "Optional[str]" = field(init=False, repr=False)
    _view: str = field(init=False, repr=False)
    _freshness_census: "Optional[_FreshnessCensus]" = field(
        default=None, init=False, repr=False
    )

    def __post_init__(self) -> None:
        self._context_signatures = tuple(
            _atom_context_signature(atom) for atom in self._atoms
        )
        # Constructors outside search build current-view spans. Search replaces
        # this fallback with the full captured interval, including atoms hidden
        # by the selected view.
        self._atom_sequence = tuple(atom.element for atom in self._atoms)
        self._sequence_view = "current"
        self._view = "current"

    # -- comments ---------------------------------------------------------

    def comment(
        self,
        text: str,
        *,
        author: str,
        initials: Optional[str] = None,
        date: Optional[dt.datetime] = None,
    ) -> "object":
        """Anchor a new comment to exactly this span's text, and return the upstream `Comment`.

        Reach for this over `Document.add_comment` when the anchor must match exact text rather
        than whole runs. Splits boundary runs, and creates `/word/comments.xml` on first use.
        Only the main document story carries comments. Refuses a protected document, a stale
        span, and a locked or data-bound control surface.
        """
        if not author:
            raise ValueError("author is required")
        _validate_xml_characters(text, argument="text")
        _validate_xml_characters(author, argument="author")
        if initials is not None:
            _validate_xml_characters(initials, argument="initials")
        if date is not None and not isinstance(date, dt.datetime):
            raise TypeError("date must be a datetime or None")
        _refuse_if_protected(
            self._document, "anchor a comment", operation_class=OP_COMMENT
        )
        main_story = next(
            story
            for story, root in _story_elements(self._document)
            if root is self._document.element
        )
        if self.story != main_story:
            raise UnsupportedStructureError(
                "comments anchor in the main document story"
                f" (span is in {self.story})"
            )
        self._validate_fresh()
        for atom in self._atoms:
            if atom.run is None:
                raise UnsupportedStructureError(
                    "span text is not inside runs; cannot anchor a comment"
                )
        from docx.commentops import _preflight_comment_add, _preflight_comment_range

        _preflight_comment_add(self._document)
        runs = []
        for atom in self._atoms:
            if atom.run is not None and not any(atom.run is run for run in runs):
                runs.append(atom.run)
        if runs:
            _preflight_comment_range(
                self._document, runs[0], runs[-1], operation="anchor a comment"
            )
        with rollback_on_error(self._document, self):
            self._isolate_edge_runs()
            runs = []
            for atom in self._atoms:
                if not any(existing is atom.run for existing in runs):
                    runs.append(atom.run)
            comments = self._document.comments
            comment = comments.add_comment(
                text=text, author=author, initials=initials or ""
            )
            comment._comment_elm.date = (  # noqa: SLF001
                date if date is not None else _clock.now()
            )
            runs[0].insert_comment_range_start_above(comment.comment_id)
            runs[-1].insert_comment_range_end_and_reference_below(
                comment.comment_id
            )
            return comment

    def _isolate_edge_runs(self) -> None:
        """Split boundary runs so this span's runs hold EXACTLY its text.

        Semantically neutral (text, order and formatting unchanged): run
        content BEFORE the first matched text node — including the split-off
        text prefix — moves to a new preceding run, and content AFTER the
        last matched node (plus the split-off tail) moves to a new following
        run, each with a clone of its source run's `rPr`. Needed so
        element-level anchors like comment range marks wrap the span text,
        not whole runs, without ever reordering visible content.
        """
        import copy

        def split_before(run, element, prefix_text: str) -> None:
            preceding = []
            for child in run:
                if child is element:
                    break
                if child.tag != _RPR:
                    preceding.append(child)
            if not preceding and not prefix_text:
                return
            new_run = OxmlElement("w:r")
            rpr = run.find(_RPR)
            if rpr is not None:
                new_run.append(copy.deepcopy(rpr))
            for child in preceding:
                run.remove(child)
                new_run.append(child)
            if prefix_text:
                new_run.add_t(prefix_text)
            run.addprevious(new_run)

        def split_after(run, element, tail_text: str) -> None:
            following = []
            seen = False
            for child in run:
                if seen and child.tag != _RPR:
                    following.append(child)
                if child is element:
                    seen = True
            if not following and not tail_text:
                return
            new_run = OxmlElement("w:r")
            rpr = run.find(_RPR)
            if rpr is not None:
                new_run.append(copy.deepcopy(rpr))
            if tail_text:
                new_run.add_t(tail_text)
            for child in following:
                run.remove(child)
                new_run.append(child)
            run.addnext(new_run)

        first = self._atoms[0]
        if first.run is not None:
            prefix = first.text[: self._start_offset] if self._start_offset > 0 else ""
            split_before(first.run, first.element, prefix)
            if prefix:
                _set_preserved_text(first.element, first.text[self._start_offset :])
                if len(self._atoms) == 1:
                    self._end_offset -= self._start_offset
                self._start_offset = 0
        last = self._atoms[-1]
        if last.run is not None:
            tail = last.text[self._end_offset :] if self._end_offset < len(last.text) else ""
            split_after(last.run, last.element, tail)
            if tail:
                _set_preserved_text(last.element, last.text[: self._end_offset])

    # -- narrowing --------------------------------------------------------

    def _synthetic_positions(self) -> "set":
        """Char positions in this span's text contributed by tab/break atoms
        (paragraph separators count too — they are equally unwritable)."""
        positions = set()
        cursor = 0
        for atom_index, text in self._in_span_pieces():
            if atom_index is None or self._atoms[atom_index].is_synthetic:
                positions.update(range(cursor, cursor + len(text)))
            cursor += len(text)
        return positions

    def _narrow_to_change(self, new_text: str) -> "Optional[Tuple[Span, str]]":
        """A sub-span covering only the changed region, or None if the trim
        cannot shrink this span (caller falls through to normal validation).

        In the kept prefix/suffix, a whitespace character in `new_text`
        aligns with an existing tab/break: callers cannot write `\\t`,
        so "Section 4. Termination" against "Section 3.<TAB>Termination"
        keeps the document's tab and changes only the "3" — matching is
        normalized, documents keep their original characters.
        """
        synthetic = self._synthetic_positions()
        old = self.text

        def aligned(old_pos: int, new_char: str) -> bool:
            return old[old_pos] == new_char or (
                old_pos in synthetic and new_char.isspace()
            )

        prefix_len = 0
        limit = min(len(old), len(new_text))
        while prefix_len < limit and aligned(prefix_len, new_text[prefix_len]):
            prefix_len += 1
        suffix_len = 0
        while (
            suffix_len < len(old) - prefix_len
            and suffix_len < len(new_text) - prefix_len
            and aligned(len(old) - suffix_len - 1, new_text[len(new_text) - suffix_len - 1])
        ):
            suffix_len += 1
        if prefix_len == 0 and suffix_len == 0:
            return None
        changed_old = self.text[prefix_len : len(self.text) - suffix_len]
        changed_new = new_text[prefix_len : len(new_text) - suffix_len]
        if not changed_old and not changed_new:
            return None
        # map the changed char range onto the atom slice (paragraph
        # separators are None pieces: a change touching one cannot narrow —
        # validation will refuse it as a cross-paragraph change)
        target_start = prefix_len
        target_end = len(self.text) - suffix_len
        position = 0
        start_idx = end_idx = None
        start_off = end_off = 0
        for atom_index, text in self._in_span_pieces():
            length = len(text)
            base = 0
            if atom_index == 0:
                base = self._start_offset
            if start_idx is None and position + length > target_start:
                if atom_index is None:
                    return None  # change begins on a paragraph separator
                start_idx = atom_index
                start_off = base + (target_start - position)
            if position + length >= target_end:
                if atom_index is None and end_idx is None:
                    return None  # change ends on a paragraph separator
                end_idx = atom_index if atom_index is not None else end_idx
                end_off = base + (target_end - position) if atom_index is not None else end_off
                break
            position += length
        if start_idx is None or end_idx is None:
            return None  # zero-length change at an edge; let validation decide
        sub_atoms = self._atoms[start_idx : end_idx + 1]
        sub_span = Span(
            text=changed_old,
            story=self.story,
            anchor=Anchor(
                story=self.story,
                index=self.anchor.index,
                content_hash=content_hash(changed_old),
            ),
            in_insert=any(a.in_insert for a in sub_atoms),
            in_delete=any(a.in_delete or a.tag == _DEL_TEXT for a in sub_atoms),
            in_content_control=any(a.sdt is not None for a in sub_atoms),
            in_text_box=any(a.in_text_box for a in sub_atoms),
            in_field=any(a.in_field for a in sub_atoms),
            crosses_paragraphs=any(
                a.paragraph is not sub_atoms[0].paragraph for a in sub_atoms
            ),
            _document=self._document,
            _atoms=list(sub_atoms),
            _start_offset=start_off,
            _end_offset=end_off,
            _norm_start=self._norm_start,
        )
        sub_span._view = self._view
        sub_span._sequence_view = self._sequence_view
        sub_span._freshness_census = self._freshness_census
        sequence_positions = {
            id(element): index for index, element in enumerate(self._atom_sequence)
        }
        sequence_start = sequence_positions.get(id(sub_atoms[0].element))
        sequence_end = sequence_positions.get(id(sub_atoms[-1].element))
        if (
            sequence_start is not None
            and sequence_end is not None
            and sequence_start <= sequence_end
        ):
            sub_span._atom_sequence = self._atom_sequence[
                sequence_start : sequence_end + 1
            ]
        return sub_span, changed_new

    # -- validation -------------------------------------------------------

    def _in_span_pieces(self) -> "List[Tuple[Optional[int], str]]":
        """(atom-index-or-None, text) pieces reconstituting this span's text.

        `None` entries are the synthetic "\\n" paragraph separators the search
        text carries between atoms of different paragraphs — part of the span
        TEXT but belonging to no atom; every offset computation must account
        for them or cross-paragraph spans desynchronize.
        """
        pieces: "List[Tuple[Optional[int], str]]" = []
        previous: "Optional[_Atom]" = None
        for index, atom in enumerate(self._atoms):
            if previous is not None and atom.paragraph is not previous.paragraph:
                pieces.append((None, "\n"))
            start = self._start_offset if index == 0 else 0
            end = self._end_offset if index == len(self._atoms) - 1 else len(atom.text)
            pieces.append((index, atom.text[start:end]))
            previous = atom
        return pieces

    def _current_slice(self) -> str:
        """The span's text as the tree holds it right now (paragraph
        boundaries render as the same "\\n" the capture used)."""
        return "".join(text for _, text in self._in_span_pieces())

    def _validate_fresh(self) -> None:
        if self._consumed:
            raise TargetNotFoundError(
                "span was consumed by a tracked or structure-preserving"
                " replace; re-find the text"
            )
        if self._current_slice() != self.text:
            raise TargetNotFoundError(
                f"span is stale: expected {self.text!r} at its anchor"
                f" ({self.anchor.to_dict()}) but the document has changed"
            )
        # Text equality is insufficient: moving the same w:t into a field,
        # hyperlink, content control, or revision changes whether/how it may be
        # edited. Compare against one shared census for a synchronous batch,
        # or rebuild the story census for an independently used span. Detached
        # and alternate-branch-switched atoms naturally disappear from it.
        story_root = next(
            (
                root
                for story_name, root in _story_elements(self._document)
                if story_name == self.story
            ),
            None,
        )
        if story_root is None:
            raise TargetNotFoundError(
                "span is stale: its story part was removed from the document"
            )
        census = self._freshness_census
        if census is None:
            current_atoms = tuple(
                _story_atoms(self._document, self.story, story_root)
            )
            current_by_element = {
                id(atom.element): atom for atom in current_atoms
            }
        else:
            current_atoms = census.atoms
            current_by_element = census.by_element
        for captured, expected in zip(self._atoms, self._context_signatures):
            current = current_by_element.get(id(captured.element))
            if current is None or _atom_context_signature(current) != expected:
                raise TargetNotFoundError(
                    "span is stale: its field, hyperlink, content-control,"
                    " revision, or containing structure has changed or was"
                    " removed"
                )
        sequence_atoms = (
            current_atoms
            if self._sequence_view is None
            else [
                atom
                for atom in current_atoms
                if _include_atom(atom, self._sequence_view)
            ]
        )
        sequence_positions = (
            census.positions
            if census is not None and self._sequence_view is None
            else {
                id(atom.element): index
                for index, atom in enumerate(sequence_atoms)
            }
        )
        expected_first = self._atom_sequence[0]
        expected_last = self._atom_sequence[-1]
        sequence_start = sequence_positions.get(id(expected_first))
        sequence_end = sequence_positions.get(id(expected_last))
        current_sequence = (
            ()
            if sequence_start is None
            or sequence_end is None
            or sequence_start > sequence_end
            else tuple(
                atom.element
                for atom in sequence_atoms[sequence_start : sequence_end + 1]
            )
        )
        if len(current_sequence) != len(self._atom_sequence) or any(
            current is not expected
            for current, expected in zip(current_sequence, self._atom_sequence)
        ):
            raise TargetNotFoundError(
                "span is stale: the text-atom sequence inside its captured"
                " interval has changed"
            )

    def _validate_replaceable(self, *, validate_bookmarks: bool = True) -> None:
        for atom in self._atoms:
            if atom.is_synthetic:
                detail = (
                    "unmodeled visible run content"
                    if atom.barrier
                    else "a tab or line break, or a no-break hyphen"
                )
                raise UnsupportedStructureError(
                    f"span crosses {detail}; replace the text"
                    " segments on either side individually"
                )
            if atom.tag == _DEL_TEXT or atom.in_delete:
                raise UnsupportedStructureError(
                    "span includes tracked-deleted text; resolve the revision"
                    " first or target visible text only"
                )
            if atom.tag == _INSTR_TEXT:
                raise UnsupportedStructureError(
                    "span includes field-instruction text; editing field code"
                    " internals is not supported"
                )
            if atom.in_field:
                raise UnsupportedStructureError(
                    "span lies inside a field result (TOC entry, page number,"
                    " date, cross-reference, …); Word regenerates field results"
                    " on update, so the edit would silently vanish"
                )
        if self.crosses_paragraphs:
            raise BoundaryViolationError(
                "span crosses a paragraph boundary; character-level replace is"
                " same-paragraph only (use docx.blocks for clause-level edits)"
            )
        scopes = {id(atom.sdt) if atom.sdt is not None else None for atom in self._atoms}
        if len(scopes) > 1:
            raise BoundaryViolationError(
                "span crosses a content-control boundary; edit inside or"
                " outside the control, not across it"
            )
        from docx.controls import (
            _refuse_control_write_restrictions,
            _validate_span_surface_edit,
        )

        controls = []
        for atom in self._atoms:
            current = atom.element.getparent()
            while current is not None:
                if current.tag == _SDT and not any(current is item for item in controls):
                    controls.append(current)
                current = current.getparent()
        for control in controls:
            _validate_span_surface_edit(control)
            _refuse_control_write_restrictions(control)
        # both modes: an untracked replace crossing a hyperlink boundary
        # would silently move text into or out of the link
        link_scopes = {
            id(atom.hyperlink) if atom.hyperlink is not None else None
            for atom in self._atoms
        }
        if len(link_scopes) > 1:
            raise BoundaryViolationError(
                "span crosses a hyperlink boundary; edit the linked text and"
                " the surrounding text separately"
            )
        hollowed = (
            _hollowed_bookmarks(self._atoms, self._end_offset)
            if validate_bookmarks
            else []
        )
        if hollowed:
            raise UnsupportedStructureError(
                f"replacing this span would hollow out bookmark(s) {hollowed}"
                " — the targets of REF/PAGEREF cross-references and TOC"
                " entries; narrow the span to inside the bookmark, or remove"
                " the bookmark deliberately first"
            )

    # -- replace ----------------------------------------------------------

    def replace(
        self,
        new_text: str,
        *,
        tracked: bool = False,
        author: Optional[str] = None,
        date: Optional[dt.datetime] = None,
        preserve_revision: bool = False,
    ) -> ReplaceResult:
        """Replace this span's text, preserving run formatting.

        Untracked, this edits in place: untouched runs retain their run properties, the
        replacement takes the start run's formatting, and the span stays usable. Tracked, it
        emits a minimal `w:del`/`w:ins` pair and consumes the span.

        `preserve_revision=True` permits a current-view span wholly owned by one existing
        `w:ins` to be corrected without changing its id, author, date, or accept/reject meaning.
        Outside revision markup it behaves like an ordinary untracked edit. The result reports
        the retained insertion in `preserved_revision_ids`; `revision_ids` remains reserved for
        newly authored revisions. Refuses mixed, nested, moved, stale, protected, field,
        control, bookmark, hyperlink, and paragraph-boundary structures before mutation.
        """
        return self._replace(
            new_text,
            tracked=tracked,
            author=author,
            date=date,
            preserve_revision=preserve_revision,
            use_transaction=self._freshness_census is None,
        )

    def _replace(
        self,
        new_text: str,
        *,
        tracked: bool,
        author: Optional[str],
        date: Optional[dt.datetime],
        preserve_revision: bool,
        use_transaction: bool,
    ) -> ReplaceResult:
        """Implementation shared with the already-transactional batch path.

        `use_transaction` is consumed by replacement plans that need multiple
        assignments; ordinary and tracked paths retain their current behavior.
        """
        if tracked and preserve_revision:
            raise ValueError(
                "tracked=True cannot be combined with preserve_revision=True"
            )
        if tracked and not author:
            raise ValueError("author is required when tracked=True")
        if tracked:
            # the w:ins/w:del identity attributes are stamped AFTER mutation
            # begins; malformed values must refuse before anything changes
            _validate_xml_characters(author, argument="author")
            if date is not None and not isinstance(date, dt.datetime):
                raise TypeError("date must be a datetime or None")
        _validate_writable_text(new_text, argument="new_text")
        _refuse_if_protected(self._document, "replace text")
        self._validate_fresh()
        if any(atom.is_synthetic for atom in self._atoms) or self.crosses_paragraphs:
            # spans matched ACROSS a tab/break/paragraph boundary may still
            # edit safely when the actual change lies within one segment:
            # narrow to the changed region; if the change itself crosses a
            # break or boundary, validation below refuses as before
            narrowed = self._narrow_to_change(new_text)
            if narrowed is not None:
                sub_span, sub_new = narrowed
                return sub_span._replace(
                    sub_new,
                    tracked=tracked,
                    author=author,
                    date=date,
                    preserve_revision=preserve_revision,
                    use_transaction=use_transaction,
                )
        preservation_noop = preserve_revision and new_text == self.text
        self._validate_replaceable(validate_bookmarks=not preservation_noop)
        preserved_revision_ids = _preserved_insertion_ids(
            self, authorize=preserve_revision
        )
        if not tracked:
            for atom in self._atoms:
                if atom.in_insert and not preserve_revision:
                    raise UnsupportedStructureError(
                        "span intersects a pending tracked insertion; an"
                        " untracked edit there would silently rewrite text the"
                        " revision history attributes to its author — accept or"
                        " reject the revision first, or use tracked=True"
                    )
        placeholder_sdts = _placeholder_controls_of(self._atoms)
        if placeholder_sdts:
            if tracked:
                raise UnsupportedStructureError(
                    "span lies in a form control still showing PLACEHOLDER"
                    " prompt text; a tracked replacement would fabricate a"
                    " deletion of text that was never real content — fill the"
                    " control untracked (or via docx.controls) first"
                )
            for sdt in placeholder_sdts:
                prompt = _placeholder_prompt_text(sdt)
                if self.text != prompt:
                    raise UnsupportedStructureError(
                        "span covers only part of a placeholder prompt;"
                        " replacing it would promote the leftover prompt text"
                        " to real content — replace the whole prompt"
                        f" ({prompt!r}) or use docx.controls.set_control_value"
                    )
        if preserve_revision and new_text == self.text:
            return ReplaceResult(
                story=self.story,
                deleted_text=self.text,
                inserted_text=new_text,
                tracked=False,
                revision_ids=(),
                preserved_revision_ids=preserved_revision_ids,
            )
        if tracked:
            result = self._tracked_replace(new_text, author=author, date=date)  # type: ignore[arg-type]
        else:
            result = self._plain_replace(
                new_text, preserved_revision_ids=preserved_revision_ids
            )
        for sdt in placeholder_sdts:
            _clear_placeholder_state(sdt)
        return result

    def _plain_replace(
        self,
        new_text: str,
        *,
        preserved_revision_ids: "Tuple[int, ...]" = (),
    ) -> ReplaceResult:
        first, last = self._atoms[0], self._atoms[-1]
        if first is last:
            text = first.text
            _set_preserved_text(
                first.element,
                text[: self._start_offset] + new_text + text[self._end_offset :],
            )
        else:
            _set_preserved_text(
                first.element, first.text[: self._start_offset] + new_text
            )
            for atom in self._atoms[1:-1]:
                _set_preserved_text(atom.element, "")
            _set_preserved_text(last.element, last.text[self._end_offset :])
        result = ReplaceResult(
            story=self.story,
            deleted_text=self.text,
            inserted_text=new_text,
            tracked=False,
            revision_ids=(),
            preserved_revision_ids=preserved_revision_ids,
        )
        self.text = new_text
        self._end_offset = self._start_offset + len(new_text)
        del self._atoms[1:]
        self._context_signatures = tuple(
            _atom_context_signature(atom) for atom in self._atoms
        )
        self._atom_sequence = (first.element,)
        self._sequence_view = "current"
        self._view = "current"
        return result

    def _tracked_replace(
        self, new_text: str, *, author: str, date: Optional[dt.datetime]
    ) -> ReplaceResult:
        from docx.oxml.revision import CT_RunTrackChange

        if self.story == "word/comments.xml":
            raise UnsupportedStructureError(
                "tracked changes inside comment text are not representable;"
                " edit the comment untracked or reply instead"
            )
        for atom in self._atoms:
            if atom.run is None:
                raise UnsupportedStructureError(
                    "matched text is not inside a run; cannot anchor a revision"
                )
        # layered revisions: a span straddling a pending w:ins would emit a
        # w:del claiming inserted text was base-document content — fabricated
        # history that corrupts reject/original views. Two shapes are safe:
        # fully inside ONE w:ins (nests correctly), and the SAME author
        # extending their own insertion where the span starts in base
        # text and ends in/at their insertion (their inserted text is simply
        # removed, never re-marked as a base-text deletion).
        enclosing = [_enclosing_insertion(atom.element) for atom in self._atoms]
        scopes = {id(e) if e is not None else None for e in enclosing}
        extends_own_insertion = False
        if len(scopes) > 1:
            non_none = [e for e in enclosing if e is not None]
            all_own = all(
                e.tag == _INS and (e.get(_W_AUTHOR) or "") == author for e in non_none
            )
            inside_seen = False
            contiguous_tail = True
            for e in enclosing:
                if e is not None:
                    inside_seen = True
                elif inside_seen:
                    contiguous_tail = False
                    break
            if not (all_own and enclosing[0] is None and contiguous_tail):
                raise UnsupportedStructureError(
                    "span overlaps a pending tracked insertion it cannot layer"
                    " over (different author, a tracked move, or base text"
                    " following the insertion); accept or reject the existing"
                    " revision first, or target text inside or outside it"
                )
            extends_own_insertion = True
        if self.text == new_text:
            raise TargetNotFoundError(
                "replacement equals the existing text; nothing to change"
            )
        prefix_len, suffix_len = _common_affix_lengths(self.text, new_text)
        first, last = self._atoms[0], self._atoms[-1]
        if first is not last:
            # kept characters must never cross run boundaries (they would
            # silently adopt another run's formatting): clamp the trim so the
            # prefix stays in the first run and the suffix in the last
            prefix_len = min(prefix_len, len(first.text) - self._start_offset)
            suffix_len = min(suffix_len, self._end_offset)
        old_mid = self.text[prefix_len : len(self.text) - suffix_len]
        new_mid = new_text[prefix_len : len(new_text) - suffix_len]
        stamp = date if date is not None else _clock.now()
        next_id = _next_revision_id(self._document)
        rpr = first.run.find(_RPR) if first.run is not None else None

        # deleted text must keep each source run's own formatting inside
        # w:del (or reject would restore it with the wrong rPr); computed
        # before mutation while atom texts are intact
        span_pieces = []
        for i, atom in enumerate(self._atoms):
            start = self._start_offset if i == 0 else 0
            end = self._end_offset if i == len(self._atoms) - 1 else len(atom.text)
            span_pieces.append((atom.run, atom.text[start:end]))
        deleted_pieces = _pieces_in_range(
            span_pieces, prefix_len, len(self.text) - suffix_len
        )
        # the inserted text renders with the first CHANGED run's formatting
        ins_source_run = deleted_pieces[0][0] if deleted_pieces else first.run
        ins_rpr = (
            ins_source_run.find(_RPR) if ins_source_run is not None else None
        )

        # -- everything validated; mutate ---------------------------------
        before = first.text[: self._start_offset]
        after = last.text[self._end_offset :]
        prefix = self.text[:prefix_len]
        suffix = self.text[len(self.text) - suffix_len :] if suffix_len else ""

        _set_preserved_text(first.element, before + prefix)
        same_run = first.run is last.run
        for atom in self._atoms[1:-1]:
            _set_preserved_text(atom.element, "")
        tail_in_new_run = first is last or same_run
        if not tail_in_new_run:
            _set_preserved_text(last.element, suffix + after)
        elif first is not last:
            _set_preserved_text(last.element, "")

        # content of the first run that FOLLOWS the matched text node (a
        # tab, a second w:t, a drawing) must move after the emitted revision
        # or the visible order scrambles
        trailing = _run_content_after(first.run, first.element)
        moved_run = None
        if trailing:
            moved_run = OxmlElement("w:r")
            if rpr is not None:
                import copy as _copy

                moved_run.append(_copy.deepcopy(rpr))
            for element in trailing:
                first.run.remove(element)
                moved_run.append(element)

        revision_ids: "List[int]" = []
        insert_point = first.run
        if extends_own_insertion:
            # pieces inside the author's own pending insertion are removed
            # outright (deleting your own unaccepted insertion leaves no
            # mark, exactly as Word behaves); only base-text pieces get w:del
            deleted_pieces = [
                (source_run, piece)
                for source_run, piece in deleted_pieces
                if source_run is None or _enclosing_insertion(source_run) is None
            ]
        if old_mid and deleted_pieces:
            del_elm = CT_RunTrackChange.new("w:del", next_id, author, stamp)
            for source_run, piece in deleted_pieces:
                source_rpr = source_run.find(_RPR) if source_run is not None else None
                del_elm.add_tracked_run(piece, source_rpr, deleted=True)
            insert_point.addnext(del_elm)
            insert_point = del_elm
            revision_ids.append(next_id)
            next_id += 1
        if new_mid:
            ins_elm = CT_RunTrackChange.new("w:ins", next_id, author, stamp)
            ins_elm.add_tracked_run(new_mid, ins_rpr, deleted=False)
            insert_point.addnext(ins_elm)
            insert_point = ins_elm
            revision_ids.append(next_id)
        if tail_in_new_run and (suffix + after):
            tail_run = _new_text_run(suffix + after, rpr)
            insert_point.addnext(tail_run)
            insert_point = tail_run
        if moved_run is not None:
            insert_point.addnext(moved_run)
        if extends_own_insertion:
            # an own-ins whose text this edit fully consumed is now an empty
            # phantom revision; drop it rather than enumerate a ghost
            for own_ins in {e for e in enclosing if e is not None}:
                if not "".join(own_ins.itertext()) and not own_ins.xpath(
                    ".//w:drawing | .//w:object"
                ):
                    own_ins.getparent().remove(own_ins)

        result = ReplaceResult(
            story=self.story,
            deleted_text=old_mid,
            inserted_text=new_mid,
            tracked=True,
            revision_ids=tuple(revision_ids),
        )
        # span state after a tracked replace is complex; force a fresh find
        self.text = new_text
        self._consumed = True
        return result


def _new_text_run(text: str, rpr: "Optional[_Element]"):
    """A new `w:r` built through the oxml layer, with `rpr` cloned in."""
    import copy

    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    run.add_t(text)
    return run


def _pieces_in_range(pieces, start: int, end: int):
    """(run, text-slice) pieces intersected with [start, end) over their
    concatenation, consecutive same-run pieces merged."""
    out = []
    position = 0
    for run, text in pieces:
        lo = max(start, position)
        hi = min(end, position + len(text))
        if hi > lo:
            fragment = text[lo - position : hi - position]
            if out and out[-1][0] is run:
                out[-1] = (run, out[-1][1] + fragment)
            else:
                out.append((run, fragment))
        position += len(text)
    return out


def _enclosing_insertion(element: "_Element") -> "Optional[_Element]":
    """The nearest insertion-like (`w:ins`/`w:moveTo`) ancestor, or None."""
    node = element.getparent()
    while node is not None:
        if node.tag in (_INS, _MOVE_TO):
            return node
        node = node.getparent()
    return None


def _revision_ancestors(element: "_Element") -> "Tuple[_Element, ...]":
    """Run-level revision wrappers containing `element`, nearest first."""
    revisions: "List[_Element]" = []
    node = element.getparent()
    while node is not None:
        if node.tag in (_INS, _DEL, _MOVE_FROM, _MOVE_TO):
            revisions.append(node)
        node = node.getparent()
    return tuple(revisions)


def _preserved_insertion_ids(
    span: Span, *, authorize: bool
) -> "Tuple[int, ...]":
    """Validate and identify the one insertion explicitly kept by `span`."""
    if not authorize:
        return ()

    selected_wrappers = [
        _enclosing_insertion(atom.element)
        for atom in span._atoms  # pyright: ignore[reportPrivateUsage]
    ]
    if all(wrapper is None for wrapper in selected_wrappers):
        return ()
    if span._view != "current":  # pyright: ignore[reportPrivateUsage]
        raise UnsupportedStructureError(
            "preserve_revision requires a span selected with view='current'"
        )
    if any(wrapper is None for wrapper in selected_wrappers):
        raise UnsupportedStructureError(
            "span mixes base text and insertion-owned text; preserve one"
            " existing insertion at a time"
        )
    wrapper = selected_wrappers[0]
    assert wrapper is not None
    if wrapper.tag != _INS:
        raise UnsupportedStructureError(
            "tracked moves cannot be edited in place; resolve the move first"
        )
    if any(candidate is not wrapper for candidate in selected_wrappers[1:]):
        raise UnsupportedStructureError(
            "span crosses multiple tracked insertion wrappers; preserve one"
            " existing insertion at a time"
        )
    for element in span._atom_sequence:  # pyright: ignore[reportPrivateUsage]
        # the captured range covers atoms hidden by the current view too, and
        # a plain replace writes only the SELECTED ones — so every element in
        # it must be owned by exactly this insertion and nothing else
        ancestors = _revision_ancestors(element)
        if len(ancestors) != 1 or ancestors[0] is not wrapper:
            raise UnsupportedStructureError(
                "span contains nested or mixed revision history that cannot be"
                " edited while preserving one insertion"
            )
    value = wrapper.get(_W_ID)
    try:
        revision_id = int(value) if value is not None else None
    except ValueError:
        revision_id = None
    if revision_id is None:
        raise UnsupportedStructureError(
            "the preserved insertion has a missing or non-decimal w:id"
        )
    return (revision_id,)


def _run_content_after(run: "Optional[_Element]", element: "_Element"):
    """`run`'s content children positioned after `element`, in order."""
    if run is None:
        return []
    trailing = []
    past = False
    for child in run:
        if past and child.tag != _RPR:
            trailing.append(child)
        if child is element:
            past = True
    return trailing


def _next_revision_id(document: "Document") -> int:
    """One above the highest revision id anywhere in the document.

    Uses plain element iteration — story roots like `w:footnotes` carry no
    registered oxml class, so prefix-based `.xpath()` is unavailable there.
    """
    w_id = qn("w:id")
    revision_tags = (
        _INS, _DEL, _MOVE_FROM, _MOVE_TO,
        qn("w:moveFromRangeStart"), qn("w:moveToRangeStart"),
        qn("w:rPrChange"), qn("w:pPrChange"), qn("w:tblPrChange"),
        qn("w:sectPrChange"), qn("w:cellIns"), qn("w:cellDel"), qn("w:cellMerge"),
    )
    highest = 0
    for _, root in _story_elements(document):
        for node in root.iter(*revision_tags):
            value = node.get(w_id)
            try:
                highest = max(highest, int(value)) if value else highest
            except ValueError:
                continue
    return highest + 1


_W_AUTHOR = qn("w:author")
_CONTROL_CHARS = ("\n", "\r", "\t", "\x0b", "\x0c")
_SDT_PR = qn("w:sdtPr")
_SHOWING_PLC_HDR = qn("w:showingPlcHdr")
_R_STYLE = qn("w:rStyle")
_BOOKMARK_START = qn("w:bookmarkStart")
_BOOKMARK_END = qn("w:bookmarkEnd")
_W_ID = qn("w:id")
_W_NAME = qn("w:name")


def _validate_xml_characters(value: str, *, argument: str) -> None:
    """Reject characters XML 1.0 cannot represent before any mutation."""
    invalid = sorted(
        {
            f"U+{ord(char):04X}"
            for char in value
            if not (
                ord(char) in (0x09, 0x0A, 0x0D)
                or 0x20 <= ord(char) <= 0xD7FF
                or 0xE000 <= ord(char) <= 0xFFFD
                or 0x10000 <= ord(char) <= 0x10FFFF
            )
        }
    )
    if invalid:
        raise ValueError(
            f"{argument} contains character(s) XML 1.0 cannot represent:"
            f" {invalid!r}"
        )


def _validate_writable_text(value: str, *, argument: str) -> None:
    """Refuse control characters in text written into `w:t` (programmer error).

    A raw newline/tab inside `w:t` is not a break in Word, but this package's
    own read-back would render it as one — the classic verified-but-false
    structure. The search side refuses spans crossing `w:br`/`w:tab` for the
    mirror reason: replacing across one would silently drop it.
    """
    _validate_xml_characters(value, argument=argument)
    found = sorted({c for c in value if c in _CONTROL_CHARS})
    if found:
        raise ValueError(
            f"{argument} contains control character(s) {found!r}: Word does"
            " not render them as breaks inside w:t — pass separate"
            " paragraphs (or a rich block list) instead"
        )


def _placeholder_controls_of(atoms: "Sequence[_Atom]"):
    """Distinct content controls of `atoms` still showing placeholder text."""
    controls = []
    for atom in atoms:
        if atom.sdt is None or any(existing is atom.sdt for existing in controls):
            continue
        sdt_pr = atom.sdt.find(_SDT_PR)
        if sdt_pr is not None and sdt_pr.find(_SHOWING_PLC_HDR) is not None:
            controls.append(atom.sdt)
    return controls


def _placeholder_prompt_text(sdt: "_Element") -> str:
    """The full prompt text a placeholder-showing control displays."""
    content = sdt.find(qn("w:sdtContent"))
    if content is None:
        return ""
    return "".join(node.text or "" for node in content.iter(_T))


def _in_nested_sdt(node: "_Element", outer_sdt: "_Element") -> bool:
    parent = node.getparent()
    while parent is not None and parent is not outer_sdt:
        if parent.tag == _SDT:
            return True
        parent = parent.getparent()
    return False


def _clear_placeholder_state(sdt: "_Element") -> None:
    """The control was really filled: drop `w:showingPlcHdr` and the
    `PlaceholderText` run style so Word stops treating it as empty.

    Nested inner controls keep their own placeholder state — only THIS
    control was filled."""
    sdt_pr = sdt.find(_SDT_PR)
    if sdt_pr is not None:
        showing = sdt_pr.find(_SHOWING_PLC_HDR)
        if showing is not None:
            sdt_pr.remove(showing)
    for r_style in sdt.findall(f".//{_R_STYLE}"):
        if r_style.get(qn("w:val")) == "PlaceholderText" and not _in_nested_sdt(
            r_style, sdt
        ):
            r_style.getparent().remove(r_style)


def _hollowed_bookmarks(atoms: "Sequence[_Atom]", end_offset: int) -> "List[str]":
    """Names of non-point bookmarks the replace would silently EMPTY.

    Character-accurate: the replace empties every middle atom, plus the last
    atom when the span consumes it entirely (`end_offset == len(last.text)`);
    the first atom always keeps its prefix + the replacement. A bookmark is
    hollowed when ALL of its text lives in elements that get emptied — the
    replacement lands outside its marker pair. Point bookmarks (`_GoBack`)
    hold no text and are transparent by design.
    """
    paragraph = atoms[0].paragraph
    if paragraph is None:
        return []
    emptied_ids = {id(atom.element) for atom in atoms[1:-1]}
    if len(atoms) > 1 and end_offset >= len(atoms[-1].text):
        emptied_ids.add(id(atoms[-1].element))
    if not emptied_ids:
        return []
    stream = list(paragraph.iter())
    starts = {}
    hollowed = []
    for position, node in enumerate(stream):
        if node.tag == _BOOKMARK_START:
            starts[node.get(_W_ID)] = (position, node.get(_W_NAME) or "")
        elif node.tag == _BOOKMARK_END:
            entry = starts.get(node.get(_W_ID))
            if entry is None:
                continue
            start_pos, name = entry
            if name == "_GoBack":
                continue
            text_elements = [
                inner
                for inner in stream[start_pos + 1 : position]
                if inner.tag in (_T, _DEL_TEXT) and (inner.text or "")
            ]
            if text_elements and all(
                id(inner) in emptied_ids for inner in text_elements
            ):
                hollowed.append(name)
    return hollowed


def _set_preserved_text(element: "_Element", text: str) -> None:
    element.text = text
    if text[:1].isspace() or text[-1:].isspace():
        element.set(_XML_SPACE, "preserve")
    elif _XML_SPACE in element.attrib:
        del element.attrib[_XML_SPACE]


def _common_affix_lengths(old: str, new: str) -> Tuple[int, int]:
    prefix = 0
    limit = min(len(old), len(new))
    while prefix < limit and old[prefix] == new[prefix]:
        prefix += 1
    suffix = 0
    while (
        suffix < len(old) - prefix
        and suffix < len(new) - prefix
        and old[len(old) - suffix - 1] == new[len(new) - suffix - 1]
    ):
        suffix += 1
    return prefix, suffix


def _spans_for_story(
    document: "Document",
    story_name: str,
    atoms: "List[_Atom]",
    needle_norm: str,
    *,
    all_atoms: "Sequence[_Atom]",
    view: str,
) -> "List[Span]":
    raw_text, char_map = _assemble(atoms)
    normalized, norm_map = _normalized_with_map(raw_text)
    all_atom_positions = {
        id(atom.element): index for index, atom in enumerate(all_atoms)
    }
    spans: "List[Span]" = []
    search_from = 0
    while True:
        found_at = normalized.find(needle_norm, search_from)
        if found_at < 0:
            break
        search_from = found_at + len(needle_norm)
        raw_start = norm_map[found_at]
        raw_end = norm_map[found_at + len(needle_norm) - 1] + 1
        # trim paragraph-separator chars at the edges (they hold no atom text)
        while raw_start < raw_end and char_map[raw_start][1] == -1:
            raw_start += 1
        while raw_end > raw_start and char_map[raw_end - 1][1] == -1:
            raw_end -= 1
        if raw_start >= raw_end:
            continue
        start_atom_idx, start_offset = char_map[raw_start]
        end_atom_idx, end_offset = char_map[raw_end - 1]
        crosses = any(
            char_map[i][1] == -1 for i in range(raw_start, raw_end)
        ) or (
            atoms[start_atom_idx].paragraph is not atoms[end_atom_idx].paragraph
        )
        span_atoms = atoms[start_atom_idx : end_atom_idx + 1]
        start_block = atoms[start_atom_idx].block_index
        span = Span(
            text=raw_text[raw_start:raw_end],
            story=story_name,
            anchor=Anchor(
                story=story_name,
                index=start_block,
                content_hash=content_hash(raw_text[raw_start:raw_end]),
            ),
            in_insert=any(a.in_insert for a in span_atoms),
            in_delete=any(a.in_delete or a.tag == _DEL_TEXT for a in span_atoms),
            in_content_control=any(a.sdt is not None for a in span_atoms),
            in_text_box=any(a.in_text_box for a in span_atoms),
            in_field=any(a.in_field for a in span_atoms),
            crosses_paragraphs=crosses,
            _document=document,
            _atoms=list(span_atoms),
            _start_offset=start_offset,
            _end_offset=end_offset + 1,
            _norm_start=found_at,
        )
        sequence_start = all_atom_positions[id(span_atoms[0].element)]
        sequence_end = all_atom_positions[id(span_atoms[-1].element)]
        span._atom_sequence = tuple(
            atom.element for atom in all_atoms[sequence_start : sequence_end + 1]
        )
        span._sequence_view = None
        span._view = view
        spans.append(span)
    return spans


def _near_distance(span_norm_start: int, near_positions: "List[int]") -> float:
    if not near_positions:
        return float("inf")
    return min(abs(span_norm_start - position) for position in near_positions)


def find_text(
    document: "Document",
    needle: str,
    *,
    nth: Optional[int] = None,
    near: Optional[str] = None,
    story: Optional[str] = None,
    view: str = "current",
) -> "List[Span]":
    """Every span of `needle` in `document`, normalized matching.

    `story` limits the search to one story part (e.g. "word/document.xml");
    `near` ranks matches by distance to the nearest occurrence of `near`'s
    normalized text in the same story; `nth` (1-based) then selects a single
    match. Matching assembles across fragmented runs and across paragraph
    boundaries (a paragraph break matches a single space in the needle).
    """
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}, got {view!r}")
    needle_norm = normalize_text(needle)
    if not needle_norm.strip():
        return []
    near_norm = normalize_text(near) if near else None

    all_spans: "List[Tuple[float, int, Span]]" = []
    order = 0
    for story_name, root in _story_elements(document):
        if story is not None and story_name != story:
            continue
        all_atoms = _story_atoms(document, story_name, root)
        atoms = [atom for atom in all_atoms if _include_atom(atom, view)]
        if not atoms:
            continue
        spans = _spans_for_story(
            document,
            story_name,
            atoms,
            needle_norm,
            all_atoms=all_atoms,
            view=view,
        )
        if not spans:
            continue
        near_positions: "List[int]" = []
        if near_norm:
            raw_text, _ = _assemble(atoms)
            normalized, _ = _normalized_with_map(raw_text)
            start = normalized.find(near_norm)
            while start >= 0:
                near_positions.append(start)
                start = normalized.find(near_norm, start + 1)
        for span in spans:
            distance = (
                _near_distance(span._norm_start, near_positions) if near_norm else 0.0
            )
            all_spans.append((distance, order, span))
            order += 1
    all_spans.sort(key=lambda item: (item[0], item[1]))
    matches = [span for _, _, span in all_spans]
    if nth is not None:
        if nth < 1 or nth > len(matches):
            return []
        return [matches[nth - 1]]
    return matches


@dataclass(frozen=True)
class ReplaceAllResult:
    """Outcome of a `replace_all` call: per-match results, never silent."""

    replaced_count: int
    results: Tuple[ReplaceResult, ...]
    refused: Tuple[dict, ...]  # {story, anchor, error, message} per refusal

    def to_dict(self) -> dict:
        return {
            "schema": "paper_replace_all",
            "version": 1,
            "replaced_count": self.replaced_count,
            "results": [result.to_dict() for result in self.results],
            "refused": list(self.refused),
        }


def replace_all(
    document: "Document",
    needle: str,
    new_text: str,
    *,
    story: Optional[str] = None,
    view: str = "current",
    tracked: bool = False,
    author: Optional[str] = None,
    date: Optional[dt.datetime] = None,
    preserve_revision: bool = False,
) -> ReplaceAllResult:
    """Replace every match of `needle` in one pass, and return a `ReplaceAllResult`.

    One scan finds all matches, then replacements apply in reverse document order within each
    story, so no pending match shifts. A refusal on one match is recorded in `refused` and the
    rest proceed; a stale span aborts the batch instead, rolling back every replacement already
    applied. Matches already equal to `new_text` are skipped. `preserve_revision` forwards the
    same existing-insertion contract to every match without adding a transaction per match.
    """
    from docx.errors import PaperRefusal

    if tracked and preserve_revision:
        raise ValueError(
            "tracked=True cannot be combined with preserve_revision=True"
        )
    if tracked and not author:
        raise ValueError("author is required when tracked=True")
    _validate_writable_text(new_text, argument="new_text")
    _refuse_if_protected(document, "replace text")
    spans = [
        span
        for span in find_text(document, needle, story=story, view=view)
        if span.text != new_text
    ]
    by_story: "dict[str, List[Span]]" = {}
    for span in spans:
        by_story.setdefault(span.story, []).append(span)
    results: "List[ReplaceResult]" = []
    refused: "List[dict]" = []
    story_roots = dict(_story_elements(document))
    for story_name, story_spans in by_story.items():
        root = story_roots.get(story_name)
        if root is None:
            raise TargetNotFoundError(
                f"story {story_name!r} was removed before replacement"
            )
        atoms = tuple(_story_atoms(document, story_name, root))
        census = _FreshnessCensus(
            atoms=atoms,
            by_element={id(atom.element): atom for atom in atoms},
            positions={
                id(atom.element): index for index, atom in enumerate(atoms)
            },
        )
        for span in story_spans:
            span._freshness_census = census
    try:
        with rollback_on_error(document):
            for story_name in sorted(by_story):
                ordered = sorted(
                    by_story[story_name], key=lambda s: s._norm_start, reverse=True
                )
                for span in ordered:
                    try:
                        results.append(
                            span.replace(
                                new_text,
                                tracked=tracked,
                                author=author,
                                date=date,
                                preserve_revision=preserve_revision,
                            )
                        )
                    except TargetNotFoundError:
                        # A stale target invalidates the captured batch. The outer
                        # transaction restores replacements already applied.
                        raise
                    except PaperRefusal as exc:
                        refused.append(
                            {
                                "story": span.story,
                                "anchor": span.anchor.to_dict(),
                                "error": type(exc).__name__,
                                "message": str(exc),
                            }
                        )
    finally:
        for story_spans in by_story.values():
            for span in story_spans:
                span._freshness_census = None
    return ReplaceAllResult(
        replaced_count=len(results), results=tuple(results), refused=tuple(refused)
    )


def find_one(
    document: "Document",
    needle: str,
    *,
    nth: Optional[int] = None,
    near: Optional[str] = None,
    story: Optional[str] = None,
    view: str = "current",
) -> Span:
    """The single span matching `needle`, or a typed refusal.

    Zero matches raise `TargetNotFoundError`. Two or more raise `AmbiguousTargetError`:
    `nth` and `story` narrow the set, while `near` only ranks `find_text` results and never
    reduces them.
    """
    matches = find_text(document, needle, nth=nth, near=near, story=story, view=view)
    if not matches:
        raise TargetNotFoundError(f"no match for {needle!r} in any story part")
    if len(matches) > 1:
        locations = ", ".join(
            f"{span.story}#{span.anchor.index}" for span in matches[:5]
        )
        raise AmbiguousTargetError(
            f"{len(matches)} matches for {needle!r} (at {locations}"
            f"{', …' if len(matches) > 5 else ''}); disambiguate with nth=,"
            " near=, or story="
        )
    return matches[0]
