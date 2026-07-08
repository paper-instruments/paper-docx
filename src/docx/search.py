"""Normalized text search over visibility-complete traversal (paper-docx).

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

from docx import _clock
from docx._normalize import normalize_text  # noqa: F401 - public re-export
from docx.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
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

_T = qn("w:t")
_DEL_TEXT = qn("w:delText")
_INSTR_TEXT = qn("w:instrText")
_TEXT_TAGS = (_T, _DEL_TEXT, _INSTR_TEXT)
#: tab/break elements contribute fixed characters to the search text so
#: needles can match across them; spans containing them refuse to replace
_SYNTHETIC_TAGS = {qn("w:tab"): "\t", qn("w:br"): "\n", qn("w:cr"): "\n"}
_R = qn("w:r")
_INS = qn("w:ins")
_DEL = qn("w:del")
_SDT = qn("w:sdt")
_HYPERLINK = qn("w:hyperlink")
_TXBX = qn("w:txbxContent")
_RPR = qn("w:rPr")
_XML_SPACE = qn("xml:space")


@dataclass
class _Atom:
    """One text node in search space, with its structural context.

    `fixed_text` marks a synthetic atom: a `w:tab`/`w:br`/`w:cr` that
    contributes a fixed character to the search text so needles can match
    across it, but that can never be edited through a span (replace refuses).
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
    in_hyperlink: bool
    fixed_text: Optional[str] = None

    @property
    def text(self) -> str:
        if self.fixed_text is not None:
            return self.fixed_text
        return self.element.text or ""

    @property
    def is_synthetic(self) -> bool:
        return self.fixed_text is not None


def _collect_block_atoms(
    story: str,
    block_index: int,
    element: "_Element",
    *,
    skip_text_boxes: bool,
    in_txbx: bool,
) -> "Iterator[_Atom]":
    """Text atoms under one block element, in document order."""

    def walk(node, paragraph, run, sdt, in_ins, in_del, in_hlink, in_box):
        for child in _first_choice_children(node):
            tag = child.tag
            if tag == _TXBX and skip_text_boxes:
                continue
            if tag in _TEXT_TAGS or tag in _SYNTHETIC_TAGS:
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
                    in_hyperlink=in_hlink,
                    fixed_text=_SYNTHETIC_TAGS.get(tag),
                )
                continue
            yield from walk(
                child,
                child if tag == qn("w:p") else paragraph,
                child if tag == _R else run,
                child if tag == _SDT else sdt,
                in_ins or tag == _INS,
                in_del or tag == _DEL,
                in_hlink or tag == _HYPERLINK,
                in_box or tag == _TXBX,
            )

    paragraph = element if element.tag == qn("w:p") else None
    yield from walk(element, paragraph, None, None, False, False, False, in_txbx)


def _story_atoms(document: "Document", story_name: str, root: "_Element") -> "List[_Atom]":
    atoms: "List[_Atom]" = []
    for kind, index, element, _in_sdt, in_txbx in _iter_block_elements(story_name, root):
        skip_boxes = kind == "paragraph"  # text-box paragraphs are their own blocks
        atoms.extend(
            _collect_block_atoms(
                story_name, index, element, skip_text_boxes=skip_boxes, in_txbx=in_txbx
            )
        )
    return atoms


def _include_atom(atom: _Atom, view: str) -> bool:
    if atom.is_synthetic:
        return not (view == "original" and atom.in_insert)
    if view == "current":
        return atom.tag == _T
    if view == "original":
        # nothing inside a pending insertion existed in the original —
        # including deletions nested within it
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
    """Outcome of a `Span.replace` call."""

    story: str
    deleted_text: str
    inserted_text: str
    tracked: bool
    revision_ids: Tuple[int, ...]

    def to_dict(self) -> dict:
        return {
            "schema": "paper_replace",
            "version": 1,
            "story": self.story,
            "deleted_text": self.deleted_text,
            "inserted_text": self.inserted_text,
            "tracked": self.tracked,
            "revision_ids": list(self.revision_ids),
        }


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
    crosses_paragraphs: bool
    _document: "Document" = field(repr=False)
    _atoms: "List[_Atom]" = field(repr=False)
    _start_offset: int = field(repr=False)  # into first atom's text
    _end_offset: int = field(repr=False)  # exclusive, into last atom's text
    _norm_start: int = field(repr=False)  # position in the story's normalized text
    _consumed: bool = field(default=False, repr=False)  # set by tracked replace

    # -- validation -------------------------------------------------------

    def _current_slice(self) -> str:
        """The span's text as the tree holds it right now (paragraph
        boundaries render as the same "\\n" the capture used)."""
        if len(self._atoms) == 1:
            return self._atoms[0].text[self._start_offset : self._end_offset]
        pieces = [self._atoms[0].text[self._start_offset :]]
        previous = self._atoms[0]
        for atom in self._atoms[1:-1]:
            if atom.paragraph is not previous.paragraph:
                pieces.append("\n")
            pieces.append(atom.text)
            previous = atom
        if self._atoms[-1].paragraph is not previous.paragraph:
            pieces.append("\n")
        pieces.append(self._atoms[-1].text[: self._end_offset])
        return "".join(pieces)

    def _validate_fresh(self) -> None:
        if self._consumed:
            raise TargetNotFoundError(
                "span was consumed by a tracked replace; re-find the text"
            )
        if self._current_slice() != self.text:
            raise TargetNotFoundError(
                f"span is stale: expected {self.text!r} at its anchor"
                f" ({self.anchor.to_dict()}) but the document has changed"
            )
        # detached atoms still hold their text; an edit landing in an orphaned
        # subtree would report success and reach nothing (§1.4's nightmare)
        story_roots = {id(root) for _, root in _story_elements(self._document)}
        for atom in self._atoms:
            top = atom.element
            while top.getparent() is not None:
                top = top.getparent()
            if id(top) not in story_roots:
                raise TargetNotFoundError(
                    "span is stale: its containing structure was removed from"
                    " the document"
                )

    def _validate_replaceable(self) -> None:
        for atom in self._atoms:
            if atom.is_synthetic:
                raise UnsupportedStructureError(
                    "span crosses a tab or line break; replace the text"
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

    # -- replace ----------------------------------------------------------

    def replace(
        self,
        new_text: str,
        *,
        tracked: bool = False,
        author: Optional[str] = None,
        date: Optional[dt.datetime] = None,
    ) -> ReplaceResult:
        """Replace this span's text, preserving run formatting.

        Untracked: surgical in-place edit — untouched runs keep their `rPr`
        byte-identical; the replacement renders with the start run's
        formatting. Tracked: emits a minimal `w:del`/`w:ins` pair (common
        prefix/suffix trimmed) stamped with `author`, `date` (default: the
        injectable clock) and a unique revision id.

        All refusal conditions are checked before any mutation (§1.3).
        """
        if tracked and not author:
            raise ValueError("author is required when tracked=True")
        self._validate_fresh()
        self._validate_replaceable()
        if tracked:
            return self._tracked_replace(new_text, author=author, date=date)  # type: ignore[arg-type]
        return self._plain_replace(new_text)

    def _plain_replace(self, new_text: str) -> ReplaceResult:
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
        )
        self.text = new_text
        self._end_offset = self._start_offset + len(new_text)
        del self._atoms[1:]
        return result

    def _tracked_replace(
        self, new_text: str, *, author: str, date: Optional[dt.datetime]
    ) -> ReplaceResult:
        import copy

        from docx.oxml.revision import CT_RunTrackChange

        for atom in self._atoms:
            if atom.in_hyperlink:
                raise UnsupportedStructureError(
                    "tracked replacement inside a hyperlink is not supported in v0"
                )
            if atom.run is None:
                raise UnsupportedStructureError(
                    "matched text is not inside a run; cannot anchor a revision"
                )
        # layered revisions: a span straddling a pending w:ins would emit a
        # w:del claiming inserted text was base-document content — fabricated
        # history that corrupts reject/original views. Fully inside ONE w:ins
        # nests correctly and is allowed.
        enclosing_insertions = {_enclosing_insertion(atom.element) for atom in self._atoms}
        if len(enclosing_insertions) > 1:
            raise UnsupportedStructureError(
                "span overlaps a pending tracked insertion; accept or reject"
                " the existing revision first, or target text inside or"
                " outside it, not across its boundary"
            )
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
        if old_mid:
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
    """The nearest `w:ins` ancestor of `element`, or None."""
    node = element.getparent()
    while node is not None:
        if node.tag == _INS:
            return node
        node = node.getparent()
    return None


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
    highest = 0
    for _, root in _story_elements(document):
        for node in root.iter(_INS, _DEL):
            value = node.get(w_id)
            try:
                highest = max(highest, int(value)) if value else highest
            except ValueError:
                continue
    return highest + 1


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
) -> "List[Span]":
    raw_text, char_map = _assemble(atoms)
    normalized, norm_map = _normalized_with_map(raw_text)
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
        spans.append(
            Span(
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
                crosses_paragraphs=crosses,
                _document=document,
                _atoms=list(span_atoms),
                _start_offset=start_offset,
                _end_offset=end_offset + 1,
                _norm_start=found_at,
            )
        )
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
        atoms = [
            atom
            for atom in _story_atoms(document, story_name, root)
            if _include_atom(atom, view)
        ]
        if not atoms:
            continue
        spans = _spans_for_story(document, story_name, atoms, needle_norm)
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

    Zero matches raise |TargetNotFoundError|; more than one (after `nth`,
    `near`, `story` disambiguation) raises |AmbiguousTargetError| — the
    library never guesses which occurrence you meant.
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
