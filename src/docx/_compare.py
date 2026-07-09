"""Compare: generate a native tracked-changes redline.

`compare(original, revised, *, author, ...)` returns a new document — the
original with `w:ins`/`w:del` markup that transforms it into the revised
text. The algebra is the contract:

* `accept_all(compare(A, B))` == B in visible text, across every story;
* `reject_all(compare(A, B))` == A;
* `compare(A, A)` emits zero revisions;
* identical inputs produce byte-identical output (with a fixed `date`).

Declared limits (typed refusals or report-only findings, never silence):
insertions/deletions only — no move or rPrChange synthesis; no cross-story
detection; paragraph add/remove outside the main body refuses; changed
merged-cell tables refuse; images/objects and formatting-only differences
are REPORTED in `CompareResult.findings`, not redlined; a block budget of
`_MAX_BLOCKS` per story refuses above it.

Public path: `docx.package.compare` (kernel re-export).
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import TYPE_CHECKING, List, Optional, Tuple

from docx import _clock
from docx.errors import PaperRefusal, UnsupportedStructureError
from docx.oxml.ns import qn
from docx.story import (
    Anchor,
    _build_block,
    _iter_block_elements,
    _story_elements,
    content_hash,
)

if TYPE_CHECKING:
    import datetime as dt

    from lxml.etree import _Element

    from docx.document import Document

_P = qn("w:p")
_TBL = qn("w:tbl")
_TR = qn("w:tr")
_TC = qn("w:tc")
_TRPR = qn("w:trPr")
_TCPR = qn("w:tcPr")
_PPR = qn("w:pPr")
_RPR = qn("w:rPr")
_R = qn("w:r")
_T = qn("w:t")
_DEL_TEXT = qn("w:delText")
_INSTR_TEXT = qn("w:instrText")
_DEL_INSTR_TEXT = qn("w:delInstrText")
_HYPERLINK = qn("w:hyperlink")
_BODY = qn("w:body")

#: paragraph children a tracked deletion leaves in place (mined from
#: docx.blocks._TRANSPARENT_PARAGRAPH_CHILDREN's rationale)
_TRANSPARENT_TAGS = tuple(
    qn(tag)
    for tag in (
        "w:bookmarkStart",
        "w:bookmarkEnd",
        "w:commentRangeStart",
        "w:commentRangeEnd",
    )
)
_PROOF_ERR = qn("w:proofErr")

#: perf budget — documented typed refusal above this many blocks per story
_MAX_BLOCKS = 10_000

#: similarity threshold below which paired blocks redline as del+ins rather
#: than a word-level edit
_PAIR_RATIO = 0.5


@dataclass(frozen=True)
class CompareFinding:
    """One difference compare REPORTS instead of redlining (declared limit)."""

    kind: str
    story: str
    detail: str

    def to_dict(self) -> dict:
        return {"kind": self.kind, "story": self.story, "detail": self.detail}


@dataclass
class CompareResult:
    """The redlined document plus everything compare could not redline."""

    document: "Document"
    findings: List[CompareFinding]
    revision_count: int
    stories: Tuple[str, ...]

    def to_dict(self) -> dict:
        return {
            "schema": "paper_compare",
            "version": 1,
            "revision_count": self.revision_count,
            "stories": list(self.stories),
            "findings": [finding.to_dict() for finding in self.findings],
        }


@dataclass
class _Ctx:
    author: str
    stamp: "dt.datetime"
    granularity: str
    story: str
    document: "Document"
    findings: List[CompareFinding]
    numbering_ids: frozenset
    #: one shared cell per compare run: ids on DETACHED clones are invisible
    #: to the attached-tree scan, so the counter is the source of truth
    id_cell: List[int]


def compare(
    original,
    revised,
    *,
    author: str,
    date: "Optional[dt.datetime]" = None,
    granularity: str = "word",
    materialize: Optional[str] = None,
) -> CompareResult:
    """A new document: `original` + tracked changes that produce `revised`.

    `original`/`revised` are .docx paths. Inputs carrying pending revisions
    refuse unless `materialize` ("accept" | "reject") routes them through
    resolution on in-memory working copies first (the files on disk are
    never touched). `granularity`: "word" (token-level edits inside matched
    paragraphs) or "block" (whole-paragraph del+ins only).
    """
    import docx as _docx

    if not author:
        raise ValueError("author is required")
    if granularity not in ("word", "block"):
        raise ValueError(
            f"granularity must be 'word' or 'block', got {granularity!r}"
        )
    if materialize not in (None, "accept", "reject"):
        raise ValueError(
            f"materialize must be None, 'accept' or 'reject', got {materialize!r}"
        )
    stamp = date if date is not None else _clock.now()
    document = _docx.Document(str(original))
    revised_doc = _docx.Document(str(revised))
    findings: List[CompareFinding] = []
    for which, doc in (("original", document), ("revised", revised_doc)):
        if len(doc.revisions):
            if materialize is None:
                raise UnsupportedStructureError(
                    f"the {which} document carries pending tracked revisions;"
                    " compare would conflate them with its own redline. Pass"
                    " materialize='accept' or 'reject' to resolve working"
                    " copies first (the input files are not modified)"
                )
            from docx.scrubbing import finalize

            finalize(doc, revisions=materialize)
    from docx.protection import acknowledge_protection, protection_status

    status = protection_status(document)
    if status.enforced:
        acknowledge_protection(document)
        findings.append(
            CompareFinding(
                kind="document_protection_present",
                story="word/settings.xml",
                detail=(
                    f"original enforces {status.edit!r} protection; the"
                    " redline was produced anyway (a new artifact) and the"
                    " setting is carried through unmodified"
                ),
            )
        )

    stories_o = {s: root for s, root in _story_elements(document)}
    stories_r = {s: root for s, root in _story_elements(revised_doc)}
    comments = "word/comments.xml"
    if _story_text_of(comments, stories_o.get(comments)) != _story_text_of(
        comments, stories_r.get(comments)
    ):
        findings.append(
            CompareFinding(
                kind="comments_differ",
                story=comments,
                detail=(
                    "comment content differs between the documents; compare"
                    " never redlines the comments story (report-only)"
                ),
            )
        )
    names_o = set(stories_o) - {comments}
    names_r = set(stories_r) - {comments}
    if names_o != names_r:
        raise UnsupportedStructureError(
            "the documents carry different story parts"
            f" (only-original: {sorted(names_o - names_r)},"
            f" only-revised: {sorted(names_r - names_o)}); compare cannot"
            " redline story-part addition or removal"
        )
    numbering_ids = _numbering_ids(document)
    compared: List[str] = []
    id_cell = [0]
    for story in sorted(names_o):
        ctx = _Ctx(
            author=author,
            stamp=stamp,
            granularity=granularity,
            story=story,
            document=document,
            findings=findings,
            numbering_ids=numbering_ids,
            id_cell=id_cell,
        )
        _compare_story(ctx, stories_o[story], stories_r[story])
        compared.append(story)
    return CompareResult(
        document=document,
        findings=findings,
        revision_count=len(document.revisions),
        stories=tuple(compared),
    )


def _story_text_of(story: str, root) -> Optional[str]:
    if root is None:
        return None
    return "\n".join(text for _kind, _el, text in _story_blocks(story, root))


def _numbering_ids(document: "Document") -> frozenset:
    try:
        from docx.numbering import list_numbering

        return frozenset(
            definition.num_id
            for definition in list_numbering(document).definitions
        )
    except Exception:
        return frozenset()


def _story_blocks(story: str, root: "_Element") -> "List[Tuple[str, _Element, str]]":
    blocks = []
    for kind, index, element, in_sdt, in_txbx in _iter_block_elements(story, root):
        block = _build_block(
            story, kind, index, element, "current", in_sdt=in_sdt, in_txbx=in_txbx
        )
        blocks.append((kind, element, block.text))
    return blocks


def _compare_story(ctx: _Ctx, root_o: "_Element", root_r: "_Element") -> None:
    blocks_o = _story_blocks(ctx.story, root_o)
    blocks_r = _story_blocks(ctx.story, root_r)
    if max(len(blocks_o), len(blocks_r)) > _MAX_BLOCKS:
        raise UnsupportedStructureError(
            f"story {ctx.story} exceeds the compare block budget"
            f" ({_MAX_BLOCKS} blocks); split the documents"
        )
    matcher = SequenceMatcher(
        None,
        [f"{k}\x00{t}" for k, _e, t in blocks_o],
        [f"{k}\x00{t}" for k, _e, t in blocks_r],
        autojunk=False,
    )
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for offset in range(i2 - i1):
                _report_formatting_difference(
                    ctx, blocks_o[i1 + offset], blocks_r[j1 + offset]
                )
            continue
        if tag == "delete":
            for kind, element, text in blocks_o[i1:i2]:
                _delete_block(ctx, kind, element, text)
            continue
        if tag == "insert":
            _insert_blocks(ctx, blocks_o, i1, blocks_r[j1:j2])
            continue
        _compare_region(ctx, blocks_o[i1:i2], blocks_r[j1:j2])


def _pair_region(old_run, new_run) -> "List[Tuple[str, Optional[int], Optional[int]]]":
    """Order-preserving best-similarity block pairing within a changed region
    (LCS-style DP maximizing summed pair ratios above `_PAIR_RATIO`)."""
    m, n = len(old_run), len(new_run)
    ratios: dict = {}

    def ratio(i: int, j: int) -> float:
        if (i, j) not in ratios:
            kind_o, _eo, text_o = old_run[i]
            kind_r, _er, text_r = new_run[j]
            if kind_o != kind_r:
                ratios[(i, j)] = 0.0
            else:
                score = SequenceMatcher(None, text_o, text_r, autojunk=False).ratio()
                ratios[(i, j)] = score if score >= _PAIR_RATIO else 0.0
        return ratios[(i, j)]

    score = [[0.0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            best = max(score[i - 1][j], score[i][j - 1])
            pair = ratio(i - 1, j - 1)
            if pair > 0.0:
                best = max(best, score[i - 1][j - 1] + pair)
            score[i][j] = best
    ops: "List[Tuple[str, Optional[int], Optional[int]]]" = []
    i, j = m, n
    while i > 0 or j > 0:
        pair = ratio(i - 1, j - 1) if (i > 0 and j > 0) else 0.0
        if pair > 0.0 and score[i][j] == score[i - 1][j - 1] + pair:
            ops.append(("pair", i - 1, j - 1))
            i -= 1
            j -= 1
        elif i > 0 and (j == 0 or score[i][j] == score[i - 1][j]):
            ops.append(("delete", i - 1, None))
            i -= 1
        else:
            ops.append(("insert", None, j - 1))
            j -= 1
    ops.reverse()
    return ops


def _compare_region(ctx: _Ctx, old_run, new_run) -> None:
    """Emit tracked changes for one changed region; a cursor threads the
    output position so unpaired insertions land in document order."""
    cursor: "Optional[_Element]" = None
    for op, index_o, index_r in _pair_region(old_run, new_run):
        if op == "pair":
            kind_o, element_o, text_o = old_run[index_o]
            _kind_r, element_r, text_r = new_run[index_r]
            if kind_o == "table":
                _compare_table(ctx, element_o, element_r)
                cursor = element_o
            elif ctx.granularity == "word":
                _replace_paragraph_text(ctx, element_o, text_o, text_r)
                cursor = element_o
            else:
                _delete_block(ctx, kind_o, element_o, text_o)
                clone = _cloned_as_insertion(ctx, "paragraph", element_r)
                _require_container_anchor(ctx, element_o)
                element_o.addnext(clone)
                cursor = clone
        elif op == "delete":
            kind_o, element_o, text_o = old_run[index_o]
            _delete_block(ctx, kind_o, element_o, text_o)
            cursor = element_o
        else:  # insert
            kind_r, element_r, _text_r = new_run[index_r]
            clone = _cloned_as_insertion(ctx, kind_r, element_r)
            if cursor is None:
                reference = old_run[0][1]
                _require_container_anchor(ctx, reference)
                reference.addprevious(clone)
            else:
                _require_container_anchor(ctx, cursor)
                cursor.addnext(clone)
            cursor = clone


def _report_formatting_difference(ctx: _Ctx, block_o, block_r) -> None:
    from lxml import etree

    _kind, element_o, text = block_o
    _kind_r, element_r, _text_r = block_r
    if etree.tostring(element_o) != etree.tostring(element_r):
        ctx.findings.append(
            CompareFinding(
                kind="formatting_only_difference",
                story=ctx.story,
                detail=(
                    "identical visible text with differing markup"
                    f" (formatting or metadata): {text[:80]!r}"
                ),
            )
        )


# ---------------------------------------------------------------------------
# whole-block emission
# ---------------------------------------------------------------------------


def _next_id(ctx: _Ctx) -> int:
    from docx.search import _next_revision_id

    value = max(ctx.id_cell[0], _next_revision_id(ctx.document))
    ctx.id_cell[0] = value + 1
    return value


def _delete_block(ctx: _Ctx, kind: str, element: "_Element", text: str) -> None:
    if kind == "table":
        for row in element.findall(_TR):
            _mark_row_deleted(ctx, row)
        return
    _mark_paragraph_deleted(ctx, element)


def _mark_paragraph_deleted(ctx: _Ctx, paragraph: "_Element") -> None:
    """Wrap the paragraph's content in `w:del` (per-run formatting kept) and
    stamp the mark deleted — the shape Word emits for a deleted paragraph."""
    from docx.blocks import _stamp_paragraph_mark
    from docx.oxml.revision import CT_RunTrackChange

    _refuse_unrepresentable_children(ctx, paragraph, verb="delete")
    revision_id = _next_id(ctx)
    current_del = None
    for child in list(paragraph):
        if child.tag == _PPR or child.tag in _TRANSPARENT_TAGS:
            current_del = None
            continue
        if child.tag == _PROOF_ERR:
            paragraph.remove(child)
            continue
        if child.tag == _HYPERLINK:
            _wrap_children_deleted(ctx, child, revision_id)
            current_del = None
            continue
        if current_del is None:
            current_del = CT_RunTrackChange.new(
                "w:del", revision_id, ctx.author, ctx.stamp
            )
            child.addprevious(current_del)
        _retag_deleted_text(child)
        current_del.append(child)
    _stamp_paragraph_mark(
        paragraph, "w:del", _next_id(ctx), ctx.author, ctx.stamp
    )


def _wrap_children_deleted(ctx: _Ctx, container: "_Element", revision_id: int) -> None:
    from docx.oxml.revision import CT_RunTrackChange

    wrapper = None
    for child in list(container):
        if child.tag != _R:
            wrapper = None
            continue
        if wrapper is None:
            wrapper = CT_RunTrackChange.new(
                "w:del", revision_id, ctx.author, ctx.stamp
            )
            child.addprevious(wrapper)
        _retag_deleted_text(child)
        wrapper.append(child)


def _retag_deleted_text(node: "_Element") -> None:
    for text_elm in node.iter(_T):
        text_elm.tag = _DEL_TEXT
    for text_elm in node.iter(_INSTR_TEXT):
        text_elm.tag = _DEL_INSTR_TEXT


_UNREPRESENTABLE_IN_DELETE = (qn("w:sdt"), qn("w:fldSimple"))
_FLD_CHAR = qn("w:fldChar")
_FLD_CHAR_TYPE = qn("w:fldCharType")


def _fldchar_unbalanced(element: "_Element") -> bool:
    """True when the element holds part of a complex field but not all of it
    — removing or inserting it whole would strand a begin/end marker."""
    depth = 0
    for fld_char in element.iter(_FLD_CHAR):
        fld_type = fld_char.get(_FLD_CHAR_TYPE)
        if fld_type == "begin":
            depth += 1
        elif fld_type == "end":
            if depth == 0:
                return True  # an end whose begin lives outside this element
            depth -= 1
    return depth != 0


def _refuse_unrepresentable_children(
    ctx: _Ctx, paragraph: "_Element", *, verb: str
) -> None:
    for child in paragraph:
        if child.tag in _UNREPRESENTABLE_IN_DELETE:
            local = child.tag.rsplit("}", 1)[-1]
            raise UnsupportedStructureError(
                f"compare cannot {verb} a paragraph containing <w:{local}>"
                f" in {ctx.story}; restructure the change or resolve it"
                " manually (a declared limit)"
            )
    if _fldchar_unbalanced(paragraph):
        raise UnsupportedStructureError(
            f"compare cannot {verb} one paragraph of a multi-paragraph"
            f" complex field in {ctx.story} — the begin/end markers would"
            " unbalance (a declared limit)"
        )
    from docx.blocks import _named_bookmarks_in

    named = _named_bookmarks_in(paragraph)
    if named:
        raise UnsupportedStructureError(
            f"compare cannot {verb} a paragraph carrying named bookmark(s)"
            f" {sorted(named)} in {ctx.story} — cross-references would"
            " dangle (a declared limit)"
        )


def _insert_blocks(
    ctx: _Ctx,
    blocks_o,
    insert_index: int,
    new_blocks,
    after_element: "Optional[_Element]" = None,
) -> None:
    """Insert sanitized clones of `new_blocks` as tracked insertions, before
    the original block at `insert_index` (or after `after_element`)."""
    clones = [
        _cloned_as_insertion(ctx, kind, element) for kind, element, _t in new_blocks
    ]
    if after_element is not None:
        reference = after_element
        for clone in clones:
            reference.addnext(clone)
            reference = clone
        return
    if insert_index < len(blocks_o):
        reference = blocks_o[insert_index][1]
        _require_container_anchor(ctx, reference)
        for clone in clones:
            reference.addprevious(clone)
        return
    if not blocks_o:
        raise UnsupportedStructureError(
            f"compare cannot insert into the empty story {ctx.story}"
        )
    reference = blocks_o[-1][1]
    _require_container_anchor(ctx, reference)
    for clone in clones:
        reference.addnext(clone)
        reference = clone


def _require_container_anchor(ctx: _Ctx, reference: "_Element") -> None:
    """Paragraph add/remove is only safe where the anchor block sits directly
    in the story container (body/header/footer root) — a text-box or
    footnote-entry neighbor would put the new block in the wrong container."""
    parent = reference.getparent()
    if parent is None or parent.tag == qn("w:txbxContent"):
        raise UnsupportedStructureError(
            f"compare cannot add or remove whole blocks next to text-box"
            f" content in {ctx.story} (a declared limit)"
        )
    if ctx.story.startswith("word/footnotes") or ctx.story.startswith(
        "word/endnotes"
    ):
        raise UnsupportedStructureError(
            f"compare cannot add or remove whole blocks in {ctx.story};"
            " note-content edits must stay within existing notes"
            " (a declared limit)"
        )


def _cloned_as_insertion(ctx: _Ctx, kind: str, element: "_Element") -> "_Element":
    from docx.blocks import _stamp_paragraph_mark, _wrap_paragraph_content_as_insertion

    clone = copy.deepcopy(element)
    _sanitize_clone(ctx, clone)
    if kind == "table":
        for row in clone.findall(_TR):
            _mark_cloned_row_inserted(ctx, row)
        return clone
    _wrap_paragraph_content_as_insertion(clone, _next_id(ctx), ctx.author, ctx.stamp)
    _stamp_paragraph_mark(clone, "w:ins", _next_id(ctx), ctx.author, ctx.stamp)
    return clone


def _sanitize_clone(ctx: _Ctx, clone: "_Element") -> None:
    """Strip everything a clone cannot carry across packages (relationship
    ids resolve against the ORIGINAL part) — each removal is a finding."""
    for node in list(clone.iter(qn("w:drawing"), qn("w:object"), qn("w:pict"))):
        holder = node.getparent()
        target = holder if (holder is not None and holder.tag == _R) else node
        if target.getparent() is None:
            continue  # already removed with its sibling
        target.getparent().remove(target)
        ctx.findings.append(
            CompareFinding(
                kind="image_or_object_change",
                story=ctx.story,
                detail=(
                    "an image/object in inserted content was not carried"
                    " into the redline (relationships do not transfer;"
                    " report-only)"
                ),
            )
        )
    for hyperlink in list(clone.iter(_HYPERLINK)):
        parent = hyperlink.getparent()
        for child in list(hyperlink):
            hyperlink.addprevious(child)
        parent.remove(hyperlink)
        ctx.findings.append(
            CompareFinding(
                kind="hyperlink_flattened",
                story=ctx.story,
                detail=(
                    "an inserted hyperlink was flattened to its text (its"
                    " relationship id does not transfer)"
                ),
            )
        )
    for sdt in list(clone.iter(qn("w:sdt"))):
        content = sdt.find(qn("w:sdtContent"))
        parent = sdt.getparent()
        if content is not None:
            for child in list(content):
                sdt.addprevious(child)
        parent.remove(sdt)
        ctx.findings.append(
            CompareFinding(
                kind="content_control_flattened",
                story=ctx.story,
                detail="an inserted content control was flattened to its content",
            )
        )
    for sect_pr in list(clone.iter(qn("w:sectPr"))):
        sect_pr.getparent().remove(sect_pr)
        ctx.findings.append(
            CompareFinding(
                kind="section_break_not_carried",
                story=ctx.story,
                detail=(
                    "a section break in inserted content was dropped (its"
                    " header/footer relationships do not transfer)"
                ),
            )
        )
    for num_pr in list(clone.iter(qn("w:numPr"))):
        num_id = num_pr.find(qn("w:numId"))
        value = num_id.get(qn("w:val")) if num_id is not None else None
        if value is None or int(value) not in ctx.numbering_ids:
            num_pr.getparent().remove(num_pr)
            ctx.findings.append(
                CompareFinding(
                    kind="numbering_reference_stripped",
                    story=ctx.story,
                    detail=(
                        f"inserted content referenced numbering id {value}"
                        " which the original does not define; the reference"
                        " was stripped"
                    ),
                )
            )
    for proof_err in list(clone.iter(_PROOF_ERR)):
        proof_err.getparent().remove(proof_err)
    stripped_anchor = False
    for tag in ("w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart",
                "w:commentRangeEnd", "w:commentReference"):
        for node in list(clone.iter(qn(tag))):
            node.getparent().remove(node)
            stripped_anchor = True
    if stripped_anchor:
        ctx.findings.append(
            CompareFinding(
                kind="anchor_not_carried",
                story=ctx.story,
                detail=(
                    "bookmark or comment anchors in inserted content were"
                    " dropped (their ids/targets do not transfer;"
                    " report-only)"
                ),
            )
        )
    if _fldchar_unbalanced(clone):
        raise UnsupportedStructureError(
            f"compare cannot insert one paragraph of a multi-paragraph"
            f" complex field in {ctx.story} — the begin/end markers would"
            " unbalance (a declared limit)"
        )


# ---------------------------------------------------------------------------
# tables
# ---------------------------------------------------------------------------


_MERGE_TAGS = (qn("w:gridSpan"), qn("w:vMerge"))


def _refuse_merged_rows(ctx: _Ctx, rows) -> None:
    for row in rows:
        for tag in _MERGE_TAGS:
            if row.find(f".//{tag}") is not None:
                raise UnsupportedStructureError(
                    f"compare cannot redline changed table rows containing"
                    f" merged cells in {ctx.story} (a declared limit)"
                )


def _row_text(row: "_Element") -> str:
    return "\x00".join(
        "".join(t.text or "" for t in cell.iter(_T)) for cell in row.findall(_TC)
    )


def _compare_table(ctx: _Ctx, table_o: "_Element", table_r: "_Element") -> None:
    rows_o = table_o.findall(_TR)
    rows_r = table_r.findall(_TR)
    matcher = SequenceMatcher(
        None, [_row_text(r) for r in rows_o], [_row_text(r) for r in rows_r],
        autojunk=False,
    )
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "delete":
            _refuse_merged_rows(ctx, rows_o[i1:i2])
            for row in rows_o[i1:i2]:
                _mark_row_deleted(ctx, row)
            continue
        if tag == "insert":
            _refuse_merged_rows(ctx, rows_r[j1:j2])
            _insert_rows(ctx, rows_o, i1, rows_r[j1:j2])
            continue
        old_rows, new_rows = rows_o[i1:i2], rows_r[j1:j2]
        _refuse_merged_rows(ctx, list(old_rows) + list(new_rows))
        paired = min(len(old_rows), len(new_rows))
        cursor = None  # last row placed in OUTPUT order
        for k in range(paired):
            if _replace_row_cells(ctx, old_rows[k], new_rows[k]):
                cursor = old_rows[k]
            else:
                _mark_row_deleted(ctx, old_rows[k])
                inserted = _insert_rows(
                    ctx, rows_o, i1 + k, [new_rows[k]], after=old_rows[k]
                )
                cursor = inserted[-1]
        for row in old_rows[paired:]:
            _mark_row_deleted(ctx, row)
            cursor = row
        if len(new_rows) > paired:
            _insert_rows(
                ctx, rows_o, i2, new_rows[paired:],
                after=cursor if cursor is not None else old_rows[-1],
            )


def _visible_paragraph_text(paragraph: "_Element") -> str:
    """Paragraph text with tabs/breaks as their synthetic characters —
    matching the atom stream _paragraph_span slices, so word-diff offsets
    never desynchronize."""
    pieces: List[str] = []
    for child in paragraph:
        if child.tag == _PPR:
            continue  # w:tabs STOP definitions live here, not tab chars
        for node in child.iter():
            if node.tag == _T:
                pieces.append(node.text or "")
            elif node.tag == qn("w:tab"):
                pieces.append("\t")
            elif node.tag in (qn("w:br"), qn("w:cr")):
                pieces.append("\n")
    return "".join(pieces)


def _replace_row_cells(ctx: _Ctx, row_o: "_Element", row_r: "_Element") -> bool:
    """Cell-wise word-level edits when the rows are shape-compatible;
    False -> caller falls back to row del+ins."""
    cells_o, cells_r = row_o.findall(_TC), row_r.findall(_TC)
    if len(cells_o) != len(cells_r) or ctx.granularity != "word":
        return False
    if row_o.find(f".//{_TBL}") is not None or row_r.find(f".//{_TBL}") is not None:
        return False  # nested-table differences need whole-row del+ins
    plan = []
    for cell_o, cell_r in zip(cells_o, cells_r):
        paragraphs_o = cell_o.findall(_P)
        paragraphs_r = cell_r.findall(_P)
        if len(paragraphs_o) != 1 or len(paragraphs_r) != 1:
            texts_o = [_visible_paragraph_text(p) for p in paragraphs_o]
            texts_r = [_visible_paragraph_text(p) for p in paragraphs_r]
            if texts_o == texts_r:
                continue
            return False
        text_o = _visible_paragraph_text(paragraphs_o[0])
        text_r = _visible_paragraph_text(paragraphs_r[0])
        if text_o != text_r:
            plan.append((paragraphs_o[0], text_o, text_r))
    for paragraph, text_o, text_r in plan:
        _replace_paragraph_text(ctx, paragraph, text_o, text_r)
    return True


def _mark_row_deleted(ctx: _Ctx, row: "_Element") -> None:
    """`w:trPr/w:del` + cell content in `w:del`/`w:delText` + del-stamped
    cell paragraph marks — the row-deletion shape, emitted."""
    from docx.oxml.parser import OxmlElement

    revision_id = _next_id(ctx)
    tr_pr = row.find(_TRPR)
    if tr_pr is None:
        tr_pr = OxmlElement("w:trPr")
        row.insert(0, tr_pr)
    marker = OxmlElement("w:del")
    marker.set(qn("w:id"), str(revision_id))
    marker.set(qn("w:author"), ctx.author)
    marker.set(qn("w:date"), _iso(ctx.stamp))
    tr_pr.append(marker)
    for cell in row.findall(_TC):
        for paragraph in cell.findall(_P):
            _mark_paragraph_deleted(ctx, paragraph)


def _mark_cloned_row_inserted(ctx: _Ctx, row: "_Element") -> None:
    from docx.blocks import _stamp_paragraph_mark, _wrap_paragraph_content_as_insertion
    from docx.oxml.parser import OxmlElement

    revision_id = _next_id(ctx)
    tr_pr = row.find(_TRPR)
    if tr_pr is None:
        tr_pr = OxmlElement("w:trPr")
        row.insert(0, tr_pr)
    marker = OxmlElement("w:ins")
    marker.set(qn("w:id"), str(revision_id))
    marker.set(qn("w:author"), ctx.author)
    marker.set(qn("w:date"), _iso(ctx.stamp))
    tr_pr.append(marker)
    for cell in row.findall(_TC):
        for paragraph in cell.findall(_P):
            _wrap_paragraph_content_as_insertion(
                paragraph, _next_id(ctx), ctx.author, ctx.stamp
            )
            _stamp_paragraph_mark(
                paragraph, "w:ins", _next_id(ctx), ctx.author, ctx.stamp
            )


def _insert_rows(ctx: _Ctx, rows_o, index: int, new_rows, after=None) -> list:
    clones = []
    for row in new_rows:
        clone = copy.deepcopy(row)
        _sanitize_clone(ctx, clone)
        _mark_cloned_row_inserted(ctx, clone)
        clones.append(clone)
    if after is not None:
        reference = after
        for clone in clones:
            reference.addnext(clone)
            reference = clone
        return clones
    if index < len(rows_o):
        for clone in clones:
            rows_o[index].addprevious(clone)
        return clones
    reference = rows_o[-1]
    for clone in clones:
        reference.addnext(clone)
        reference = clone
    return clones


def _iso(stamp: "dt.datetime") -> str:
    from docx.oxml.simpletypes import ST_DateTime

    # the same serialization the oxml layer uses for w:ins/w:del dates —
    # one logical edit must never carry two different timestamps
    return ST_DateTime.convert_to_xml(stamp)


# ---------------------------------------------------------------------------
# word-level paragraph edits (through the Span machinery)
# ---------------------------------------------------------------------------


def _replace_paragraph_text(
    ctx: _Ctx, paragraph: "_Element", text_o: str, text_r: str
) -> None:
    """Token-level tracked edits inside one matched paragraph; whole-block
    del+ins fallback when the span machinery refuses a region.

    A refusal can arrive AFTER earlier regions already emitted revisions;
    falling back on the half-redlined paragraph would nest fresh w:del
    inside w:del — so the pristine paragraph is swapped back in before the
    fallback runs."""
    regions = _token_regions(text_o, text_r)
    pristine = copy.deepcopy(paragraph)
    try:
        for start, end, replacement in reversed(regions):
            span = _paragraph_span(ctx, paragraph, start, end)
            if span is None:
                raise UnsupportedStructureError("region unmappable to atoms")
            span.replace(
                replacement, tracked=True, author=ctx.author, date=ctx.stamp
            )
    except PaperRefusal:
        parent = paragraph.getparent()
        if parent is not None:
            parent.replace(paragraph, pristine)
            paragraph = pristine
        _fallback_paragraph_replace(ctx, paragraph, text_r)


def _fallback_paragraph_replace(
    ctx: _Ctx, paragraph: "_Element", text_r: str
) -> None:
    from docx.blocks import _stamp_paragraph_mark, _wrap_paragraph_content_as_insertion
    from docx.oxml.parser import OxmlElement

    if (
        paragraph.find(f".//{_FLD_CHAR}") is not None
        or paragraph.find(f".//{qn('w:fldSimple')}") is not None
    ):
        ctx.findings.append(
            CompareFinding(
                kind="field_flattened",
                story=ctx.story,
                detail=(
                    "a changed paragraph containing a field was redlined as"
                    " whole-paragraph delete+insert; the replacement carries"
                    " the field's TEXT, not the field itself (report-only)"
                ),
            )
        )

    replacement = OxmlElement("w:p")
    p_pr = paragraph.find(_PPR)
    if p_pr is not None:
        cloned_ppr = copy.deepcopy(p_pr)
        for r_pr in list(cloned_ppr.findall(_RPR)):
            cloned_ppr.remove(r_pr)  # mark stamps belong to the new mark only
        replacement.append(cloned_ppr)
    run = OxmlElement("w:r")
    run.add_t(text_r)
    replacement.append(run)
    _wrap_paragraph_content_as_insertion(
        replacement, _next_id(ctx), ctx.author, ctx.stamp
    )
    _stamp_paragraph_mark(replacement, "w:ins", _next_id(ctx), ctx.author, ctx.stamp)
    _mark_paragraph_deleted(ctx, paragraph)
    paragraph.addnext(replacement)


def _token_regions(old: str, new: str) -> "List[Tuple[int, int, str]]":
    """Changed character regions (old-start, old-end, replacement) from a
    token-level diff; zero-width regions are widened by one anchor char so
    every region maps to at least one atom."""
    tokens_o = re.findall(r"\S+|\s+", old)
    tokens_r = re.findall(r"\S+|\s+", new)
    offsets = [0]
    for token in tokens_o:
        offsets.append(offsets[-1] + len(token))
    regions: "List[Tuple[int, int, str]]" = []
    matcher = SequenceMatcher(None, tokens_o, tokens_r, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        start, end = offsets[i1], offsets[i2]
        replacement = "".join(tokens_r[j1:j2])
        if start == end:  # pure insertion: widen over one anchor char
            if start > 0:
                start -= 1
                replacement = old[start] + replacement
            elif end < len(old):
                replacement = replacement + old[end]
                end += 1
            else:  # empty original paragraph text
                regions.append((0, 0, replacement))
                continue
        regions.append((start, end, replacement))
    return regions


def _paragraph_span(ctx: _Ctx, paragraph: "_Element", start: int, end: int):
    """A Span over [start, end) of the paragraph's current-view text, built
    with the exact conventions of `docx.search._spans_for_story`."""
    from docx.search import Span, _include_atom, _story_atoms

    root = next(
        (root for s, root in _story_elements(ctx.document) if s == ctx.story), None
    )
    if root is None:
        return None
    atoms = [
        atom
        for atom in _story_atoms(ctx.document, ctx.story, root)
        if atom.paragraph is paragraph and _include_atom(atom, "current")
    ]
    if not atoms or start >= end:
        return None
    positions = []  # (atom_index, offset) per character
    for index, atom in enumerate(atoms):
        for offset in range(len(atom.text)):
            positions.append((index, offset))
    if end > len(positions):
        return None
    start_atom, start_offset = positions[start]
    end_atom, end_offset = positions[end - 1]
    span_atoms = atoms[start_atom : end_atom + 1]
    text = "".join(a.text for a in atoms)[start:end]
    return Span(
        text=text,
        story=ctx.story,
        anchor=Anchor(
            story=ctx.story,
            index=span_atoms[0].block_index,
            content_hash=content_hash(text),
        ),
        in_insert=any(a.in_insert for a in span_atoms),
        in_delete=any(a.in_delete or a.tag == _DEL_TEXT for a in span_atoms),
        in_content_control=any(a.sdt is not None for a in span_atoms),
        in_text_box=any(a.in_text_box for a in span_atoms),
        in_field=any(a.in_field for a in span_atoms),
        crosses_paragraphs=False,
        _document=ctx.document,
        _atoms=list(span_atoms),
        _start_offset=start_offset,
        _end_offset=end_offset + 1,
        _norm_start=0,
    )
