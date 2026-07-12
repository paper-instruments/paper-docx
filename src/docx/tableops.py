"""Narrow, guarded table operations (paper-docx).

Guards are CELL-WISE: an operation refuses when the cells or rows
it actually touches participate in a merge or hold a nested table — a merged
header row (the default shape of real tables) no longer blocks edits to plain
data cells. The failure mode this module exists to prevent is unchanged: a
"clever" edit that silently reshuffles a complex region. Cell updates route
through `Span.replace`, so formatting preservation, tracked changes and
refusal atomicity all come from the same machinery as body text.
"""

from __future__ import annotations

import copy
import datetime as dt
from typing import TYPE_CHECKING, Optional, Sequence, Tuple

from docx import _clock
from docx.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.oxml.ns import qn
from docx.oxml.revision import CT_RunTrackChange
from docx.protection import _refuse_if_protected
from docx.search import (
    ReplaceResult,
    Span,
    _collect_block_atoms,
    _next_revision_id,
    normalize_text,
)
from docx.story import Anchor, _iter_block_elements, _story_elements, content_hash

if TYPE_CHECKING:
    from docx.document import Document
    from docx.table import Table, _Cell

_T = qn("w:t")


_TC_PR = qn("w:tcPr")
_GRID_SPAN = qn("w:gridSpan")
_V_MERGE = qn("w:vMerge")
_TBL = qn("w:tbl")
_TC = qn("w:tc")
_SDT = qn("w:sdt")
_W_VAL = qn("w:val")
_FLD_CHAR = qn("w:fldChar")


def _refuse_whole_cell_semantics(tc) -> None:
    """Refuse semantics that may begin/end outside the chosen paragraph."""
    from docx.revision import _MARKUP_SCAN_TAGS

    if (
        next(tc.iter(_FLD_CHAR), None) is not None
        or next(tc.iter(qn("w:fldSimple")), None) is not None
    ):
        raise UnsupportedStructureError(
            "cell contains a Word field; update its field/result explicitly"
        )
    revision = next((node for node in tc.iter() if node.tag in _MARKUP_SCAN_TAGS), None)
    if revision is not None:
        raise UnsupportedStructureError(
            "cell contains pending revision markup; resolve it before replacing the cell"
        )


def _cell_is_merged(tc) -> bool:
    tc_pr = tc.find(_TC_PR)
    if tc_pr is None:
        return False
    return tc_pr.find(_GRID_SPAN) is not None or tc_pr.find(_V_MERGE) is not None


def _cell_has_nested_table(tc) -> bool:
    return tc.find(_TBL) is not None


def _refuse_complex_cell(tc, *, row: int, column: int) -> None:
    """Cell-wise guard: only the TARGET cell's complexity refuses.

    A merged header row no longer poisons edits to the plain data cells
    below it — merged-header tables are the default shape of real tables."""
    if _cell_is_merged(tc):
        raise UnsupportedStructureError(
            f"cell ({row}, {column}) participates in a merge (gridSpan/vMerge);"
            " editing it would silently reshuffle the merged region"
        )
    if _cell_has_nested_table(tc):
        raise UnsupportedStructureError(
            f"cell ({row}, {column}) contains a nested table; address the inner"
            " table directly instead"
        )


def _row_cells(tr):
    return tr.findall(_TC)


def _row_has_vmerge(tr) -> bool:
    return any(
        tc.find(_TC_PR) is not None and tc.find(_TC_PR).find(_V_MERGE) is not None
        for tc in _row_cells(tr)
    )


def _row_has_nested_table(tr) -> bool:
    return any(_cell_has_nested_table(tc) for tc in _row_cells(tr))


def _row_continues_merge_from_above(tr) -> bool:
    for tc in _row_cells(tr):
        tc_pr = tc.find(_TC_PR)
        v_merge = tc_pr.find(_V_MERGE) if tc_pr is not None else None
        if v_merge is not None and (v_merge.get(_W_VAL) or "continue") == "continue":
            return True
    return False


def _refuse_tracked_template_row(tr, *, row: int) -> None:
    from docx.revision import _MARKUP_SCAN_TAGS

    revision = next(
        (node for node in tr.iter() if node.tag in _MARKUP_SCAN_TAGS),
        None,
    )
    if revision is None:
        return
    marker = revision.tag.rpartition("}")[2]
    raise UnsupportedStructureError(
        f"template row {row} contains tracked revision metadata ({marker});"
        " resolve its pending revisions before copying formatting from it"
    )


def _refuse_row_op(table: "Table", *, affected_rows, splits_before=None) -> None:
    """Row-op guard: refuse only when the AFFECTED rows intersect a
    vertical merge or hold a nested table, or when insertion would split a
    merge; horizontal merges elsewhere in the table are none of our business."""
    for index in affected_rows:
        tr = table.rows[index]._tr
        if _row_has_vmerge(tr):
            raise UnsupportedStructureError(
                f"row {index} participates in a vertical merge; structural row"
                " operations there would corrupt the merged region"
            )
        if _row_has_nested_table(tr):
            raise UnsupportedStructureError(
                f"row {index} contains a nested table; refusing rather than"
                " duplicating or destroying it wholesale"
            )
    if (
        splits_before is not None
        and splits_before < len(table.rows)
        and _row_continues_merge_from_above(table.rows[splits_before]._tr)
    ):
        raise UnsupportedStructureError(
            f"inserting between rows {splits_before - 1} and {splits_before}"
            " would split a vertical merge"
        )


def _document_of(table: "Table") -> "Document":
    part = table.part
    document = getattr(part, "document", None)
    if document is None:
        raise UnsupportedStructureError(
            "tables outside the main document story are not supported"
        )
    return document


def _cell_text(cell: "_Cell") -> str:
    return "\n".join(p.text for p in cell.paragraphs)


def find_table(document: "Document", *, near_text: str) -> "Table":
    """The table whose cell text contains `near_text` (normalized matching).

    Zero matching tables raise |TargetNotFoundError|; more than one raise
    |AmbiguousTargetError| — make `near_text` more specific.
    """
    needle = normalize_text(near_text)
    matches = []
    for table in document.tables:
        text = normalize_text(
            "\n".join(_cell_text(cell) for row in table.rows for cell in row.cells)
        )
        if needle in text:
            matches.append(table)
    if not matches:
        raise TargetNotFoundError(f"no table contains {near_text!r}")
    if len(matches) > 1:
        raise AmbiguousTargetError(
            f"{len(matches)} tables contain {near_text!r}; use more specific text"
        )
    return matches[0]


def _locate_table_block(document: "Document", table: "Table") -> Tuple[str, int]:
    for story, root in _story_elements(document):
        for kind, index, element, _sdt, _txbx in _iter_block_elements(story, root):
            if kind == "table" and element is table._tbl:
                return story, index
    raise TargetNotFoundError("table is not part of this document's traversal")


def _cell_at(table: "Table", row: int, column: int) -> "_Cell":
    rows = table.rows
    if not 0 <= row < len(rows):
        raise TargetNotFoundError(f"row {row} does not exist (0..{len(rows) - 1})")
    row_proxy = rows[row]
    offset = column - row_proxy.grid_cols_before
    cells = row_proxy.cells
    if offset < 0 or offset >= len(cells):
        raise TargetNotFoundError(
            f"layout-grid column {column} has no cell in row {row}"
        )
    return cells[offset]


def update_cell(
    table: "Table",
    row: int,
    column: int,
    new_text: str,
    *,
    tracked: bool = False,
    author: Optional[str] = None,
    date: Optional[dt.datetime] = None,
) -> ReplaceResult:
    """Replace the visible text of one cell (0-based `row`/`column`).

    Routed through `Span.replace`: the first run's formatting carries the new
    text, other formatting is untouched, and `tracked=True` produces a real
    revision. Complex tables and multi-paragraph cells are refused.
    """
    if tracked and not author:
        raise ValueError("author is required when tracked=True")
    from docx.search import _validate_writable_text

    _validate_writable_text(new_text, argument="new_text")
    document = _document_of(table)
    _refuse_if_protected(document, "update a table cell")
    cell = _cell_at(table, row, column)
    _refuse_complex_cell(cell._tc, row=row, column=column)
    _refuse_whole_cell_semantics(cell._tc)

    populated = [p for p in cell.paragraphs if p.text]
    if len(populated) > 1:
        raise UnsupportedStructureError(
            "cell holds multiple paragraphs of text; update them individually"
            " through docx.search spans"
        )
    story, block_index = _locate_table_block(document, table)
    target_paragraph = populated[0] if populated else cell.paragraphs[0]
    all_atoms = list(
        _collect_block_atoms(
            story,
            block_index,
            target_paragraph._p,
            skip_text_boxes=False,
            in_txbx=False,
        )
    )
    if any(atom.is_synthetic for atom in all_atoms):
        raise UnsupportedStructureError(
            "cell text contains a tab, break, no-break hyphen, or other"
            " visible run content that cannot be represented by new_text;"
            " update its text segments individually instead"
        )
    atoms = [atom for atom in all_atoms if atom.tag == _T]
    if not atoms:
        if next(cell._tc.iter(_SDT), None) is not None:
            raise UnsupportedStructureError(
                "empty cell contains a content control; set it through"
                " Control.set_value() instead"
            )
        return _fill_empty_cell(
            document, story, target_paragraph, new_text,
            tracked=tracked, author=author, date=date,
        )
    text = "".join(atom.text for atom in atoms)
    span = Span(
        text=text,
        story=story,
        anchor=Anchor(story=story, index=block_index, content_hash=content_hash(text)),
        in_insert=any(a.in_insert for a in atoms),
        in_delete=False,
        in_content_control=any(a.sdt is not None for a in atoms),
        in_text_box=any(a.in_text_box for a in atoms),
        in_field=any(a.in_field for a in atoms),
        crosses_paragraphs=False,
        _document=document,
        _atoms=atoms,
        _start_offset=0,
        _end_offset=len(atoms[-1].text),
        _norm_start=0,
    )
    return span.replace(new_text, tracked=tracked, author=author, date=date)


def _fill_empty_cell(
    document: "Document",
    story: str,
    paragraph,
    new_text: str,
    *,
    tracked: bool,
    author: Optional[str],
    date: Optional[dt.datetime],
) -> ReplaceResult:
    if not tracked:
        paragraph.add_run(new_text)
        return ReplaceResult(
            story=story, deleted_text="", inserted_text=new_text,
            tracked=False, revision_ids=(),
        )
    stamp = date if date is not None else _clock.now()
    revision_id = _next_revision_id(document)
    ins = CT_RunTrackChange.new("w:ins", revision_id, author, stamp)
    ins.add_tracked_run(new_text, None, deleted=False)
    paragraph._p.append(ins)
    return ReplaceResult(
        story=story, deleted_text="", inserted_text=new_text,
        tracked=True, revision_ids=(revision_id,),
    )


def insert_row_after(
    table: "Table",
    row: int,
    values: Sequence[str],
    *,
    copy_format_from: Optional[int] = None,
) -> None:
    """Insert a new row after `row` (0-based), copying formatting from
    `copy_format_from` (default: the anchor row) and filling `values`."""
    _refuse_if_protected(_document_of(table), "insert a table row")
    rows = table.rows
    if not 0 <= row < len(rows):
        raise TargetNotFoundError(f"row {row} does not exist (0..{len(rows) - 1})")
    template_index = row if copy_format_from is None else copy_format_from
    if not 0 <= template_index < len(rows):
        raise TargetNotFoundError(
            f"copy_format_from row {template_index} does not exist"
        )
    _refuse_row_op(table, affected_rows={template_index}, splits_before=row + 1)
    from docx.search import _validate_writable_text

    for position, value in enumerate(values):
        _validate_writable_text(value, argument=f"values[{position}]")
    column_count = len(rows[row].cells)
    if len(values) > column_count:
        raise ValueError(
            f"{len(values)} values for a {column_count}-column table"
        )
    # a horizontally merged template row repeats its merged tc through
    # rows[..].cells, so positional value assignment would silently drop or
    # misplace values — refuse instead
    template_tr = rows[template_index]._tr
    _refuse_tracked_template_row(template_tr, row=template_index)
    if any(
        tc.find(_TC_PR) is not None and tc.find(_TC_PR).find(_GRID_SPAN) is not None
        for tc in _row_cells(template_tr)
    ):
        raise UnsupportedStructureError(
            f"template row {template_index} contains horizontally merged cells"
            " (gridSpan); positional values cannot be assigned unambiguously —"
            " copy formatting from an unmerged row"
        )

    # Populate the copied row while detached. Any refusal or unexpected error
    # leaves the table tree untouched.
    new_tr = copy.deepcopy(rows[template_index]._tr)
    from docx.table import _Cell

    detached_cells = tuple(_Cell(tc, table) for tc in _row_cells(new_tr))
    for index, cell in enumerate(detached_cells):
        _set_cell_text_keeping_format(
            cell, values[index] if index < len(values) else ""
        )
    rows[row]._tr.addnext(new_tr)


def _set_cell_text_keeping_format(cell: "_Cell", text: str) -> None:
    """Replace a copied cell's text, keeping its paragraph and run formatting
    (the upstream `.text` setter would drop the template's run properties)."""
    if next(iter(cell._tc.iter(_SDT)), None) is not None:
        raise UnsupportedStructureError(
            "template cell contains a content control that cannot be populated"
            " safely; nothing was changed"
        )
    if not cell.paragraphs:
        raise UnsupportedStructureError(
            "template cell has no direct paragraph that can be populated safely;"
            " nothing was changed"
        )
    paragraph = cell.paragraphs[0]
    template_rpr = None
    for run in paragraph.runs:
        rpr = run._r.find(qn("w:rPr"))
        if rpr is not None:
            template_rpr = copy.deepcopy(rpr)
            break
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    paragraph.clear()
    run = paragraph.add_run(text)
    if template_rpr is not None:
        run._r.insert(0, template_rpr)


def delete_row(table: "Table", row: int) -> None:
    """Delete row `row` (0-based). The last remaining row is refused —
    a rowless table is not valid WordprocessingML."""
    _refuse_if_protected(_document_of(table), "delete a table row")
    rows = table.rows
    if not 0 <= row < len(rows):
        raise TargetNotFoundError(f"row {row} does not exist (0..{len(rows) - 1})")
    _refuse_row_op(table, affected_rows={row})
    if len(rows) == 1:
        raise UnsupportedStructureError(
            "deleting the last remaining row would leave an invalid table;"
            " remove the table itself instead"
        )
    tr = rows[row]._tr
    tr.getparent().remove(tr)
