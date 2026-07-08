"""Narrow, guarded table operations (paper-docx).

Every operation refuses loudly on tables with merged cells (`vMerge`,
`gridSpan`) or nested tables — the failure mode this module exists to prevent
is a "clever" edit that silently reshuffles a complex table. Cell updates
route through `Span.replace` (Phase 5), so formatting preservation, tracked
changes and refusal atomicity all come from the same machinery as body text.
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


def _is_complex(table: "Table") -> bool:
    return bool(table._tbl.xpath(".//w:vMerge | .//w:gridSpan | .//w:tbl//w:tbl"))


def _refuse_complex(table: "Table") -> None:
    if _is_complex(table):
        raise UnsupportedStructureError(
            "table has merged cells or nested tables; refusing rather than"
            " silently corrupting its structure"
        )


def _document_of(table: "Table") -> "Document":
    part = table.part
    document = getattr(part, "document", None)
    if document is None:
        raise UnsupportedStructureError(
            "tables outside the main document story are not supported in v0"
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
    cells = rows[row].cells
    if not 0 <= column < len(cells):
        raise TargetNotFoundError(
            f"column {column} does not exist (0..{len(cells) - 1})"
        )
    return cells[column]


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
    _refuse_complex(table)
    document = _document_of(table)
    cell = _cell_at(table, row, column)

    populated = [p for p in cell.paragraphs if p.text]
    if len(populated) > 1:
        raise UnsupportedStructureError(
            "cell holds multiple paragraphs of text; update them individually"
            " through docx.search spans"
        )
    story, block_index = _locate_table_block(document, table)
    target_paragraph = populated[0] if populated else cell.paragraphs[0]
    atoms = [
        atom
        for atom in _collect_block_atoms(
            story, block_index, target_paragraph._p, skip_text_boxes=False, in_txbx=False
        )
        if atom.tag == _T
    ]
    if not atoms:
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
    _refuse_complex(table)
    rows = table.rows
    if not 0 <= row < len(rows):
        raise TargetNotFoundError(f"row {row} does not exist (0..{len(rows) - 1})")
    template_index = row if copy_format_from is None else copy_format_from
    if not 0 <= template_index < len(rows):
        raise TargetNotFoundError(
            f"copy_format_from row {template_index} does not exist"
        )
    column_count = len(rows[row].cells)
    if len(values) > column_count:
        raise ValueError(
            f"{len(values)} values for a {column_count}-column table"
        )

    # -- validated; mutate --
    new_tr = copy.deepcopy(rows[template_index]._tr)
    rows[row]._tr.addnext(new_tr)
    for index, cell in enumerate(table.rows[row + 1].cells):
        cell.text = values[index] if index < len(values) else ""


def delete_row(table: "Table", row: int) -> None:
    """Delete row `row` (0-based). The last remaining row is refused —
    a rowless table is not valid WordprocessingML."""
    _refuse_complex(table)
    rows = table.rows
    if not 0 <= row < len(rows):
        raise TargetNotFoundError(f"row {row} does not exist (0..{len(rows) - 1})")
    if len(rows) == 1:
        raise UnsupportedStructureError(
            "deleting the last remaining row would leave an invalid table;"
            " remove the table itself instead"
        )
    tr = rows[row]._tr
    tr.getparent().remove(tr)
