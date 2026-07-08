"""Revision-resolution completion (PLAN-v0.11 Phases 0-2).

Phase 0 pins that v0.1's detection fires on every revision shape the v0.11
pipeline builds on (the ground we stand on); Phases 1-2 pin the resolution
semantics themselves against the hand-computed accepted ground truth.
"""

from __future__ import annotations

from pathlib import Path

import pytest

import docx
from docx.errors import UnsupportedStructureError

from .harness.contract import save_and_reopen
from .harness.paths import fixture_path

MULTIROUND = "generated/redline/multiround.docx"
MULTIROUND_ACCEPTED = "generated/redline/multiround-accepted.docx"
ROW_REVISIONS = "generated/feature-isolated/row-revisions.docx"
FORMAT_RICH = "generated/feature-isolated/format-changes-rich.docx"
PARAGRAPH_MERGE = "generated/feature-isolated/paragraph-merge.docx"
MINIMAL_TABLE = "generated/feature-isolated/table-merged-nested.docx"
GAUNTLET = "generated/gauntlet/gauntlet.docx"


def _doc(relpath: str):
    return docx.Document(str(fixture_path(relpath)))


class DescribePhase0Detection:
    """v0.1 detection fires on every shape v0.11 resolves (Phase 0 gate)."""

    def it_enumerates_every_revision_in_the_multiround_redline(self):
        revisions = _doc(MULTIROUND).revisions
        assert len(revisions) == 20
        assert {r.author for r in revisions} == {"Alice Editor", "Bob Reviewer"}

    def it_enumerates_the_move_pair_with_mark_stamps(self):
        revisions = _doc(MULTIROUND).revisions
        move_from = [r for r in revisions if r.revision_type == "move_from"]
        move_to = [r for r in revisions if r.revision_type == "move_to"]
        assert len(move_from) == 2 and len(move_to) == 2  # wrapper + mark each
        assert any(r.is_paragraph_mark for r in move_from)
        assert any(r.is_paragraph_mark for r in move_to)

    def it_enumerates_rich_format_changes_including_the_mark_change(self):
        revisions = _doc(FORMAT_RICH).revisions
        format_changes = [r for r in revisions if r.revision_type == "format_change"]
        assert len(format_changes) == 3
        assert sum(1 for r in format_changes if r.is_paragraph_mark) == 1

    def it_enumerates_both_paragraph_mark_revisions(self):
        revisions = _doc(PARAGRAPH_MERGE).revisions
        marks = [(r.revision_type, r.author) for r in revisions if r.is_paragraph_mark]
        assert sorted(marks) == [
            ("deletion", "Bob Reviewer"),
            ("insertion", "Alice Editor"),
        ]

    def it_enumerates_all_ten_row_revision_nodes(self):
        revisions = _doc(ROW_REVISIONS).revisions
        assert len(revisions) == 10  # 2 trPr markers + 4 content + 4 mark stamps
        assert {r.author for r in revisions} == {"Alice Editor", "Bob Reviewer"}

    def it_finds_zero_revisions_in_the_accepted_ground_truth(self):
        assert len(_doc(MULTIROUND_ACCEPTED).revisions) == 0


def _rescan(document) -> dict:
    from docx.revision import _remaining_markup

    return _remaining_markup(document)


def _body_texts(document) -> list:
    from docx.story import iter_blocks

    return [b.text for b in iter_blocks(document) if b.story == "word/document.xml"]


class DescribeFormatChangeResolution:
    """Phase 1: w:rPrChange/w:pPrChange accept and reject."""

    def it_accepts_a_run_format_change_keeping_current_properties(
        self, tmp_path: Path
    ):
        document = _doc(FORMAT_RICH)
        for revision in document.revisions:
            if not revision.is_paragraph_mark and "Delivery" in revision.text:
                revision.accept()
                break
        else:
            pytest.fail("run format change not found")
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        paragraph = next(
            p for p in reopened.paragraphs if "Delivery" in p.text
        )
        assert paragraph.runs[0].font.bold is True
        assert all(
            r.revision_type != "format_change" or "Delivery" not in r.text
            for r in reopened.revisions
        )

    def it_rejects_a_run_format_change_restoring_stored_properties(
        self, tmp_path: Path
    ):
        document = _doc(FORMAT_RICH)
        for revision in document.revisions:
            if not revision.is_paragraph_mark and "Delivery" in revision.text:
                revision.reject()
                break
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        run = next(p for p in reopened.paragraphs if "Delivery" in p.text).runs[0]
        assert run.font.bold is None  # bold dropped
        assert run.font.italic is True  # stored previous restored
        assert run.font.size is not None and run.font.size.pt == 14  # sz 28 half-points

    def it_accepts_a_paragraph_format_change(self, tmp_path: Path):
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        document = _doc(FORMAT_RICH)
        for revision in document.revisions:
            if "right-aligned" in revision.text:
                revision.accept()
                break
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        paragraph = next(p for p in reopened.paragraphs if "right-aligned" in p.text)
        assert paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER

    def it_rejects_a_paragraph_format_change_restoring_stored_properties(
        self, tmp_path: Path
    ):
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        document = _doc(FORMAT_RICH)
        for revision in document.revisions:
            if "right-aligned" in revision.text:
                revision.reject()
                break
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        paragraph = next(p for p in reopened.paragraphs if "right-aligned" in p.text)
        assert paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT

    def it_resolves_the_paragraph_mark_format_change_both_ways(
        self, tmp_path: Path
    ):
        for accept in (True, False):
            document = _doc(FORMAT_RICH)
            mark = next(r for r in document.revisions if r.is_paragraph_mark)
            assert mark.revision_type == "format_change"
            (mark.accept if accept else mark.reject)()
            reopened = save_and_reopen(document, tmp_path / f"out-{accept}.docx")
            # only the mark change resolved; the run/paragraph changes remain
            assert _rescan(reopened) == {"pPrChange": 1, "rPrChange": 1}
            texts = _body_texts(reopened)
            assert "The paragraph mark itself was re-formatted." in texts

    def it_accepts_all_and_a_rescan_finds_zero_markup(self, tmp_path: Path):
        document = _doc(FORMAT_RICH)
        assert document.revisions.accept_all() == 3
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert _rescan(reopened) == {}
        assert len(reopened.revisions) == 0

    def it_rejects_all_and_a_rescan_finds_zero_markup(self, tmp_path: Path):
        document = _doc(FORMAT_RICH)
        assert document.revisions.reject_all() == 3
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert _rescan(reopened) == {}


class DescribeRowRevisionResolution:
    """Phase 1: w:trPr row markers — the ghost-row defect closed."""

    def _grid(self, document) -> list:
        return [
            [cell.text for cell in row.cells] for row in document.tables[0].rows
        ]

    def it_accepts_all_keeping_inserted_and_dropping_deleted_rows(
        self, tmp_path: Path
    ):
        document = _doc(ROW_REVISIONS)
        document.revisions.accept_all()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert self._grid(reopened) == [["Item", "Amount"], ["Filing fee", "$100"]]
        assert _rescan(reopened) == {}

    def it_rejects_all_restoring_the_original_table(self, tmp_path: Path):
        document = _doc(ROW_REVISIONS)
        document.revisions.reject_all()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert self._grid(reopened) == [["Item", "Amount"], ["Old charge", "$50"]]
        assert _rescan(reopened) == {}

    def it_classifies_row_markers_distinctly(self):
        revisions = _doc(ROW_REVISIONS).revisions
        types = sorted(
            r.revision_type
            for r in revisions
            if r.revision_type.startswith("row_")
        )
        assert types == ["row_deletion", "row_insertion"]

    def it_accepts_a_single_row_insertion_removing_only_the_marker(
        self, tmp_path: Path
    ):
        document = _doc(ROW_REVISIONS)
        row_ins = next(
            r for r in document.revisions if r.revision_type == "row_insertion"
        )
        row_ins.accept()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert len(reopened.tables[0].rows) == 3  # row kept, marker gone
        assert "row_insertion" not in {r.revision_type for r in reopened.revisions}

    def it_rejects_a_single_row_insertion_removing_the_row(self, tmp_path: Path):
        document = _doc(ROW_REVISIONS)
        row_ins = next(
            r for r in document.revisions if r.revision_type == "row_insertion"
        )
        row_ins.reject()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        # inserted row gone; the deleted-marked row remains (its delText
        # content is invisible to upstream cell.text — still pending)
        assert len(reopened.tables[0].rows) == 2
        assert "Filing fee" not in reopened.element.xml
        assert "row_deletion" in {r.revision_type for r in reopened.revisions}

    def it_refuses_to_empty_a_table_atomically(self):
        import copy as copy_mod

        document = _doc(ROW_REVISIONS)
        table = document.tables[0]._tbl
        # strip the table down to ONLY the deleted-marked row
        for row in list(table.tr_lst):
            if "Old charge" not in "".join(row.itertext()):
                table.remove(row)
        before = copy_mod.deepcopy(document.element.xml)
        with pytest.raises(UnsupportedStructureError, match="only row|every row"):
            document.revisions.accept_all()
        assert document.element.xml == before  # nothing half-resolved


class DescribeParagraphMarkResolution:
    """Phase 1: the pilcrow itself, both directions explicitly."""

    def it_accepts_a_deleted_mark_merging_with_the_next_paragraph(
        self, tmp_path: Path
    ):
        document = _doc(PARAGRAPH_MERGE)
        mark = next(
            r
            for r in document.revisions
            if r.is_paragraph_mark and r.revision_type == "deletion"
        )
        mark.accept()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        texts = _body_texts(reopened)
        assert "This sentence continues onto the following line." in texts
        assert "onto the following line." not in texts

    def it_rejects_a_deleted_mark_keeping_both_paragraphs(self, tmp_path: Path):
        document = _doc(PARAGRAPH_MERGE)
        mark = next(
            r
            for r in document.revisions
            if r.is_paragraph_mark and r.revision_type == "deletion"
        )
        mark.reject()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        texts = _body_texts(reopened)
        assert "This sentence continues " in texts
        assert "onto the following line." in texts

    def it_accepts_an_inserted_mark_keeping_the_split(self, tmp_path: Path):
        document = _doc(PARAGRAPH_MERGE)
        mark = next(
            r
            for r in document.revisions
            if r.is_paragraph_mark and r.revision_type == "insertion"
        )
        mark.accept()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        texts = _body_texts(reopened)
        assert "A tracked split divides " in texts
        assert "this once-single sentence." in texts

    def it_rejects_an_inserted_mark_merging_the_split_back(self, tmp_path: Path):
        document = _doc(PARAGRAPH_MERGE)
        mark = next(
            r
            for r in document.revisions
            if r.is_paragraph_mark and r.revision_type == "insertion"
        )
        mark.reject()
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        texts = _body_texts(reopened)
        assert "A tracked split divides this once-single sentence." in texts

    def it_resolves_the_whole_fixture_both_ways_to_zero_markup(
        self, tmp_path: Path
    ):
        for accept in (True, False):
            document = _doc(PARAGRAPH_MERGE)
            resolver = (
                document.revisions.accept_all if accept else document.revisions.reject_all
            )
            assert resolver() == 2
            reopened = save_and_reopen(document, tmp_path / f"out-{accept}.docx")
            assert _rescan(reopened) == {}


class DescribeExoticTypesStayRefused:
    """Phase 1: everything not resolved is enumerated and refused BY NAME."""

    @pytest.mark.parametrize(
        ("markup", "expected_type"),
        [
            (
                '<w:tblPrChange w:id="900" w:author="A" w:date="2026-06-01T09:30:00Z">'
                "<w:tblPr/></w:tblPrChange>",
                "table_property_change",
            ),
            (
                '<w:tcPrChange w:id="901" w:author="A" w:date="2026-06-01T09:30:00Z">'
                "<w:tcPr/></w:tcPrChange>",
                "table_property_change",
            ),
            (
                '<w:cellMerge w:id="902" w:author="A" w:date="2026-06-01T09:30:00Z"/>',
                "cell_revision",
            ),
        ],
    )
    def it_enumerates_and_refuses_each_exotic_type_by_name(
        self, markup: str, expected_type: str
    ):
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls, qn

        document = _doc(MINIMAL_TABLE)
        table = document.tables[0]._tbl
        first_row = table.tr_lst[0]
        cell = first_row.tc_lst[0]
        if "tblPrChange" in markup:
            table.tblPr.append(parse_xml(f"<w:x {nsdecls('w')}>{markup}</w:x>")[0])
        else:
            tc_pr = cell.get_or_add_tcPr()
            tc_pr.append(parse_xml(f"<w:x {nsdecls('w')}>{markup}</w:x>")[0])
        revisions = document.revisions
        exotic = [r for r in revisions if r.revision_type == expected_type]
        assert exotic, f"{expected_type} not enumerated"
        with pytest.raises(UnsupportedStructureError, match=expected_type):
            exotic[0].accept()
        with pytest.raises(UnsupportedStructureError, match="not resolve"):
            revisions.accept_all()
        assert revisions.remaining_unsupported() == {expected_type: 1}

