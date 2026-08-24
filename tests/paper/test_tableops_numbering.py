"""Tests for docx.tableops and docx.numbering."""

from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path

import pytest
from lxml import etree

import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.numbering import (
    apply_list_style,
    apply_numbering,
    ensure_bullet_definition,
    list_numbering,
)
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.shared import Inches
from docx.tableops import delete_row, find_table, insert_row_after, update_cell

from .harness.contract import (
    assert_changed_parts,
    assert_refusal_atomic,
    save_and_reopen,
)
from .harness.paths import fixture_path

MINIMAL = "generated/minimal-clean/minimal.docx"
COMPLEX_TABLE = "generated/feature-isolated/table-merged-nested.docx"
NUMBERING = "generated/feature-isolated/numbering-custom.docx"
NUMBERING_LO = "libreoffice/feature-isolated/numbering-custom.docx"

FROZEN = dt.datetime(2026, 7, 7, 12, 0, 0, tzinfo=dt.timezone.utc)
W = nsdecls("w")


def _doc(relpath: str):
    return docx.Document(str(fixture_path(relpath)))


def _doc_with_simple_table():
    document = _doc(MINIMAL)
    table = document.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            table.cell(r, c).text = f"cell {r}{c}"
    return document, table


class DescribeFindTable:
    @pytest.mark.parametrize(
        ("near_text", "match"),
        [("", "exact"), ("   \t", "normalized")],
    )
    def it_refuses_targets_that_contain_no_searchable_text(self, near_text: str, match: str):
        document, _ = _doc_with_simple_table()

        with pytest.raises(TargetNotFoundError, match="searchable text"):
            find_table(document, near_text=near_text, match=match)

    def it_finds_the_table_by_exact_cell_text_by_default(self):
        document, table = _doc_with_simple_table()
        found = find_table(document, near_text="cell 10")
        assert found.cell(1, 0).text == table.cell(1, 0).text

    def it_keeps_default_exact_matching_case_and_typography_sensitive(self):
        document, table = _doc_with_simple_table()
        table.cell(1, 0).text = "Résumé – Total"
        for near_text in ("RÉSUMÉ – TOTAL", "Résumé - Total"):
            with pytest.raises(TargetNotFoundError, match="no table"):
                find_table(document, near_text=near_text)

    def it_supports_explicit_normalized_matching(self):
        document, table = _doc_with_simple_table()
        table.cell(1, 0).text = "Résumé – Total"
        found = find_table(
            document,
            near_text="RÉSUMÉ - TOTAL",
            match="normalized",
        )
        assert found.cell(1, 0).text == table.cell(1, 0).text

    def it_does_not_match_an_exact_query_across_adjacent_cells(self):
        document, _ = _doc_with_simple_table()
        with pytest.raises(TargetNotFoundError, match="no table"):
            find_table(document, near_text="cell 00\ncell 01")

    def it_does_not_match_a_normalized_query_across_adjacent_cells(self):
        document = _doc(MINIMAL)
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "Account"
        table.cell(0, 1).text = "Balance"
        with pytest.raises(TargetNotFoundError, match="no table"):
            find_table(
                document,
                near_text="Account Balance",
                match="normalized",
            )

    def it_matches_paragraph_boundaries_within_one_cell(self):
        document = _doc(MINIMAL)
        table = document.add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        cell.text = "Account"
        cell.add_paragraph("Balance")

        exact = find_table(document, near_text="Account\nBalance")
        normalized = find_table(
            document,
            near_text="ACCOUNT   BALANCE",
            match="normalized",
        )
        assert exact.cell(0, 0).text == table.cell(0, 0).text
        assert normalized.cell(0, 0).text == table.cell(0, 0).text

    def it_counts_a_merged_physical_cell_only_once(self):
        document = _doc(MINIMAL)
        table = document.add_table(rows=1, cols=2)
        cell = table.cell(0, 0).merge(table.cell(0, 1))
        cell.text = "merged marker"

        found = find_table(document, near_text="merged marker")
        assert found.cell(0, 0).text == table.cell(0, 0).text

    def it_counts_a_table_only_once_when_several_cells_match(self):
        document = _doc(MINIMAL)
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "marker in first cell"
        table.cell(0, 1).text = "marker in second cell"

        found = find_table(document, near_text="marker", match="exact")
        assert found.cell(0, 1).text == table.cell(0, 1).text

    def it_refuses_when_no_table_matches(self):
        document, _ = _doc_with_simple_table()
        with pytest.raises(TargetNotFoundError, match="no table"):
            find_table(
                document,
                near_text="NOTHING LIKE THIS",
                match="normalized",
            )

    def it_counts_ambiguity_between_tables(self):
        document, first = _doc_with_simple_table()
        first.cell(0, 0).text = "cell 10 in another cell"
        second = document.add_table(rows=1, cols=1)
        second.cell(0, 0).text = "cell 10 duplicate"
        with pytest.raises(AmbiguousTargetError, match="2 tables"):
            find_table(document, near_text="cell 10", match="exact")

    def it_rejects_an_invalid_match_policy(self):
        document, _ = _doc_with_simple_table()
        with pytest.raises(
            ValueError,
            match=r"match must be one of .* got 'prefix'",
        ):
            find_table(document, near_text="cell 10", match="prefix")


class DescribeUpdateCell:
    def it_replaces_cell_text_through_the_span_machinery(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        result = update_cell(table, 0, 1, "updated value")
        assert result.deleted_text == "cell 01"
        assert result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert reopened.tables[0].cell(0, 1).text == "updated value"

    def it_fills_an_empty_cell(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        table.cell(1, 1).paragraphs[0].clear()
        result = update_cell(table, 1, 1, "was empty")
        assert not result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert reopened.tables[0].cell(1, 1).text == "was empty"

    def it_supports_tracked_updates(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        result = update_cell(table, 0, 0, "cell 99", tracked=True, author="Carol QA", date=FROZEN)
        assert result.tracked
        assert result.revision_ids
        assert not result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert reopened.tables[0].cell(0, 0).text.startswith("cell")
        revisions = reopened.revisions
        assert {r.author for r in revisions} == {"Carol QA"}
        reopened.revisions.accept_all()
        assert reopened.tables[0].cell(0, 0).text == "cell 99"

    def it_refuses_merged_target_cells_atomically(self):
        document = _doc(COMPLEX_TABLE)
        assert_refusal_atomic(
            document,
            lambda doc: update_cell(doc.tables[0], 0, 0, "x"),  # gridSpan cell
            UnsupportedStructureError,
        )
        assert_refusal_atomic(
            document,
            lambda doc: update_cell(doc.tables[0], 1, 2, "x"),  # vMerge cell
            UnsupportedStructureError,
        )
        assert_refusal_atomic(
            document,
            lambda doc: update_cell(doc.tables[0], 2, 0, "x"),  # nested table
            UnsupportedStructureError,
        )

    def and_it_updates_plain_cells_of_the_same_table(self, tmp_path: Path):
        """Cell-wise guards: the merge no longer poisons the table."""
        document = _doc(COMPLEX_TABLE)
        update_cell(document.tables[0], 1, 0, "updated plain cell")
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert reopened.tables[0].cell(1, 0).text == "updated plain cell"

    def it_refuses_out_of_range_addresses(self):
        _, table = _doc_with_simple_table()
        with pytest.raises(TargetNotFoundError, match="row 7"):
            update_cell(table, 7, 0, "x")
        with pytest.raises(TargetNotFoundError, match="column 9"):
            update_cell(table, 0, 9, "x")

    def it_requires_an_author_when_tracked(self):
        _, table = _doc_with_simple_table()
        with pytest.raises(ValueError, match="author"):
            update_cell(table, 0, 0, "x", tracked=True)

    def it_replaces_a_uniformly_formatted_fragmented_cell(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.clear()
        paragraph.add_run("cell ").bold = True
        paragraph.add_run("value").bold = True

        result = update_cell(table, 0, 0, "updated")

        assert result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "uniform-cell.docx")
        runs = reopened.tables[0].cell(0, 0).paragraphs[0].runs
        assert "".join(run.text for run in runs if run.bold) == "updated"

    def it_updates_a_mixed_format_cell_using_the_start_run(self):
        document, table = _doc_with_simple_table()
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.clear()
        paragraph.add_run("Label: ").bold = True
        paragraph.add_run("value")

        result = update_cell(table, 0, 0, "updated")

        assert not result.preserved_formatting_regions
        assert paragraph.text == "updated"
        assert paragraph.runs[0].bold

    def it_refuses_a_marker_divided_cell_atomically(self):
        document, table = _doc_with_simple_table()
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.clear()
        first = paragraph.add_run("cell ")
        first._r.addnext(parse_xml(f'<w:proofErr {W} w:type="spellStart"/>'))
        paragraph.add_run("value")

        assert_refusal_atomic(
            document,
            lambda _document: update_cell(table, 0, 0, "updated"),
            UnsupportedStructureError,
        )

    def it_accepts_single_node_complete_cell_formatting(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        run = table.cell(0, 0).paragraphs[0].runs[0]
        run._r.get_or_add_rPr().append(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:shd {W} w:fill="FFFF00"/>')
        )

        result = update_cell(table, 0, 0, "updated")

        assert result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "shaded-cell.docx")
        paragraph = reopened.tables[0].cell(0, 0).paragraphs[0]
        assert paragraph.text == "updated"
        assert paragraph._p.xpath(  # pyright: ignore[reportPrivateUsage]
            'w:r/w:rPr/w:shd[@w:fill="FFFF00"]'
        )

    def it_refuses_ambiguous_repeated_cell_affixes_atomically(self):
        document, table = _doc_with_simple_table()
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.clear()
        paragraph.add_run("Term").bold = True
        paragraph.add_run("Term")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: update_cell(table, 0, 0, "Term"),
            UnsupportedStructureError,
        )

        assert "exact affix alignment is ambiguous" in str(refusal)

    def it_updates_a_character_style_divided_cell_using_the_start_style(self):
        document, table = _doc_with_simple_table()
        first_style = document.styles.add_style("Cell First", WD_STYLE_TYPE.CHARACTER)
        first_style.font.bold = True
        second_style = document.styles.add_style("Cell Second", WD_STYLE_TYPE.CHARACTER)
        second_style.font.italic = True
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.clear()
        paragraph.add_run("cell ", style=first_style)
        paragraph.add_run("value", style=second_style)

        result = update_cell(table, 0, 0, "updated")

        assert not result.preserved_formatting_regions
        assert paragraph.text == "updated"
        assert paragraph.runs[0].style.name == "Cell First"

    def it_refuses_a_hyperlink_scope_divided_cell_atomically(self):
        document, table = _doc_with_simple_table()
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.clear()
        paragraph._p.append(
            parse_xml(
                f'<w:hyperlink {W} w:anchor="target">'
                "<w:r><w:t>cell </w:t></w:r>"
                "</w:hyperlink>"
            )
        )
        paragraph.add_run("value")

        assert_refusal_atomic(
            document,
            lambda _document: update_cell(table, 0, 0, "updated"),
            BoundaryViolationError,
        )


class DescribeRowOperations:
    def it_inserts_a_row_copying_neighbor_formatting(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        insert_row_after(table, 0, ["new a", "new b"])
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        grid = [[c.text for c in row.cells] for row in reopened.tables[0].rows]
        assert grid == [["cell 00", "cell 01"], ["new a", "new b"], ["cell 10", "cell 11"]]

    def it_pads_missing_values_with_empty_cells(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        insert_row_after(table, 1, ["only one"])
        reopened = save_and_reopen(document, tmp_path / "padded.docx")
        assert [c.text for c in reopened.tables[0].rows[2].cells] == [
            "only one",
            "",
        ]

    def it_preserves_one_complete_uniform_formatting_outcome(self, tmp_path: Path):
        document = _doc(MINIMAL)
        table = document.add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        table.rows[0].height = Inches(0.4)
        for text in ("uniform ", "template"):
            run = paragraph.add_run(text)
            rpr = run._r.get_or_add_rPr()  # pyright: ignore[reportPrivateUsage]
            highlight = OxmlElement("w:highlight")
            highlight.set(qn("w:val"), "yellow")
            language = OxmlElement("w:lang")
            language.set(qn("w:val"), "en-US")
            rpr.extend((highlight, language))
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "D9EAF7")
        cell._tc.get_or_add_tcPr().append(  # pyright: ignore[reportPrivateUsage]
            shading
        )

        original = tmp_path / "before.docx"
        modified = tmp_path / "after.docx"
        document.save(str(original))
        tc_pr = cell._tc.tcPr  # pyright: ignore[reportPrivateUsage]
        p_pr = paragraph._p.pPr  # pyright: ignore[reportPrivateUsage]
        r_pr = paragraph.runs[0]._r.rPr  # pyright: ignore[reportPrivateUsage]
        tr_pr = table.rows[0]._tr.trPr  # pyright: ignore[reportPrivateUsage]
        assert tc_pr is not None
        assert p_pr is not None
        assert r_pr is not None
        assert tr_pr is not None
        expected_tcpr = etree.tostring(tc_pr, method="c14n")
        expected_ppr = etree.tostring(p_pr, method="c14n")
        expected_rpr = etree.tostring(r_pr, method="c14n")
        expected_trpr = etree.tostring(tr_pr, method="c14n")

        insert_row_after(table, 0, ["replacement value"])
        reopened = save_and_reopen(document, modified)
        inserted = reopened.tables[0].cell(1, 0)
        assert inserted.text == "replacement value"
        assert [run.text for run in inserted.paragraphs[0].runs] == ["replacement value"]
        inserted_tcpr = inserted._tc.tcPr  # pyright: ignore[reportPrivateUsage]
        inserted_ppr = inserted.paragraphs[0]._p.pPr  # pyright: ignore[reportPrivateUsage]
        inserted_rpr = inserted.paragraphs[0].runs[0]._r.rPr  # pyright: ignore[reportPrivateUsage]
        inserted_trpr = reopened.tables[0].rows[1]._tr.trPr  # pyright: ignore[reportPrivateUsage]
        assert inserted_tcpr is not None
        assert inserted_ppr is not None
        assert inserted_rpr is not None
        assert inserted_trpr is not None
        assert etree.tostring(inserted_tcpr, method="c14n") == expected_tcpr
        assert etree.tostring(inserted_ppr, method="c14n") == expected_ppr
        assert etree.tostring(inserted_rpr, method="c14n") == expected_rpr
        assert etree.tostring(inserted_trpr, method="c14n") == expected_trpr
        assert_changed_parts(original, modified, {"word/document.xml"})

    def it_rejects_too_many_values(self):
        _, table = _doc_with_simple_table()
        with pytest.raises(ValueError, match="values"):
            insert_row_after(table, 0, ["a", "b", "c"])

    def it_deletes_a_row(self, tmp_path: Path):
        document, table = _doc_with_simple_table()
        delete_row(table, 0)
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        grid = [[c.text for c in row.cells] for row in reopened.tables[0].rows]
        assert grid == [["cell 10", "cell 11"]]

    def it_refuses_deleting_the_last_row(self):
        document, table = _doc_with_simple_table()
        delete_row(table, 0)
        assert_refusal_atomic(
            document,
            lambda doc: delete_row(table, 0),
            UnsupportedStructureError,
        )

    def it_refuses_row_ops_intersecting_a_vertical_merge(self):
        document = _doc(COMPLEX_TABLE)
        with pytest.raises(UnsupportedStructureError, match="vertical merge"):
            insert_row_after(document.tables[0], 1, ["x"])  # template row merged
        with pytest.raises(UnsupportedStructureError, match="vertical merge"):
            delete_row(document.tables[0], 1)  # vMerge restart row
        with pytest.raises(UnsupportedStructureError, match="split"):
            # clean template, but the insertion point splits rows 1-2's merge
            insert_row_after(document.tables[0], 1, ["x"], copy_format_from=0)

    def it_refuses_horizontally_merged_template_rows(self):
        """A gridSpan template row repeats its merged tc through .cells, so
        positional values would silently misassign."""
        document = _doc(COMPLEX_TABLE)
        with pytest.raises(UnsupportedStructureError, match="gridSpan"):
            insert_row_after(document.tables[0], 0, ["a", "b", "c"])

    @pytest.mark.parametrize(
        ("omitted_side", "property_tag"),
        [("start", "w:gridBefore"), ("end", "w:gridAfter")],
    )
    def it_refuses_nonrectangular_copied_templates_before_mutation(
        self, omitted_side: str, property_tag: str
    ):
        document = _doc(MINIMAL)
        table = document.add_table(rows=1, cols=3)
        row = table.rows[0]
        omitted = row._tr.tc_lst[0 if omitted_side == "start" else -1]  # pyright: ignore[reportPrivateUsage]
        row._tr.remove(omitted)  # pyright: ignore[reportPrivateUsage]
        grid_omission = OxmlElement(property_tag)
        grid_omission.set(qn("w:val"), "1")
        row._tr.get_or_add_trPr().append(grid_omission)  # pyright: ignore[reportPrivateUsage]
        before = table._tbl.xml  # pyright: ignore[reportPrivateUsage]

        with pytest.raises(UnsupportedStructureError, match="nonrectangular"):
            insert_row_after(table, 0, ["a", "b", "c"])

        assert table._tbl.xml == before  # pyright: ignore[reportPrivateUsage]
        assert len(table.rows) == 1

    def and_it_allows_deleting_a_grid_span_row(self, tmp_path: Path):
        """Deleting a whole horizontally merged row is unambiguous."""
        document = _doc(COMPLEX_TABLE)
        delete_row(document.tables[0], 0)
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        assert len(reopened.tables[0].rows) == 2


class DescribeListNumbering:
    def it_reports_definitions_with_levels(self):
        report = list_numbering(_doc(NUMBERING))
        custom = next(d for d in report.definitions if d.num_id == 42)
        assert custom.abstract_num_id == 90
        assert [(lvl.level, lvl.num_fmt) for lvl in custom.levels] == [
            (0, "decimal"),
            (1, "lowerLetter"),
        ]

    def it_reports_numbered_paragraphs_with_anchors_and_levels(self):
        report = list_numbering(_doc(NUMBERING))
        assert [(p.num_id, p.level, p.text) for p in report.numbered_paragraphs] == [
            (42, 0, "First numbered item"),
            (42, 0, "Second numbered item"),
            (42, 1, "Nested lettered item"),
        ]

    def it_reads_the_libreoffice_remapped_definition(self):
        """LibreOffice round-trip remapped numId 42 -> 7 (frozen in the sidecar)."""
        report = list_numbering(_doc(NUMBERING_LO))
        assert [(p.num_id, p.text) for p in report.numbered_paragraphs] == [
            (7, "First numbered item"),
            (7, "Second numbered item"),
            (7, "Nested lettered item"),
        ]

    def it_serializes_deterministically(self):
        import json

        payload = list_numbering(_doc(NUMBERING)).to_dict()
        assert payload["schema"] == "paper_numbering"
        assert payload["version"] == 2
        assert json.dumps(payload) == json.dumps(list_numbering(_doc(NUMBERING)).to_dict())

    def it_reports_numbered_table_cell_paragraphs_without_top_level_block_identity(self):
        document = docx.Document()
        table = document.add_table(rows=1, cols=1)
        paragraph = table.cell(0, 0).paragraphs[0]
        paragraph.add_run("numbered cell")
        num_id = ensure_bullet_definition(document)
        apply_numbering(paragraph, num_id=num_id)
        (reported,) = list_numbering(document).numbered_paragraphs
        assert reported.text == "numbered cell"
        assert reported.table_cell == (0, 0, 0)


class DescribeApplyNumbering:
    def it_applies_an_existing_definition(self, tmp_path: Path):
        document = _doc(NUMBERING)
        paragraph = document.add_paragraph("Newly numbered item")
        apply_numbering(paragraph, num_id=42, level=1)
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        report = list_numbering(reopened)
        assert (42, 1, "Newly numbered item") in [
            (p.num_id, p.level, p.text) for p in report.numbered_paragraphs
        ]

    def it_refuses_undefined_definitions(self):
        document = _doc(NUMBERING)
        paragraph = document.add_paragraph("target")
        assert_refusal_atomic(
            document,
            lambda doc: apply_numbering(paragraph, num_id=999),
            TargetNotFoundError,
        )

    def it_refuses_undefined_levels(self):
        document = _doc(NUMBERING)
        paragraph = document.add_paragraph("target")
        with pytest.raises(TargetNotFoundError, match="level 7"):
            apply_numbering(paragraph, num_id=42, level=7)

    def it_refuses_when_no_numbering_part_exists(self, tmp_path: Path):
        import re

        stripped = tmp_path / "no-numbering.docx"
        with zipfile.ZipFile(fixture_path(MINIMAL)) as zin, zipfile.ZipFile(
            stripped, "w"
        ) as zout:
            for name in zin.namelist():
                if "numbering" in name:
                    continue
                blob = zin.read(name)
                if name == "word/_rels/document.xml.rels":
                    blob = re.sub(rb"<Relationship [^>]*numbering[^>]*/>", b"", blob)
                zout.writestr(name, blob)
        document = docx.Document(str(stripped))
        paragraph = document.add_paragraph("target")
        with pytest.raises(TargetNotFoundError, match="does not exist"):
            apply_numbering(paragraph, num_id=1)


class DescribeApplyListStyle:
    def it_applies_an_existing_style(self):
        document = _doc(MINIMAL)
        paragraph = document.add_paragraph("styled")
        apply_list_style(paragraph, "Heading 2")
        assert paragraph.style.name == "Heading 2"

    def it_refuses_undefined_styles(self):
        document = _doc(MINIMAL)
        paragraph = document.add_paragraph("styled")
        assert_refusal_atomic(
            document,
            lambda doc: apply_list_style(paragraph, "No Such List Style"),
            TargetNotFoundError,
        )
