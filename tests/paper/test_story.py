"""Tests for docx.story — visibility-complete traversal (Phase 3).

Sidecar-driven: counts and texts per fixture assert against the hand-verified
ground truth; the gauntlet proves nothing visible is missed; the LibreOffice
textbox fixture proves mc:AlternateContent fallbacks are not double-counted.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import List

import pytest

import docx
from docx._normalize import normalize_text
from docx.story import Block, iter_blocks, outline, story_parts

from .harness.paths import fixture_path, sidecar_path

TRACKED = "generated/feature-isolated/tracked-ins-del.docx"
TEXTBOX = "generated/feature-isolated/textbox.docx"
TEXTBOX_LO = "libreoffice/feature-isolated/textbox.docx"
CONTROLS = "generated/feature-isolated/content-control.docx"
NOTES = "generated/feature-isolated/footnotes-endnotes.docx"
HDRFTR = "generated/feature-isolated/header-footer-sections.docx"
TABLES = "generated/feature-isolated/table-merged-nested.docx"
FRAGMENTED = "generated/feature-isolated/fragmented-runs.docx"
GAUNTLET = "generated/gauntlet/gauntlet.docx"
MINIMAL = "generated/minimal-clean/minimal.docx"

GOLDEN_DIR = Path(__file__).parent / "golden"


def _doc(relpath: str):
    return docx.Document(str(fixture_path(relpath)))


def _ground_truth(relpath: str) -> dict:
    return json.loads(sidecar_path(fixture_path(relpath)).read_text(encoding="utf-8"))[
        "ground_truth"
    ]


def _blocks(relpath: str, view: str = "current") -> List[Block]:
    return list(iter_blocks(_doc(relpath), view=view))


class DescribeStoryParts:
    def it_lists_every_story_part_in_pinned_order(self):
        assert story_parts(_doc(GAUNTLET)) == (
            "word/document.xml",
            "word/header1.xml",
            "word/header2.xml",
            "word/header3.xml",
            "word/footer1.xml",
            "word/footnotes.xml",
            "word/endnotes.xml",
            "word/comments.xml",
        )

    def it_sees_footnotes_and_endnotes_stories(self):
        assert story_parts(_doc(NOTES)) == (
            "word/document.xml",
            "word/footnotes.xml",
            "word/endnotes.xml",
        )


class DescribeTrackedViews:
    def it_shows_the_accepted_text_in_the_current_view(self):
        truth = _ground_truth(TRACKED)["revised_paragraph"]
        block = _blocks(TRACKED, "current")[1]
        assert block.text == truth["visible_text_current"]
        assert block.in_insert and not block.in_delete

    def it_shows_the_pre_change_text_in_the_original_view(self):
        truth = _ground_truth(TRACKED)["revised_paragraph"]
        block = _blocks(TRACKED, "original")[1]
        assert block.text == truth["visible_text_original"]
        assert block.in_delete and not block.in_insert

    def it_shows_everything_in_the_all_view(self):
        block = _blocks(TRACKED, "all")[1]
        assert "forty-two" in block.text and "forty-seven" in block.text
        assert block.in_insert and block.in_delete

    def it_rejects_unknown_views(self):
        with pytest.raises(ValueError, match="view"):
            list(iter_blocks(_doc(MINIMAL), view="fancy"))


class DescribeTextBoxVisibility:
    @pytest.mark.parametrize("relpath", [TEXTBOX, TEXTBOX_LO])
    def it_emits_text_box_content_exactly_once(self, relpath: str):
        """The LibreOffice variant duplicates w:txbxContent in an
        mc:AlternateContent fallback; traversal must yield ONE block."""
        truth = _ground_truth(relpath)
        boxed = [b for b in _blocks(relpath) if b.in_text_box]
        assert [b.text for b in boxed] == truth["textbox_visible_texts"]

    @pytest.mark.parametrize("relpath", [TEXTBOX, TEXTBOX_LO])
    def it_keeps_text_box_text_out_of_the_host_paragraph(self, relpath: str):
        host_texts = [b.text for b in _blocks(relpath) if not b.in_text_box]
        assert not any("inside the text box" in text for text in host_texts)

    @pytest.mark.parametrize("relpath", [TEXTBOX, TEXTBOX_LO])
    def it_counts_one_traversed_text_box(self, relpath: str):
        counts = outline(_doc(relpath)).blind_region_counts
        assert counts["text_boxes"] == 1


class DescribeContentControls:
    def it_emits_block_control_content_flagged(self):
        truth = _ground_truth(CONTROLS)["block_control"]
        flagged = [b for b in _blocks(CONTROLS) if b.in_content_control]
        assert truth["text"] in [b.text for b in flagged]

    def it_keeps_inline_control_text_in_its_host_paragraph(self):
        truth = _ground_truth(CONTROLS)["inline_control"]
        texts = [b.text for b in _blocks(CONTROLS)]
        assert truth["host_paragraph_text"] in texts


class DescribeNoteStories:
    def it_reads_footnote_text_with_preserved_leading_space(self):
        truth = _ground_truth(NOTES)
        footnote_blocks = [b for b in _blocks(NOTES) if b.story == "word/footnotes.xml"]
        assert [b.text for b in footnote_blocks] == [truth["footnote"]["text"]]

    def it_reads_endnote_text_and_skips_separator_plumbing(self):
        truth = _ground_truth(NOTES)
        endnote_blocks = [b for b in _blocks(NOTES) if b.story == "word/endnotes.xml"]
        assert [b.text for b in endnote_blocks] == [truth["endnote"]["text"]]


class DescribeHeaderFooterStories:
    def it_reads_every_header_and_footer_text(self):
        truth = _ground_truth(HDRFTR)
        by_story = {}
        for block in _blocks(HDRFTR):
            by_story.setdefault(block.story, []).append(block.text)
        header_texts = sorted(
            text for story, texts in by_story.items()
            for text in texts
            if story.startswith("word/header")
        )
        assert header_texts == sorted(
            [
                truth["section1"]["header_text"],
                truth["section1"]["first_page_header_text"],
                truth["section2"]["header_text"],
            ]
        )
        footer_texts = sorted(
            text for story, texts in by_story.items()
            for text in texts
            if story.startswith("word/footer")
        )
        assert footer_texts == sorted(
            [truth["section1"]["footer_text"], truth["section2"]["footer_text"]]
        )


class DescribeTableBlocks:
    def it_reports_the_table_shape(self):
        (table_block,) = [b for b in _blocks(TABLES) if b.kind == "table"]
        assert table_block.table is not None
        assert table_block.table.rows == 3
        assert table_block.table.columns == 3
        assert table_block.table.has_merges
        assert table_block.table.has_nested_table

    def it_owns_all_text_inside_the_table_including_nested_cells(self):
        (table_block,) = [b for b in _blocks(TABLES) if b.kind == "table"]
        for expected in ("R0C0", "R1C1", "N00", "N11"):
            assert expected in table_block.text

    def it_does_not_emit_cell_paragraphs_as_separate_blocks(self):
        paragraph_texts = [b.text for b in _blocks(TABLES) if b.kind == "paragraph"]
        assert not any("N00" in text or "R0C0" in text for text in paragraph_texts)


class DescribeFragmentedRuns:
    def it_reassembles_fragmented_paragraph_text_exactly(self):
        truth = _ground_truth(FRAGMENTED)
        texts = [b.text for b in _blocks(FRAGMENTED)]
        assert truth["rate_paragraph"]["text"] in texts
        assert truth["nbsp_paragraph_text"] in texts


class DescribeGauntletCompleteness:
    """CONVENTIONS Phase 3 exit test: nothing visible is missed."""

    def it_sees_every_ground_truth_text_somewhere(self):
        truth = _ground_truth(GAUNTLET)
        union = "\n".join(b.text for b in _blocks(GAUNTLET, view="all"))
        expected_snippets = [
            truth["fragmented"]["rate_paragraph_text"],
            *truth["tracked"]["tracked_insertions"]["inserted_texts"],
            *truth["tracked"]["tracked_deletions"]["deleted_texts"],
            truth["content_controls"]["block_control"]["text"],
            truth["content_controls"]["inline_control"]["text"],
            *truth["textbox_visible_texts"],
            truth["notes"]["footnote"]["text"],
            truth["notes"]["endnote"]["text"],
            *truth["sections"]["header_texts"],
            *truth["sections"]["footer_texts"],
            truth["comments"][0]["text"],
            *truth["numbering"]["numbered_paragraph_texts"],
            "N00",  # nested table cell
        ]
        missing = [s for s in expected_snippets if s not in union]
        assert not missing, f"visible text missed by traversal: {missing!r}"


class DescribeAnchors:
    def it_numbers_blocks_contiguously_from_zero_per_story(self):
        blocks = _blocks(GAUNTLET)
        by_story = {}
        for block in blocks:
            by_story.setdefault(block.story, []).append(block.index)
        for story, indices in by_story.items():
            assert indices == list(range(len(indices))), story

    def it_hashes_normalized_block_text(self):
        block = _blocks(MINIMAL)[0]
        expected = hashlib.sha256(
            normalize_text(block.text).encode("utf-8")
        ).hexdigest()[:8]
        assert block.anchor.content_hash == expected
        assert block.anchor.story == block.story and block.anchor.index == block.index


class DescribeInspectionDeterminism:
    """CONVENTIONS §4 invariant: same input -> byte-identical JSON, run twice."""

    @pytest.mark.parametrize("relpath", [MINIMAL, GAUNTLET, TEXTBOX_LO])
    def it_produces_byte_identical_json_across_runs(self, relpath: str):
        first = json.dumps(outline(_doc(relpath)).to_dict(), ensure_ascii=False)
        second = json.dumps(outline(_doc(relpath)).to_dict(), ensure_ascii=False)
        assert first == second

    def it_matches_the_golden_outline_for_the_minimal_fixture(self):
        golden = (GOLDEN_DIR / "outline-minimal.json").read_text(encoding="utf-8")
        actual = json.dumps(outline(_doc(MINIMAL)).to_dict(), indent=2, ensure_ascii=False) + "\n"
        assert actual == golden, (
            "outline JSON shape drifted from the golden; if deliberate, update"
            " tests/paper/golden/outline-minimal.json in the same reviewed commit"
        )


class DescribeBlindRegionCounts:
    def it_counts_traversed_regions_on_the_gauntlet(self):
        counts = outline(_doc(GAUNTLET)).blind_region_counts
        assert counts == {
            "tracked_insertions": 2,
            "tracked_deletions": 1,
            "moves": 2,  # one moveFrom + one moveTo
            "format_changes": 2,  # rPrChange + pPrChange
            "content_controls": 3,  # block + inline + placeholder form control
            "text_boxes": 1,
            "fields": 2,  # fldSimple + one complex fldChar field
            "math": 0,
            "embedded_objects": 0,
            "alt_chunks": 0,
            "hidden_text": 0,
        }
