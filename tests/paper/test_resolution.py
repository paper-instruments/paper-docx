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

from .harness.paths import fixture_path

MULTIROUND = "generated/redline/multiround.docx"
MULTIROUND_ACCEPTED = "generated/redline/multiround-accepted.docx"
ROW_REVISIONS = "generated/feature-isolated/row-revisions.docx"
FORMAT_RICH = "generated/feature-isolated/format-changes-rich.docx"
PARAGRAPH_MERGE = "generated/feature-isolated/paragraph-merge.docx"
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
