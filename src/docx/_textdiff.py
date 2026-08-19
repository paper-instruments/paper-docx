"""Human-readable semantic text diff (paper-docx).

Promoted from the reference `diff_docx.py`: `diff_package` says WHICH parts
changed; this says WHAT changed, as a unified diff over story-labeled block
text — the verification lens a human (or an agent writing an email summary)
actually reads. `pending_changes` diffs one document's "original" view
against its "current" view: what would change if every revision were
accepted.

Public import path: `docx.package.text_diff` / `docx.package.pending_changes`.
"""

from __future__ import annotations

import difflib
from dataclasses import dataclass
from typing import TYPE_CHECKING, Dict, List, Tuple, Union

if TYPE_CHECKING:
    import os

_PathLike = Union[str, "os.PathLike[str]"]


@dataclass(frozen=True)
class StoryTextDiff:
    story: str
    diff_lines: Tuple[str, ...]  # unified diff, empty when the story is unchanged

    def to_dict(self) -> dict:
        return {"story": self.story, "diff_lines": list(self.diff_lines)}


@dataclass(frozen=True)
class TextDiff:
    """Per-story unified diffs of visible block text."""

    stories: Tuple[StoryTextDiff, ...]

    @property
    def changed_line_count(self) -> int:
        return sum(
            1
            for story in self.stories
            for line in story.diff_lines
            if line.startswith(("+", "-")) and not line.startswith(("+++", "---"))
        )

    @property
    def is_empty(self) -> bool:
        return all(not story.diff_lines for story in self.stories)

    def to_dict(self) -> dict:
        return {
            "schema": "paper_text_diff",
            "version": 1,
            "changed_line_count": self.changed_line_count,
            "stories": [story.to_dict() for story in self.stories],
        }


def _story_lines(document, view: str) -> "Dict[str, List[str]]":
    from docx.story import iter_blocks

    lines: "Dict[str, List[str]]" = {}
    for block in iter_blocks(document, view=view):
        # no block index in the label: one insertion would shift every
        # later index and report the whole document as changed
        lines.setdefault(block.story, []).append(f"[{block.kind}] {block.text}")
    return lines


def _diff(
    before: "Dict[str, List[str]]",
    after: "Dict[str, List[str]]",
    from_label: str,
    to_label: str,
) -> TextDiff:
    stories = []
    for story in sorted(set(before) | set(after)):
        diff_lines = tuple(
            difflib.unified_diff(
                before.get(story, []),
                after.get(story, []),
                fromfile=f"{from_label}:{story}",
                tofile=f"{to_label}:{story}",
                lineterm="",
            )
        )
        stories.append(StoryTextDiff(story=story, diff_lines=diff_lines))
    return TextDiff(stories=tuple(stories))


def text_diff(path_a: _PathLike, path_b: _PathLike, *, view: str = "current") -> TextDiff:
    """What changed, textually, between the documents at `path_a` and `path_b`."""
    import docx

    before = _story_lines(docx.Document(str(path_a)), view)
    after = _story_lines(docx.Document(str(path_b)), view)
    return _diff(before, after, str(path_a), str(path_b))


def pending_changes(path: _PathLike) -> TextDiff:
    """What the document's pending revisions would change if all were accepted.

    Diffs the "original" view against the "current" view, so an empty `TextDiff` means no
    visible TEXT change rather than a clean document: formatting-only, table-structure and
    same-place move revisions produce no diff lines. Ask `Document.revisions` whether markup
    remains.
    """
    import docx

    document = docx.Document(str(path))
    before = _story_lines(document, "original")
    after = _story_lines(document, "current")
    return _diff(before, after, "original", "current")
