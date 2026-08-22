"""Tests for docx.blocks — anchor-relative block operations."""

from __future__ import annotations

import copy
import datetime as dt
import shutil
from dataclasses import replace
from pathlib import Path
from typing import Callable, Union

import pytest

import docx
from docx.blocks import (
    BlockEditResult,
    BlockTarget,
    RichParagraph,
    TextRun,
    insert_blocks_after,
    insert_section_after,
    tracked_delete_paragraphs,
    tracked_replace_paragraphs,
)
from docx.document import Document
from docx.errors import (
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.protection import set_protection
from docx.search import Span, find_one
from docx.story import Block, iter_blocks, outline

from .harness.contract import assert_changed_parts, assert_refusal_atomic, save_and_reopen
from .harness.paths import fixture_path

MINIMAL = "generated/minimal-clean/minimal.docx"
TABLES = "generated/feature-isolated/table-merged-nested.docx"
CONTROLS = "generated/feature-isolated/content-control.docx"
FRAGMENTED = "generated/feature-isolated/fragmented-runs.docx"
GAUNTLET = "generated/gauntlet/gauntlet.docx"

FROZEN = dt.datetime(2026, 7, 7, 12, 0, 0, tzinfo=dt.timezone.utc)


def _doc(relpath: str):
    return docx.Document(str(fixture_path(relpath)))


def _texts(document, view: str = "current"):
    return [b.text for b in iter_blocks(document, view=view)]


def _memory_doc(*texts: str):
    document = docx.Document()
    for text in texts:
        document.add_paragraph(text)
    return document


def _cross_paragraph_target(document):
    return find_one(document, "alpha end\nbeta start")


class DescribeOneParagraphTargets:
    @pytest.mark.parametrize(
        ("operation", "as_string"),
        [
            (
                lambda document, target: insert_section_after(
                    document, target, heading="wrong", paragraphs=[]
                ),
                True,
            ),
            (
                lambda document, target: insert_blocks_after(
                    document,
                    target,
                    blocks=[RichParagraph(runs=[TextRun("wrong")])],
                ),
                False,
            ),
        ],
    )
    def it_refuses_string_derived_and_supplied_cross_paragraph_targets_atomically(
        self, operation, as_string: bool
    ):
        document = _memory_doc("alpha end", "beta start", "after")
        span = _cross_paragraph_target(document)
        target = span.text if as_string else span
        set_protection(document, edit="readOnly")

        error = assert_refusal_atomic(
            document,
            lambda doc: operation(doc, target),
            BoundaryViolationError,
        )

        assert "wholly within one paragraph" in str(error)
        assert "explicit live Block" in str(error)

    @pytest.mark.parametrize("crosses_paragraphs", [False, True])
    def it_checks_both_summary_evidence_and_concrete_atom_identity(
        self, crosses_paragraphs: bool
    ):
        document = _memory_doc("alpha end", "beta start")
        if crosses_paragraphs:
            span = find_one(document, "alpha end")
            span.crosses_paragraphs = True
        else:
            span = _cross_paragraph_target(document)
            span.crosses_paragraphs = False

        with pytest.raises(BoundaryViolationError, match="one paragraph"):
            insert_section_after(document, span, heading="wrong", paragraphs=[])

    @pytest.mark.parametrize("invalid_endpoint", ["start", "end"])
    @pytest.mark.parametrize("operation", ["delete", "replace"])
    def it_validates_both_range_endpoints_before_protection(
        self, invalid_endpoint: str, operation: str
    ):
        document = _memory_doc("before", "alpha end", "beta start", "after")
        cross = _cross_paragraph_target(document)
        start = cross if invalid_endpoint == "start" else "before"
        end = cross if invalid_endpoint == "end" else "after"
        set_protection(document, edit="readOnly")

        def mutate(doc):
            if operation == "delete":
                return tracked_delete_paragraphs(
                    doc, start, end_anchor=end, author="Reviewer"
                )
            return tracked_replace_paragraphs(
                doc, start, ["wrong"], end_anchor=end, author="Reviewer"
            )

        assert_refusal_atomic(document, mutate, BoundaryViolationError)

    def it_keeps_a_live_one_paragraph_span_as_a_valid_target(self):
        document = _memory_doc("before", "target text", "after")
        target = find_one(document, "target")

        insert_section_after(document, target, heading="inserted", paragraphs=[])

        assert _texts(document) == ["before", "target text", "inserted", "after"]

    def it_keeps_stale_span_taxonomy_before_protection(self):
        document = _memory_doc("target text")
        target = find_one(document, "target")
        document.paragraphs[0].text = "changed text"
        set_protection(document, edit="readOnly")

        assert_refusal_atomic(
            document,
            lambda doc: insert_section_after(
                doc, target, heading="wrong", paragraphs=[]
            ),
            TargetNotFoundError,
        )

    def it_keeps_missing_paragraph_taxonomy(self):
        document = _memory_doc("target text")
        target = find_one(document, "target")
        target._atoms[0].paragraph = None  # noqa: SLF001 - defensive same-package shape

        with pytest.raises(TargetNotFoundError, match="not inside a paragraph"):
            insert_section_after(document, target, heading="wrong", paragraphs=[])


class DescribeLiveBlockTargets:
    def it_follows_the_exact_element_across_an_index_shift(self):
        document = _memory_doc("A", "target", "C")
        target = tuple(iter_blocks(document))[1]
        document.paragraphs[0].insert_paragraph_before("new first")
        insert_section_after(document, target, heading="after target", paragraphs=[])
        assert _texts(document) == ["new first", "A", "target", "after target", "C"]

    @pytest.mark.parametrize("replacement", ["target", "TARGET", " target ", "t­arget"])
    def it_refuses_detached_or_replaced_elements_even_with_equivalent_text(
        self, replacement: str
    ):
        document = _memory_doc("A", "target", "C")
        target = tuple(iter_blocks(document))[1]
        old = target._element  # noqa: SLF001
        new = copy.deepcopy(old)
        for text in new.iter(qn("w:t")):
            text.text = replacement
        old.getparent().replace(old, new)
        before = document.element.xml
        with pytest.raises(TargetNotFoundError, match="stale"):
            insert_section_after(document, target, heading="wrong", paragraphs=[])
        assert document.element.xml == before

    def it_refuses_a_detached_element(self):
        document = _memory_doc("A", "target", "C")
        target = tuple(iter_blocks(document))[1]
        target._element.getparent().remove(target._element)  # noqa: SLF001
        with pytest.raises(TargetNotFoundError, match="stale"):
            insert_section_after(document, target, heading="wrong", paragraphs=[])

    def it_refuses_an_element_reparented_into_different_structure(self):
        document = _memory_doc("A", "target", "C")
        target = tuple(iter_blocks(document))[1]
        wrapper = OxmlElement("w:sdt")
        content = OxmlElement("w:sdtContent")
        wrapper.append(content)
        document.element.body.insert(1, wrapper)
        content.append(target._element)  # noqa: SLF001
        with pytest.raises(TargetNotFoundError, match="reparented"):
            insert_section_after(document, target, heading="wrong", paragraphs=[])

    def it_refuses_a_live_block_from_another_document(self):
        source = _memory_doc("foreign")
        destination = _memory_doc("local")
        foreign = next(iter(iter_blocks(source)))
        before = (source.element.xml, destination.element.xml)
        with pytest.raises(BoundaryViolationError, match="different document"):
            insert_section_after(destination, foreign, heading="wrong", paragraphs=[])
        assert (source.element.xml, destination.element.xml) == before

    def it_refuses_a_manually_detached_snapshot_instead_of_re_resolving_it(self):
        document = _memory_doc("target")
        snapshot = replace(
            next(iter(iter_blocks(document))),
            _document=None,
            _element=None,
            _story_root=None,
            _parent=None,
            _container_elements=(),
        )
        with pytest.raises(TargetNotFoundError, match="not a live block"):
            insert_section_after(document, snapshot, heading="wrong", paragraphs=[])

    @pytest.mark.parametrize(
        "changes",
        [
            {"_view": "bogus"},
            {"kind": "table"},
            {"story": "word/missing.xml"},
            {"_story_root": OxmlElement("w:document")},
            {"_container_elements": ()},
        ],
    )
    def it_refuses_live_blocks_with_contradictory_attachment_state(self, changes):
        document = _memory_doc("target")
        target = replace(next(iter(iter_blocks(document))), **changes)
        with pytest.raises(TargetNotFoundError, match="stale"):
            insert_section_after(document, target, heading="wrong", paragraphs=[])

    def it_refuses_a_live_table_block_as_a_paragraph_anchor(self):
        document = _memory_doc("paragraph")
        document.add_table(rows=1, cols=1).cell(0, 0).text = "cell"
        table = next(block for block in iter_blocks(document) if block.kind == "table")
        with pytest.raises(UnsupportedStructureError, match="table block"):
            insert_section_after(document, table, heading="wrong", paragraphs=[])


LiveTarget = Union[Block, Span]
BlockMutation = Callable[[Document, BlockTarget], BlockEditResult]


def _live_paragraph_target(
    document: Document, kind: str, view: str, text: str
) -> LiveTarget:
    if kind == "block":
        return next(
            block
            for block in iter_blocks(document, view=view)
            if block.kind == "paragraph" and block.text == text
        )
    return find_one(document, text, view=view)


def _insert_section(document: Document, target: BlockTarget) -> BlockEditResult:
    return insert_section_after(document, target, heading="new", paragraphs=[])


def _insert_blocks(document: Document, target: BlockTarget) -> BlockEditResult:
    return insert_blocks_after(
        document,
        target,
        blocks=[RichParagraph(runs=[TextRun("new")])],
    )


def _tracked_delete(document: Document, target: BlockTarget) -> BlockEditResult:
    return tracked_delete_paragraphs(document, target, author="A", date=FROZEN)


def _tracked_replace(document: Document, target: BlockTarget) -> BlockEditResult:
    return tracked_replace_paragraphs(
        document, target, ["new"], author="A", date=FROZEN
    )


_BLOCK_MUTATIONS: tuple[object, ...] = (
    pytest.param(_insert_section, id="insert-section"),
    pytest.param(_insert_blocks, id="insert-blocks"),
    pytest.param(_tracked_delete, id="tracked-delete"),
    pytest.param(_tracked_replace, id="tracked-replace"),
)


class DescribeMutationProjectionAuthority:
    @pytest.mark.parametrize("operation", _BLOCK_MUTATIONS)
    @pytest.mark.parametrize("target_kind", ["block", "span"])
    @pytest.mark.parametrize("view", ["original", "all"])
    def it_refuses_historical_live_mutation_anchors_atomically(
        self, operation: BlockMutation, target_kind: str, view: str
    ):
        document = _memory_doc("before", "target", "after")
        target = _live_paragraph_target(document, target_kind, view, "target")

        refusal = assert_refusal_atomic(
            document,
            lambda doc: operation(doc, target),
            UnsupportedStructureError,
        )

        assert f"view='{view}'" in str(refusal)
        assert 'view="current"' in str(refusal)
        if isinstance(target, Span):
            target._validate_fresh()  # noqa: SLF001  # pyright: ignore[reportPrivateUsage]
        else:
            assert target.text == "target"

    @pytest.mark.parametrize("historical_endpoint", ["start", "end"])
    def it_checks_both_range_projections_before_document_protection(
        self, historical_endpoint: str
    ):
        document = _memory_doc("start", "end")
        current_start = find_one(document, "start")
        current_end = find_one(document, "end")
        historical_start = find_one(document, "start", view="original")
        historical_end = find_one(document, "end", view="original")
        start = historical_start if historical_endpoint == "start" else current_start
        end = historical_end if historical_endpoint == "end" else current_end
        set_protection(document, edit="readOnly")

        refusal = assert_refusal_atomic(
            document,
            lambda doc: tracked_delete_paragraphs(
                doc, start, end_anchor=end, author="A", date=FROZEN
            ),
            UnsupportedStructureError,
        )

        assert "view='original'" in str(refusal)

    @pytest.mark.parametrize("operation", _BLOCK_MUTATIONS)
    @pytest.mark.parametrize("target_kind", ["block", "span"])
    def it_keeps_current_live_targets_authoritative(
        self, operation: BlockMutation, target_kind: str
    ):
        document = _memory_doc("before", "target", "after")
        target = _live_paragraph_target(document, target_kind, "current", "target")

        result = operation(document, target)

        assert result.inserted_blocks or result.deleted_blocks

    def it_does_not_resolve_a_historical_only_exact_string_destination(self):
        document = docx.Document()
        paragraph = document.add_paragraph()._p  # pyright: ignore[reportPrivateUsage]
        deletion = OxmlElement("w:del")
        deletion.set(qn("w:id"), "1")
        run = OxmlElement("w:r")
        deleted_text = OxmlElement("w:delText")
        deleted_text.text = "historical only"
        run.append(deleted_text)
        deletion.append(run)
        paragraph.append(deletion)
        assert find_one(document, "historical only", view="original")

        assert_refusal_atomic(
            document,
            lambda doc: insert_section_after(
                doc, "historical only", heading="new", paragraphs=[]
            ),
            TargetNotFoundError,
        )

    @pytest.mark.parametrize(
        "operation",
        [tracked_delete_paragraphs, tracked_replace_paragraphs],
    )
    @pytest.mark.parametrize("historical_endpoint", ["start", "end"])
    @pytest.mark.parametrize("target_kind", ["block", "span"])
    @pytest.mark.parametrize("view", ["original", "all"])
    def it_validates_each_tracked_range_endpoint_before_mutation(
        self,
        operation: Callable[..., BlockEditResult],
        historical_endpoint: str,
        target_kind: str,
        view: str,
    ):
        document = _memory_doc("start", "end", "after")
        current_start = _live_paragraph_target(document, target_kind, "current", "start")
        current_end = _live_paragraph_target(document, target_kind, "current", "end")
        historical_start = _live_paragraph_target(document, target_kind, view, "start")
        historical_end = _live_paragraph_target(document, target_kind, view, "end")
        start = historical_start if historical_endpoint == "start" else current_start
        end = historical_end if historical_endpoint == "end" else current_end

        def mutate(doc: Document) -> BlockEditResult:
            if operation is tracked_replace_paragraphs:
                return operation(
                    doc,
                    start,
                    ["replacement"],
                    end_anchor=end,
                    author="A",
                    date=FROZEN,
                )
            return operation(doc, start, end_anchor=end, author="A", date=FROZEN)

        refusal = assert_refusal_atomic(
            document, mutate, UnsupportedStructureError
        )

        assert f"view='{view}'" in str(refusal)
        assert 'view="current"' in str(refusal)


class DescribeLegacyAnchorRefusal:
    @pytest.mark.parametrize(
        "operation",
        [
            lambda document, anchor: insert_section_after(
                document, anchor, heading="wrong", paragraphs=[]
            ),
            lambda document, anchor: tracked_delete_paragraphs(
                document, anchor, author="A"
            ),
            lambda document, anchor: tracked_replace_paragraphs(
                document, anchor, ["wrong"], author="A"
            ),
            lambda document, anchor: insert_blocks_after(
                document, anchor, blocks=[RichParagraph(runs=[TextRun("valid")])]
            ),
        ],
    )
    def it_refuses_generic_anchor_evidence_for_every_mutation_family(self, operation):
        document = _memory_doc("target")
        anchor = next(iter(iter_blocks(document))).anchor
        before = document.element.xml
        with pytest.raises(UnsupportedStructureError, match="inert location evidence"):
            operation(document, anchor)
        assert document.element.xml == before


class DescribeInsertSectionAfter:
    def it_inserts_a_heading_and_body_after_a_string_anchor(self, tmp_path: Path):
        document = _doc(MINIMAL)
        result = insert_section_after(
            document,
            "First body paragraph",
            heading="New Section",
            paragraphs=["Alpha body.", "Beta body."],
        )
        assert result.inserted_blocks == 3
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        blocks = list(iter_blocks(reopened))
        assert [b.text for b in blocks[1:5]] == [
            "First body paragraph with perfectly ordinary text.",
            "New Section",
            "Alpha body.",
            "Beta body.",
        ]
        assert blocks[2].style_id == "Heading2"

    def it_accepts_a_live_block_object(self):
        document = _doc(MINIMAL)
        anchor = outline(document).blocks[1]
        insert_section_after(
            document, anchor, heading="Anchored Section", paragraphs=["Body."]
        )
        assert "Anchored Section" in _texts(document)

    def it_validates_style_ids_before_mutating(self):
        document = _doc(MINIMAL)
        assert_refusal_atomic(
            document,
            lambda doc: insert_section_after(
                doc, "First body paragraph", heading="X", paragraphs=["Y"],
                heading_style="NoSuchStyle99",
            ),
            TargetNotFoundError,
        )

    def it_requires_an_author_when_tracked(self):
        document = _doc(MINIMAL)
        with pytest.raises(ValueError, match="author"):
            insert_section_after(
                document, "First body paragraph", heading="X", paragraphs=[],
                tracked=True,
            )

    def it_makes_tracked_insertions_reject_away_completely(self):
        """The paragraph-mark stamp means rejection leaves no husk behind."""
        document = _doc(MINIMAL)
        before = _texts(document)
        insert_section_after(
            document,
            "First body paragraph",
            heading="Tracked Section",
            paragraphs=["Tracked body."],
            tracked=True,
            author="Carol QA",
            date=FROZEN,
        )
        assert "Tracked Section" in _texts(document)
        document.revisions.reject_all()
        assert _texts(document) == before

    def it_keeps_the_changed_part_budget(self, tmp_path: Path):
        source = fixture_path(MINIMAL)
        working = tmp_path / "work.docx"
        shutil.copyfile(source, working)
        document = docx.Document(str(working))
        insert_section_after(
            document, "First body paragraph", heading="Budget", paragraphs=["B."]
        )
        out = tmp_path / "out.docx"
        docx.package.patch_save(working, document, out)
        assert_changed_parts(working, out, {"word/document.xml"})


class DescribeTrackedDeleteParagraphs:
    def it_preserves_each_runs_formatting_inside_the_deletion(self):
        """Runs move into w:del with their own rPr (reject restores exactly)."""
        document = _doc(FRAGMENTED)
        result = tracked_delete_paragraphs(
            document, "Consulting rate", count=1, author="Carol QA", date=FROZEN
        )
        assert result.deleted_blocks == 1
        assert result.deleted_text[0].startswith("Consulting rate: $75–100/hr")
        (deletion,) = document.element.body.xpath("//w:p/w:del")
        runs = deletion.findall(qn("w:r"))
        assert len(runs) == 8, "each original run must survive inside w:del"
        assert deletion.xpath(".//w:delText")
        assert not deletion.xpath(".//w:t")
        bold_runs = [r for r in runs if r.find(qn("w:rPr")) is not None
                     and r.find(qn("w:rPr")).find(qn("w:b")) is not None]
        assert len(bold_runs) == 3, "bold formatting must survive deletion markup"

    def it_stamps_the_paragraph_mark_so_accept_removes_the_paragraph(self):
        document = _doc(MINIMAL)
        tracked_delete_paragraphs(
            document, "First body paragraph", count=1, author="Carol QA", date=FROZEN
        )
        document.revisions.accept_all()
        texts = _texts(document)
        assert "First body paragraph with perfectly ordinary text." not in texts
        assert "" not in texts, "accepting a deletion must not leave an empty husk"

    def it_selects_a_range_by_end_anchor(self):
        document = _doc(MINIMAL)
        result = tracked_delete_paragraphs(
            document,
            "First body paragraph",
            end_anchor="Second body paragraph",
            author="Carol QA",
            date=FROZEN,
        )
        assert result.deleted_blocks == 2

    def it_refuses_ranges_that_do_not_share_one_parent(self):
        """Body paragraph -> table-cell paragraph: never silently corrupted."""
        document = _doc(TABLES)
        assert_refusal_atomic(
            document,
            lambda doc: tracked_delete_paragraphs(
                doc, "Paragraph before the merged",
                end_anchor="N11",  # nested-table cell paragraph
                author="Carol QA", date=FROZEN,
            ),
            BoundaryViolationError,
        )

    def it_counts_ranges_among_siblings_so_tables_are_bracketed_not_selected(self):
        """count=2 from the paragraph before a table selects the paragraphs
        AROUND it (siblings); the table itself is untouched."""
        document = _doc(TABLES)
        result = tracked_delete_paragraphs(
            document, "Paragraph before the merged", count=2,
            author="Carol QA", date=FROZEN,
        )
        assert result.deleted_blocks == 2
        assert result.deleted_text[1] == "Paragraph after the tables."
        assert len(document.tables) == 1  # table survives

    def it_refuses_paragraphs_with_inline_controls(self):
        document = _doc(CONTROLS)
        assert_refusal_atomic(
            document,
            lambda doc: tracked_delete_paragraphs(
                doc, "Inline control follows", count=1,
                author="Carol QA", date=FROZEN,
            ),
            UnsupportedStructureError,
        )

    def it_refuses_cross_story_ranges(self):
        document = _doc(GAUNTLET)
        with pytest.raises(BoundaryViolationError, match="different story"):
            tracked_delete_paragraphs(
                document,
                "Gauntlet section two body",
                end_anchor="Gauntlet header, section one",
                author="Carol QA",
                date=FROZEN,
            )

    def it_refuses_counts_past_the_last_paragraph(self):
        document = _doc(MINIMAL)
        with pytest.raises(TargetNotFoundError, match="past the last"):
            tracked_delete_paragraphs(
                document, "Second body paragraph", count=10,
                author="Carol QA", date=FROZEN,
            )


class DescribeTrackedReplaceParagraphs:
    def it_deletes_the_range_and_inserts_replacements(self, tmp_path: Path):
        document = _doc(MINIMAL)
        result = tracked_replace_paragraphs(
            document,
            "Second body paragraph",
            ["Replacement one.", "Replacement two."],
            author="Carol QA",
            date=FROZEN,
        )
        assert result.deleted_blocks == 1
        assert result.inserted_blocks == 2
        assert len(result.revision_ids) == 3
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        texts = _texts(reopened)
        assert "Replacement one." in texts
        assert "Replacement two." in texts

    def it_keeps_the_changed_part_budget(self, tmp_path: Path):
        source = fixture_path(MINIMAL)
        working = tmp_path / "work.docx"
        shutil.copyfile(source, working)
        document = docx.Document(str(working))
        tracked_replace_paragraphs(
            document, "Second body paragraph", ["R."], author="Carol QA", date=FROZEN
        )
        out = tmp_path / "out.docx"
        docx.package.patch_save(working, document, out)
        assert_changed_parts(working, out, {"word/document.xml"})

    @pytest.mark.lo_smoke
    def it_produces_output_libreoffice_can_open(self, tmp_path: Path):
        from .harness.lo import assert_libreoffice_opens

        document = _doc(MINIMAL)
        tracked_replace_paragraphs(
            document, "Second body paragraph", ["LO check."],
            author="Carol QA", date=FROZEN,
        )
        out = tmp_path / "out.docx"
        document.save(str(out))
        assert_libreoffice_opens(out)
