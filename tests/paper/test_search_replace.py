"""Tests for docx.search: find_text/Span, run-preserving replace, and
tracked replace over the revision vocabulary.

Span mapping is tested BY USE (perform a replace, assert the outcome) —
never by inspecting private offsets.
"""

from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path
from typing import Any, cast

import pytest
from lxml import etree

import docx
import docx._clock
from docx.enum.style import WD_STYLE_TYPE
from docx.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.search import find_one, find_text, normalize_text
from docx.story import iter_blocks

from .harness.contract import assert_changed_parts, assert_refusal_atomic, save_and_reopen
from .harness.paths import fixture_path

FRAGMENTED = "generated/feature-isolated/fragmented-runs.docx"
TRACKED = "generated/feature-isolated/tracked-ins-del.docx"
CONTROLS = "generated/feature-isolated/content-control.docx"
GAUNTLET = "generated/gauntlet/gauntlet.docx"
MINIMAL = "generated/minimal-clean/minimal.docx"

RATE_TEXT = "$75–100/hr on a “full-service” basis"
FROZEN = dt.datetime(2026, 7, 7, 12, 0, 0, tzinfo=dt.timezone.utc)
W = nsdecls("w")


def _doc(relpath: str):
    return docx.Document(str(fixture_path(relpath)))


class DescribeNormalizeText:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("‘quoted’", "'quoted'"),
            ("“quoted”", '"quoted"'),
            ("75–100", "75-100"),
            ("em—dash", "em-dash"),
            ("minus−sign", "minus-sign"),
            ("no break", "no break"),  # NBSP
            ("figure space", "figure space"),
            ("narrow break", "narrow break"),
            ("thin space", "thin space"),
            ("soft­hyphen", "softhyphen"),
            ("tab\there", "tab here"),
            ("cr\rhere", "cr here"),
            ("many   spaces\n\n here", "many spaces here"),
            ("CaseFOLD", "casefold"),
        ],
    )
    def it_applies_each_pinned_normalization_rule(self, raw: str, expected: str):
        assert normalize_text(raw) == expected


class DescribeFindText:
    def it_matches_ascii_needles_against_smart_characters_across_runs(self):
        span = find_one(
            _doc(FRAGMENTED),
            '$75-100/hr on a "full-service" basis',
            match="normalized",
        )
        assert span.text == RATE_TEXT  # raw text captured verbatim, 8 runs deep
        assert span.match_policy == "normalized"
        assert not span.crosses_paragraphs

    def it_matches_through_no_break_spaces(self):
        span = find_one(
            _doc(FRAGMENTED), "Net 30 payment terms", match="normalized"
        )
        assert "Net 30" in span.text  # captured text preserves the raw NBSP

    def it_matches_across_a_paragraph_boundary(self):
        document = _doc(MINIMAL)
        needle = "ordinary text.\nSecond body paragraph"
        spans = find_text(document, needle)
        assert len(spans) == 1
        assert spans[0].text == needle
        assert spans[0].crosses_paragraphs
        assert find_one(document, needle).crosses_paragraphs

    @pytest.mark.parametrize(
        ("document_text", "needle"),
        [
            ("Company", "company"),
            ("“quoted”", '"quoted"'),
            ("Net Income – Adjusted", "Net Income - Adjusted"),
            ("Net 30", "Net 30"),
            ("alpha   beta", "alpha beta"),
            ("soft­hyphen", "softhyphen"),
        ],
    )
    def it_keeps_literal_distinctions_by_default(
        self, document_text: str, needle: str
    ):
        document = docx.Document()
        document.add_paragraph(document_text)
        assert find_text(document, needle) == []
        assert find_one(document, needle, match="normalized").text == document_text

    def it_matches_exact_text_through_run_fragmentation(self):
        span = find_one(_doc(FRAGMENTED), RATE_TEXT)
        assert span.text == RATE_TEXT
        assert span.match_policy == "exact"

    def it_treats_exact_whitespace_literally(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("alpha ")
        paragraph.add_run(" beta")
        assert find_one(document, "  ").text == "  "
        assert find_text(document, " ")
        assert find_text(document, "   ") == []
        assert find_text(document, " ", match="normalized") == []

    def it_uses_only_newline_as_the_exact_paragraph_separator(self):
        document = _doc(MINIMAL)
        assert find_text(document, "ordinary text.\nSecond body paragraph")
        assert find_text(document, "ordinary text. Second body paragraph") == []
        assert find_text(document, "ordinary text.\rSecond body paragraph") == []
        assert find_text(document, "\n") == []
        assert find_text(document, "\nSecond body paragraph") == []
        assert find_text(document, "") == []
        assert find_text(document, "\N{SOFT HYPHEN}", match="normalized") == []
        assert find_text(
            document,
            "ordinary text. Second body paragraph",
            match="normalized",
        )

        line_break_document = docx.Document()
        paragraph = line_break_document.add_paragraph()
        paragraph.add_run("alpha").add_break()
        paragraph.add_run("beta")
        line_break = find_one(line_break_document, "\n")
        assert line_break.text == "\n"
        assert not line_break.crosses_paragraphs

    @pytest.mark.parametrize("api", [find_text, find_one])
    def it_rejects_an_invalid_policy_before_scanning(self, api):
        with pytest.raises(ValueError, match="match must be one of"):
            api(object(), "target", match="fuzzy")

    def it_returns_matches_in_document_order_with_nth_selection(self):
        document = _doc(TRACKED)
        matches = find_text(document, "Paragraph")
        assert len(matches) == 2
        assert find_text(document, "Paragraph", nth=2)[0].text == matches[1].text
        assert find_text(document, "Paragraph", nth=3) == []

    def it_scopes_to_a_story_part(self):
        document = _doc(GAUNTLET)
        everywhere = find_text(document, "Gauntlet header, section one")
        assert {span.story for span in everywhere} == {"word/header1.xml"}
        assert find_text(document, "Gauntlet header, section one",
                         story="word/document.xml") == []

    def it_ranks_by_proximity_to_the_near_text(self):
        document = _doc(GAUNTLET)
        matches = find_text(
            document, "Gauntlet numbered item", near="item two"
        )
        assert [span.anchor.index for span in matches] == [27, 26]

    def it_applies_the_same_policy_to_target_and_near(self):
        document = docx.Document()
        document.add_paragraph("Near")
        document.add_paragraph("target")
        document.add_paragraph("near")
        document.add_paragraph("target")
        exact = find_text(document, "target", near="near")[0]
        normalized = find_text(
            document, "TARGET", near="NEAR", match="normalized"
        )[0]
        assert exact.anchor.index == 3
        assert normalized.anchor.index == 1

    def it_keeps_tied_near_candidates_in_stable_document_order(self):
        document = docx.Document()
        document.add_paragraph("target")
        document.add_paragraph("nearby")
        document.add_paragraph("target")
        document.add_paragraph("padding that makes the next candidate farther")
        document.add_paragraph("target")

        matches = find_text(document, "target", near="nearby")

        assert [span.anchor.index for span in matches] == [0, 2, 4]

    def it_keeps_the_complete_candidate_set_when_context_is_missing(self):
        document = docx.Document()
        document.add_paragraph("target")
        document.add_paragraph("target")

        ordinary = find_text(document, "target")
        ranked = find_text(document, "target", near="absent context")

        assert [span.anchor.index for span in ranked] == [
            span.anchor.index for span in ordinary
        ]

    def it_keeps_raw_order_separate_from_normalized_match_offsets(self):
        document = docx.Document()
        document.add_paragraph("Straße x Straße")
        spans = find_text(document, "strasse", match="normalized")
        assert [span.text for span in spans] == ["Straße", "Straße"]
        assert [span._raw_start for span in spans] == [0, 9]
        assert [span._match_start for span in spans] == [0, 10]

    def it_honors_the_view_parameter(self):
        document = _doc(TRACKED)
        assert find_text(document, "forty-two") == []  # deleted text, current view
        assert len(find_text(document, "forty-two", view="original")) == 1
        assert len(find_text(document, "forty-two", view="all")) == 1
        assert find_text(document, "forty-seven", view="original") == []

    def it_finds_text_inside_text_boxes_and_controls(self):
        document = _doc(GAUNTLET)
        boxed = find_one(document, "Text living inside the text box.")
        assert boxed.in_text_box
        controlled = find_one(document, "controlled text")
        assert controlled.in_content_control

    def it_uses_the_nearest_of_multiple_context_occurrences(self):
        document = docx.Document()
        document.add_paragraph("nearby")
        document.add_paragraph("target")
        document.add_paragraph("filler filler")
        document.add_paragraph("target")
        document.add_paragraph("padding")
        document.add_paragraph("nearby")

        matches = find_text(document, "target", near="nearby")

        assert [span.anchor.index for span in matches] == [1, 3]

    def it_ranks_eligible_stories_before_ineligible_ones_without_hiding_targets(
        self,
    ):
        document = docx.Document()
        document.add_paragraph("target")
        document.sections[0].header.paragraphs[0].text = "nearby target"

        matches = find_text(document, "target", near="nearby")

        assert [span.story for span in matches] == [
            "word/header1.xml",
            "word/document.xml",
        ]

    @pytest.mark.parametrize(
        ("wrapper", "excluded_view", "eligible_views"),
        [
            (
                f'<w:ins {W} w:id="4" w:author="Editor">'
                "<w:r><w:t>nearby</w:t></w:r></w:ins>",
                "original",
                ("current", "all"),
            ),
            (
                f'<w:del {W} w:id="5" w:author="Editor">'
                "<w:r><w:delText>nearby</w:delText></w:r></w:del>",
                "current",
                ("original", "all"),
            ),
        ],
    )
    def it_applies_view_scope_to_context(
        self, wrapper: str, excluded_view: str, eligible_views: tuple[str, str]
    ):
        document = docx.Document()
        document.add_paragraph("target")
        paragraph = document.add_paragraph("target ")
        paragraph._p.append(parse_xml(wrapper))

        excluded = find_text(document, "target", near="nearby", view=excluded_view)
        assert [span.anchor.index for span in excluded] == [0, 1]
        for view in eligible_views:
            ranked = find_text(document, "target", near="nearby", view=view)
            assert [span.anchor.index for span in ranked] == [1, 0]

    @pytest.mark.parametrize("match", ["exact", "normalized"])
    def it_rejects_near_with_nth_before_result_state(self, match: str):
        document = docx.Document()
        document.add_paragraph("target")
        document.add_paragraph("target")

        for needle in ("target", "missing target"):
            with pytest.raises(ValueError, match="mutually exclusive"):
                find_text(
                    document,
                    needle,
                    near="missing context",
                    nth=99,
                    match=match,
                )


class DescribeFindOne:
    def it_refuses_zero_matches(self):
        with pytest.raises(TargetNotFoundError, match="no match"):
            find_one(_doc(MINIMAL), "text that does not exist anywhere")

    def it_returns_one_match(self):
        span = find_one(_doc(MINIMAL), "Second body paragraph")
        assert span.text == "Second body paragraph"

    def it_refuses_ambiguity_without_disambiguators(self):
        with pytest.raises(AmbiguousTargetError, match="disambiguate"):
            find_one(_doc(TRACKED), "Paragraph")

    def it_resolves_ambiguity_with_nth(self):
        span = find_one(_doc(TRACKED), "Paragraph", nth=1)
        assert span.text == "Paragraph"

    def it_resolves_ambiguity_with_story(self):
        document = docx.Document()
        document.add_paragraph("target")
        document.sections[0].header.paragraphs[0].text = "target"

        span = find_one(document, "target", story="word/header1.xml")

        assert span.story == "word/header1.xml"

    def it_rejects_the_removed_near_keyword(self):
        with pytest.raises(TypeError, match="unexpected keyword argument 'near'"):
            find_one(
                _doc(MINIMAL),
                "target",
                near="context",  # type: ignore[call-arg]
            )


class DescribePlainReplace:
    def it_preserves_untouched_run_formatting(self, tmp_path: Path):
        document = _doc(FRAGMENTED)
        find_one(document, "$75–100/hr").replace("$85–110/hr")
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        paragraph = reopened.paragraphs[0]
        assert paragraph.text == (
            "Consulting rate: $85–110/hr on a “full-service” basis"
            " — travel time billed at $37.50/hr."
        )
        italic_runs = [r.text for r in paragraph.runs if r.italic]
        assert "".join(italic_runs) == "“full-service”"  # untouched formatting island
        bold_text = "".join(r.text for r in paragraph.runs if r.bold)
        assert bold_text.startswith("$85")

    def it_inherits_the_start_run_across_a_bold_to_italic_transition(self, tmp_path: Path):
        document = _doc(FRAGMENTED)
        span = find_one(document, "100/hr on a “full-")
        result = span.replace("90/hr on any “full-")

        assert not result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "start-run-inheritance.docx")
        assert any(
            run.bold and "90/hr on any " in run.text for run in reopened.paragraphs[0].runs
        )
        assert any(run.italic and "“full-" in run.text for run in reopened.paragraphs[0].runs)

    def it_restores_text_and_formatting_when_inverting_a_uniform_span(
        self, tmp_path: Path
    ):
        """Invariant: replace(x->y) then (y->x) restores text and formatting.

        Holds fully for spans of uniform formatting (here: the italic
        island, split across two runs)."""
        document = _doc(FRAGMENTED)
        find_one(document, "full-service").replace("bespoke")
        find_one(document, "bespoke").replace("full-service")
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        paragraph = reopened.paragraphs[0]
        assert paragraph.text == (
            "Consulting rate: $75–100/hr on a “full-service” basis"
            " — travel time billed at $37.50/hr."
        )
        italic_text = "".join(r.text for r in paragraph.runs if r.italic)
        assert italic_text == "“full-service”"
        bold_text = "".join(r.text for r in paragraph.runs if r.bold)
        assert bold_text == "$75–100/hr"

    def it_replaces_a_mixed_span_using_the_start_run_formatting(self):
        document = _doc(FRAGMENTED)
        span = find_one(document, RATE_TEXT)
        result = span.replace("something else entirely")

        assert not result.preserved_formatting_regions
        assert any(
            run.bold and run.text == "something else entirely"
            for run in document.paragraphs[0].runs
        )

    def it_inherits_italic_when_the_changed_interval_starts_in_italics(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Al").italic = True
        paragraph.add_run("pha").bold = True

        result = find_one(document, "Alpha").replace("Omega")

        assert not result.preserved_formatting_regions
        assert paragraph.text == "Omega"
        assert any(run.italic and run.text == "Omega" for run in paragraph.runs)
        assert not any(run.bold and run.text for run in paragraph.runs)

    def it_preserves_an_unchanged_prefix_before_start_run_inheritance(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("PRE").underline = True
        paragraph.add_run("Al").bold = True
        paragraph.add_run("pha").italic = True

        result = find_one(document, "PREAlpha").replace("PREOmega")

        assert not result.preserved_formatting_regions
        assert paragraph.text == "PREOmega"
        assert any(run.underline and run.text == "PRE" for run in paragraph.runs)
        assert any(run.bold and run.text == "Omega" for run in paragraph.runs)
        assert not any(run.italic and run.text for run in paragraph.runs)

    def it_leaves_later_suffix_runs_untouched_when_inheriting_from_the_start_run(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Al").bold = True
        paragraph.add_run("pha").italic = True
        paragraph.add_run(" END").underline = True

        result = find_one(document, "Alpha END").replace("Omega END")

        assert not result.preserved_formatting_regions
        assert paragraph.text == "Omega END"
        assert any(run.bold and run.text == "Omega" for run in paragraph.runs)
        assert any(run.underline and run.text == " END" for run in paragraph.runs)

    def it_preserves_exact_affixes_in_their_own_formatting_regions(
        self, tmp_path: Path
    ):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("prefix ").bold = True
        paragraph.add_run("middle").italic = True
        paragraph.add_run(" suffix").underline = True

        result = find_one(document, "prefix middle suffix").replace(
            "prefix changed suffix"
        )

        assert result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "affixes.docx")
        runs = reopened.paragraphs[0].runs
        assert [(run.text, run.bold, run.italic, run.underline) for run in runs] == [
            ("prefix ", True, None, None),
            ("changed", None, True, None),
            (" suffix", None, None, True),
        ]

    @pytest.mark.parametrize("same_format", [False, True])
    def it_refuses_ambiguous_repeated_affixes_atomically(self, same_format: bool):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Term").bold = True
        second = paragraph.add_run("Term")
        if same_format:
            second.bold = True
        span = find_one(document, "TermTerm")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("Term"),
            UnsupportedStructureError,
        )

        assert "exact affix alignment is ambiguous" in str(refusal)
        assert "re-find" in str(refusal)
        assert "'payment'" in str(refusal)
        assert "'settlement'" in str(refusal)
        assert span.replace("TermTerm").preserved_formatting_regions

    def it_refuses_an_insertion_at_a_formatting_boundary_atomically(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("A").bold = True
        paragraph.add_run("B")
        span = find_one(document, "AB")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("AXB"),
            UnsupportedStructureError,
        )

        assert "insertion point has competing" in str(refusal)
        assert span.replace("AB").preserved_formatting_regions

    def it_refuses_an_insertion_at_an_unlisted_wrapper_boundary_atomically(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("A")
        paragraph._p.append(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:customXml {W}><w:r><w:t>B</w:t></w:r></w:customXml>')
        )
        span = find_one(document, "AB")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("AXB"),
            UnsupportedStructureError,
        )

        assert "insertion point has competing" in str(refusal)

    def it_allows_an_insertion_between_equivalent_destinations(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("A").bold = True
        paragraph.add_run("B").bold = True

        find_one(document, "AB").replace("AXB")

        assert paragraph.text == "AXB"
        assert "".join(run.text for run in paragraph.runs if run.bold) == "AXB"

    def it_refuses_an_insertion_across_a_proofing_marker_atomically(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("A")
        first.bold = True
        first._r.addnext(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:proofErr {W} w:type="spellStart"/>')
        )
        paragraph.add_run("B").bold = True
        span = find_one(document, "AB")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("AXB"),
            UnsupportedStructureError,
        )

        assert "positional marker" in str(refusal)

    def it_refuses_an_insertion_across_a_bookmark_boundary_atomically(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("A")
        first.bold = True
        first._r.addnext(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:bookmarkStart {W} w:id="42" w:name="target"/>')
        )
        last = paragraph.add_run("B")
        last.bold = True
        last._r.addnext(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:bookmarkEnd {W} w:id="42"/>')
        )
        span = find_one(document, "AB")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("AXB"),
            UnsupportedStructureError,
        )

        assert "bookmark 'target'" in str(refusal)

    def it_refuses_replacement_across_equivalent_but_distinct_inline_wrappers(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph._p.append(  # pyright: ignore[reportPrivateUsage]
            parse_xml(
                f'<w:smartTag {W} w:uri="urn:test" w:element="same">'
                "<w:r><w:t>A</w:t></w:r></w:smartTag>"
            )
        )
        paragraph._p.append(  # pyright: ignore[reportPrivateUsage]
            parse_xml(
                f'<w:smartTag {W} w:uri="urn:test" w:element="same">'
                "<w:r><w:t>B</w:t></w:r></w:smartTag>"
            )
        )

        span = find_one(document, "AB")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("XY"),
            UnsupportedStructureError,
        )

        assert "separate inline wrapper owners" in str(refusal)

    def it_replaces_equivalent_fragmented_runs_and_consumes_the_span(
        self, tmp_path: Path
    ):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("frag").bold = True
        paragraph.add_run("mented").bold = True
        span = find_one(document, "fragmented")

        first = span.replace("unified")
        with pytest.raises(TargetNotFoundError, match="consumed.*re-find"):
            span.replace("renewed")
        replacement_span = find_one(document, "unified")
        second = replacement_span.replace("renewed")

        assert first.preserved_formatting_regions
        assert second.preserved_formatting_regions
        with pytest.raises(TargetNotFoundError, match="consumed.*re-find"):
            replacement_span.replace("again")
        reopened = save_and_reopen(document, tmp_path / "fragmented.docx")
        assert reopened.paragraphs[0].text == "renewed"
        assert "".join(run.text for run in reopened.paragraphs[0].runs if run.bold) == "renewed"

    def it_keeps_a_plain_noop_reusable_and_consumes_complete_deletion(self):
        document = docx.Document()
        document.add_paragraph("target")
        span = find_one(document, "target")

        result = span.replace("target")
        assert result.preserved_formatting_regions
        span.replace("target")
        span.replace("")

        with pytest.raises(TargetNotFoundError, match="re-find"):
            span.replace("again")

    def it_updates_xml_space_for_an_ordinary_replacement(self, tmp_path: Path):
        document = docx.Document()
        document.add_paragraph("plain")

        result = find_one(document, "plain").replace(" edged ")

        assert result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "space.docx")
        element = reopened.paragraphs[0]._p.find(".//" + qn("w:t"))
        assert element is not None
        assert element.text == " edged "
        assert element.get(qn("xml:space")) == "preserve"

    def it_clears_placeholder_state_after_an_ordinary_replacement(self):
        document = docx.Document()
        body = cast(Any, document.element).body
        body.insert(
            0,
            parse_xml(
                f'<w:p {W}><w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr>'
                '<w:sdtContent><w:r><w:rPr>'
                '<w:rStyle w:val="PlaceholderText"/></w:rPr>'
                '<w:t>Click or tap here to enter text.</w:t>'
                '</w:r></w:sdtContent></w:sdt></w:p>'
            ),
        )

        find_one(document, "Click or tap here to enter text.").replace("Filled")

        assert not body.xpath("//w:sdtPr/w:showingPlcHdr")
        assert not body.xpath('//w:rStyle[@w:val="PlaceholderText"]')

    def it_reports_start_run_inheritance_for_lexically_different_run_properties(
        self,
    ):
        document = docx.Document()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("alpha")
        second = paragraph.add_run("beta")
        first.bold = True
        second.bold = True
        second._r.get_or_add_rPr().find(qn("w:b")).set(qn("w:val"), "1")
        span = find_one(document, "alphabeta")

        result = span.replace("changed")

        assert not result.preserved_formatting_regions
        assert paragraph.text == "changed"
        assert paragraph.runs[0].bold

    def it_accepts_single_node_complete_run_formatting(self, tmp_path: Path):
        document = docx.Document()
        paragraph = document.add_paragraph("target")
        rpr = paragraph.runs[0]._r.get_or_add_rPr()  # pyright: ignore[reportPrivateUsage]
        rpr.append(parse_xml(f'<w:shd {W} w:fill="FFFF00"/>'))

        result = find_one(document, "target").replace("changed")

        assert result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "single-node-shading.docx")
        assert reopened.paragraphs[0].text == "changed"
        assert reopened.paragraphs[0]._p.xpath(  # pyright: ignore[reportPrivateUsage]
            'w:r/w:rPr/w:shd[@w:fill="FFFF00"]'
        )

    def it_accepts_fragmented_identical_complete_run_properties(self, tmp_path: Path):
        document = docx.Document()
        paragraph = document.add_paragraph()
        for text in ("alpha", "beta"):
            run = paragraph.add_run(text)
            run._r.get_or_add_rPr().append(  # pyright: ignore[reportPrivateUsage]
                parse_xml(f'<w:shd {W} w:fill="FFFF00"/>')
            )

        result = find_one(document, "alphabeta").replace("changed")

        assert result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "fragmented-shading.docx")
        assert reopened.paragraphs[0].text == "changed"
        assert reopened.paragraphs[0]._p.xpath(  # pyright: ignore[reportPrivateUsage]
            'w:r/w:rPr/w:shd[@w:fill="FFFF00"]'
        )

    def it_inherits_the_starting_character_style(self):
        document = docx.Document()
        first_style = document.styles.add_style("Replacement First", WD_STYLE_TYPE.CHARACTER)
        first_style.font.bold = True
        second_style = document.styles.add_style("Replacement Second", WD_STYLE_TYPE.CHARACTER)
        second_style.font.italic = True
        paragraph = document.add_paragraph()
        paragraph.add_run("alpha", style=first_style)
        paragraph.add_run("beta", style=second_style)
        span = find_one(document, "alphabeta")

        result = span.replace("changed")

        assert not result.preserved_formatting_regions
        assert paragraph.text == "changed"
        assert paragraph.runs[0].style.name == "Replacement First"

    def it_refuses_a_positional_marker_inside_the_changed_interval(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("alpha")
        paragraph.add_run("beta")
        first._r.addnext(parse_xml(f'<w:proofErr {W} w:type="spellStart"/>'))
        span = find_one(document, "alphabeta")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("changed"),
            UnsupportedStructureError,
        )
        assert "positional marker" in str(refusal)

    def it_preserves_a_marker_wholly_inside_an_unchanged_affix(self, tmp_path: Path):
        document = docx.Document()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("prefix")
        marker = parse_xml(f'<w:proofErr {W} w:type="spellStart"/>')
        first._r.addnext(marker)
        paragraph.add_run(" target")

        find_one(document, "prefix target").replace("prefix changed")

        reopened = save_and_reopen(document, tmp_path / "marker-affix.docx")
        children = list(reopened.paragraphs[0]._p)
        assert [child.tag for child in children] == [
            qn("w:r"), qn("w:proofErr"), qn("w:r")
        ]

    def it_keeps_the_changed_part_budget_to_the_document_part(self, tmp_path: Path):
        source = fixture_path(FRAGMENTED)
        working = tmp_path / "work.docx"
        shutil.copyfile(source, working)
        document = docx.Document(str(working))
        find_one(document, "$75–100/hr").replace("$95–120/hr")
        out = tmp_path / "out.docx"
        docx.package.patch_save(working, document, out)
        assert_changed_parts(working, out, {"word/document.xml"})

    def it_keeps_a_header_edit_to_its_own_story_part(self, tmp_path: Path):
        source = fixture_path(GAUNTLET)
        working = tmp_path / "header-work.docx"
        shutil.copyfile(source, working)
        document = docx.Document(str(working))
        result = find_one(document, "Gauntlet header, section one").replace(
            "Reviewed header, section one"
        )
        out = tmp_path / "header-out.docx"
        docx.package.patch_save(working, document, out)
        assert result.story == "word/header1.xml"
        assert result.preserved_formatting_regions
        assert_changed_parts(working, out, {"word/header1.xml"})

    @pytest.mark.lo_smoke
    def it_writes_a_safe_fragmented_replacement_libreoffice_can_open(
        self, tmp_path: Path
    ):
        from .harness.lo import assert_libreoffice_opens

        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("frag").bold = True
        paragraph.add_run("mented").bold = True
        find_one(document, "fragmented").replace("unified")
        out = tmp_path / "safe-fragmented.docx"
        document.save(out)
        assert_libreoffice_opens(out)


class DescribePreservationPolicies:
    def _insertion_document(self, *, revision_id: str = "41"):
        document = docx.Document()
        insertion = parse_xml(
            f'<w:ins {W} w:id="{revision_id}" w:author="Alice"'
            ' w:date="2026-07-07T12:00:00Z">'
            "<w:r><w:t>pending wording</w:t></w:r></w:ins>"
        )
        document.add_paragraph()._p.append(insertion)
        return document, insertion

    def it_preserves_an_existing_insertion_and_its_projections(self, tmp_path: Path):
        document, insertion = self._insertion_document()
        attributes = dict(insertion.attrib)
        original = [b.text for b in iter_blocks(document, view="original")]
        span = find_one(document, "pending")
        result = span.replace(
            "revised", preserve_revision=True
        )
        assert result.preserved_revision_ids == (41,)
        assert result.preserved_formatting_regions
        assert not result.preserved_structure
        assert dict(insertion.attrib) == attributes
        assert [b.text for b in iter_blocks(document, view="original")] == original
        with pytest.raises(TargetNotFoundError, match="consumed.*re-find"):
            span.replace("again", preserve_revision=True)
        path = tmp_path / "preserved-insertion.docx"
        document.save(path)
        accepted = docx.Document(path)
        accepted.revisions.accept_all()
        assert "revised wording" in [b.text for b in iter_blocks(accepted)]
        rejected = docx.Document(path)
        rejected.revisions.reject_all()
        assert "revised wording" not in [b.text for b in iter_blocks(rejected)]

    def it_keeps_safe_insertion_refusal_as_the_default(self):
        document, _ = self._insertion_document()
        with pytest.raises(UnsupportedStructureError, match="pending tracked insertion"):
            find_one(document, "pending").replace("revised")

    def it_refuses_a_base_text_and_insertion_crossing(self):
        document, insertion = self._insertion_document()
        insertion.addprevious(parse_xml(f'<w:r {W}><w:t>base </w:t></w:r>'))
        with pytest.raises(UnsupportedStructureError, match="mixes base text"):
            find_one(document, "base pending").replace(
                "combined", preserve_revision=True
            )

    def it_refuses_a_revision_nested_inside_the_preserved_insertion(self):
        document = docx.Document()
        document.add_paragraph()._p.append(
            parse_xml(
                f'<w:ins {W} w:id="41" w:author="Alice"'
                ' w:date="2026-07-07T12:00:00Z">'
                '<w:r><w:t xml:space="preserve">pending </w:t></w:r>'
                '<w:del w:id="42" w:author="Bob" w:date="2026-07-07T12:00:00Z">'
                '<w:r><w:delText xml:space="preserve">dropped </w:delText>'
                "</w:r></w:del>"
                "<w:r><w:t>wording</w:t></w:r></w:ins>"
            )
        )
        # the deletion is hidden from the current view but sits BETWEEN the
        # matched atoms; an in-place edit would leave it stranded
        with pytest.raises(UnsupportedStructureError, match="nested or mixed"):
            find_one(document, "pending wording").replace(
                "revised", preserve_revision=True
            )

    def it_refuses_a_span_crossing_two_insertions(self):
        document, insertion = self._insertion_document()
        insertion.addnext(
            parse_xml(
                f'<w:ins {W} w:id="55" w:author="Bob"'
                ' w:date="2026-07-07T12:00:00Z">'
                "<w:r><w:t> and more</w:t></w:r></w:ins>"
            )
        )
        with pytest.raises(UnsupportedStructureError, match="crosses multiple"):
            find_one(document, "wording and more").replace(
                "revised", preserve_revision=True
            )

    def it_refuses_a_tracked_move_destination(self):
        document = docx.Document()
        document.add_paragraph()._p.append(
            parse_xml(
                f'<w:moveTo {W} w:id="9" w:author="Alice"'
                ' w:date="2026-07-07T12:00:00Z">'
                "<w:r><w:t>moved wording</w:t></w:r></w:moveTo>"
            )
        )
        with pytest.raises(UnsupportedStructureError, match="tracked moves"):
            find_one(document, "moved wording").replace(
                "revised", preserve_revision=True
            )

    def it_refuses_preservation_outside_the_current_view(self):
        document, _ = self._insertion_document()
        span = find_text(document, "pending wording", view="all")[0]
        with pytest.raises(UnsupportedStructureError, match="view='current'"):
            span.replace("revised", preserve_revision=True)

    def it_preserves_the_exact_text_element_graph_and_distribution(self, tmp_path: Path):
        document = docx.Document()
        paragraph = document.add_paragraph()
        for text in ("alpha", " pay", "ment", " terms"):
            paragraph.add_run(text)
        paragraph.runs[1].bold = True
        proofing_marker = parse_xml(f'<w:proofErr {W} w:type="spellStart"/>')
        paragraph.runs[1]._r.addnext(proofing_marker)
        elements = tuple(paragraph._p.iter(qn("w:t")))
        runs = tuple(paragraph._p.iter(qn("w:r")))
        graph = tuple((e, e.tag, tuple(e.attrib.items()), tuple(e)) for e in elements)
        run_properties = tuple(
            etree.tostring(run.find(qn("w:rPr")))
            if run.find(qn("w:rPr")) is not None else None
            for run in runs
        )
        result = find_one(document, "alpha payment terms").replace(
            "alpha settlement terms", preserve_structure=True
        )
        assert result.preserved_structure
        assert [e.text for e in elements] == ["alpha", " set", "tlem", "ent terms"]
        assert tuple((e, e.tag, tuple(e.attrib.items()), tuple(e)) for e in elements) == graph
        assert tuple(paragraph._p.iter(qn("w:r"))) == runs
        assert proofing_marker.getparent() is paragraph._p
        assert tuple(
            etree.tostring(run.find(qn("w:rPr")))
            if run.find(qn("w:rPr")) is not None else None
            for run in runs
        ) == run_properties
        reopened = save_and_reopen(document, tmp_path / "exact.docx")
        assert reopened.paragraphs[-1].text == "alpha settlement terms"

    def it_keeps_empty_nodes_and_consumes_a_mutated_exact_span(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        for text in ("ab", "cd", "ef"):
            paragraph.add_run(text)
        elements = tuple(paragraph._p.iter(qn("w:t")))
        span = find_one(document, "abcdef")
        span.replace("x", preserve_structure=True)
        assert [e.text for e in elements] == ["x", "", ""]
        with pytest.raises(TargetNotFoundError, match="consumed.*re-find"):
            span.replace("again")
        assert find_one(document, "x").text == "x"

    def it_refuses_exact_edge_whitespace_without_changing_xml_space(self):
        document = docx.Document()
        paragraph = document.add_paragraph("plain")
        element = paragraph._p.find(".//" + qn("w:t"))
        assert element is not None
        assert element.get(qn("xml:space")) is None
        before = etree.tostring(paragraph._p)
        with pytest.raises(UnsupportedStructureError, match="edge whitespace"):
            find_one(document, "plain").replace(" plain", preserve_structure=True)
        assert etree.tostring(paragraph._p) == before

    def it_refuses_an_exact_plan_that_would_hollow_a_bookmark(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("outside")
        start = parse_xml(f'<w:bookmarkStart {W} w:id="9" w:name="target"/>')
        first._r.addnext(start)
        inside = paragraph.add_run("inside")
        end = parse_xml(f'<w:bookmarkEnd {W} w:id="9"/>')
        inside._r.addnext(end)
        before = etree.tostring(paragraph._p)
        with pytest.raises(UnsupportedStructureError, match="hollow"):
            find_one(document, "outsideinside").replace(
                "x", preserve_structure=True
            )
        assert etree.tostring(paragraph._p) == before

    def it_refuses_hollowing_a_bookmark_whose_markers_span_paragraphs(self):
        document = docx.Document()
        first = document.add_paragraph()._p
        first.append(
            parse_xml(f'<w:bookmarkStart {W} w:id="10" w:name="target"/>')
        )
        document.add_paragraph("inside")
        last = document.add_paragraph()._p
        last.append(parse_xml(f'<w:bookmarkEnd {W} w:id="10"/>'))
        before = document.element.xml

        with pytest.raises(UnsupportedStructureError, match="hollow"):
            find_one(document, "inside").replace("", preserve_structure=True)

        assert document.element.xml == before

    def it_keeps_a_revision_preserving_noop_reusable_across_a_bookmark(self):
        document = docx.Document()
        insertion = parse_xml(
            f'<w:ins {W} w:id="42" w:author="Alice">'
            '<w:r><w:t>outside</w:t></w:r>'
            '<w:bookmarkStart w:id="11" w:name="target"/>'
            '<w:r><w:t>inside</w:t></w:r>'
            '<w:bookmarkEnd w:id="11"/>'
            "</w:ins>"
        )
        document.add_paragraph()._p.append(insertion)
        span = find_one(document, "outsideinside")
        before = document.element.xml

        result = span.replace("outsideinside", preserve_revision=True)

        assert result.preserved_revision_ids == (42,)
        assert result.preserved_formatting_regions
        assert document.element.xml == before
        span.replace("outsideinside", preserve_revision=True)

    @pytest.mark.parametrize("preserve_revision", [False, True])
    def it_keeps_a_base_text_noop_reusable_across_a_bookmark(
        self, preserve_revision: bool
    ):
        document = docx.Document()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("outside")
        first._r.addnext(
            parse_xml(f'<w:bookmarkStart {W} w:id="11" w:name="target"/>')
        )
        inside = paragraph.add_run("inside")
        inside._r.addnext(parse_xml(f'<w:bookmarkEnd {W} w:id="11"/>'))
        before = document.element.xml
        span = find_one(document, "outsideinside")

        result = span.replace(
            "outsideinside", preserve_revision=preserve_revision
        )

        assert result.preserved_revision_ids == ()
        assert result.preserved_formatting_regions
        assert document.element.xml == before
        span.replace("outsideinside", preserve_revision=preserve_revision)

    def it_leaves_an_exact_noop_reusable(self):
        document = docx.Document()
        document.add_paragraph("same")
        span = find_one(document, "same")
        before = document.element.xml
        assert span.replace("same", preserve_structure=True).preserved_structure
        assert not span.replace(
            "same", preserve_structure=True
        ).preserved_formatting_regions
        assert document.element.xml == before
        span.replace("same", preserve_structure=True)

    def it_combines_revision_and_structure_preservation(self):
        document, insertion = self._insertion_document()
        text_element = next(insertion.iter(qn("w:t")))
        result = find_one(document, "pending").replace(
            "current", preserve_revision=True, preserve_structure=True
        )
        assert result.preserved_structure
        assert result.preserved_revision_ids == (41,)
        assert text_element.getparent() is not None

    def it_keeps_an_exact_patch_save_to_the_changed_story(self, tmp_path: Path):
        source = fixture_path(FRAGMENTED)
        working = tmp_path / "work.docx"
        shutil.copyfile(source, working)
        document = docx.Document(str(working))
        find_one(document, "$75–100/hr").replace(
            "$85–110/hr", preserve_structure=True
        )
        out = tmp_path / "out.docx"
        docx.package.patch_save(working, document, out)
        assert_changed_parts(working, out, {"word/document.xml"})

    @pytest.mark.parametrize("revision_id", ["bad", ""])
    def it_refuses_unreportable_insertion_ids(self, revision_id: str):
        document, insertion = self._insertion_document(revision_id=revision_id)
        if not revision_id:
            del insertion.attrib[qn("w:id")]
        with pytest.raises(UnsupportedStructureError, match="w:id"):
            find_one(document, "pending").replace("current", preserve_revision=True)

    @pytest.mark.parametrize("policy", ["preserve_structure", "preserve_revision"])
    def it_refuses_tracked_preservation_combinations(self, policy: str):
        document = docx.Document()
        document.add_paragraph("target")
        with pytest.raises(ValueError, match=policy):
            find_one(document, "target").replace(
                "changed", tracked=True, author="Editor", **{policy: True}
            )


class DescribeReplaceRefusals:
    def it_narrows_a_cross_paragraph_match_to_a_same_paragraph_change(self):
        document = docx.Document()
        document.add_paragraph("alpha end")
        document.add_paragraph("beta start")
        span = find_one(document, "alpha end\nbeta start")

        span.replace("ALPHA end beta start")

        assert [paragraph.text for paragraph in document.paragraphs] == [
            "ALPHA end",
            "beta start",
        ]

    def it_refuses_spans_over_deleted_text(self):
        document = _doc(TRACKED)
        span = find_one(document, "forty-two", view="all")
        with pytest.raises(UnsupportedStructureError, match="tracked-deleted"):
            span.replace("anything")

    def it_refuses_cross_paragraph_spans(self):
        document = _doc(MINIMAL)
        span = find_one(document, "ordinary text.\nSecond body paragraph")
        with pytest.raises(BoundaryViolationError, match="paragraph boundary"):
            span.replace("anything")

    def it_refuses_spans_crossing_a_content_control_boundary(self):
        document = _doc(CONTROLS)
        span = find_one(document, "follows: controlled")
        with pytest.raises(BoundaryViolationError, match="content-control"):
            span.replace("anything")

    def it_refuses_stale_spans(self):
        document = _doc(MINIMAL)
        span = find_one(document, "perfectly ordinary text")
        find_one(document, "perfectly ordinary").replace("thoroughly mundane")
        with pytest.raises(TargetNotFoundError, match="stale"):
            span.replace("anything")

    def it_refuses_atomically(self):
        """A refused replace leaves no trace, in memory or on disk."""
        document = _doc(CONTROLS)
        span = find_one(document, "follows: controlled")
        assert_refusal_atomic(
            document,
            lambda doc: span.replace("anything"),
            BoundaryViolationError,
            on_disk=(fixture_path(CONTROLS),),
        )


class DescribeTrackedReplace:
    def it_rolls_back_a_late_direct_tracked_failure_and_keeps_the_span_reusable(
        self, monkeypatch: pytest.MonkeyPatch
    ):
        from docx.oxml.revision import CT_RunTrackChange

        document = docx.Document()
        paragraph = document.add_paragraph("Alpha")
        span = find_one(document, "Alpha")
        before = document.element.xml  # pyright: ignore[reportUnknownMemberType]

        def fail_revision_creation(*_args, **_kwargs):
            raise RuntimeError("injected revision creation failure")

        monkeypatch.setattr(
            CT_RunTrackChange,
            "new",
            staticmethod(fail_revision_creation),
        )
        with pytest.raises(RuntimeError, match="injected revision"):
            span.replace("Zulu", tracked=True, author="Carol QA", date=FROZEN)

        assert document.element.xml == before  # pyright: ignore[reportUnknownMemberType]
        assert paragraph.text == "Alpha"
        assert span.replace("Zulu").inserted_text == "Zulu"

    def it_refuses_repeated_affix_ambiguity_atomically(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Term")
        paragraph.add_run("Term")
        span = find_one(document, "TermTerm")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace(
                "Term", tracked=True, author="Carol QA", date=FROZEN
            ),
            UnsupportedStructureError,
        )

        assert "exact affix alignment is ambiguous" in str(refusal)
        assert "exact substring" in str(refusal)
        assert not paragraph._p.xpath(  # pyright: ignore[reportPrivateUsage]
            ".//w:ins | .//w:del"
        )
        assert span.replace("TermTerm!").tracked is False

    def it_tracks_mixed_formatting_with_the_start_run_properties(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Al").bold = True
        paragraph.add_run("pha").italic = True
        span = find_one(document, "Alpha")

        result = span.replace("Omega", tracked=True, author="Carol QA", date=FROZEN)

        assert result.deleted_text == "Alpha"
        assert result.inserted_text == "Omega"
        (inserted_rpr,) = paragraph._p.xpath("w:ins/w:r/w:rPr")  # pyright: ignore[reportPrivateUsage]
        assert inserted_rpr.find(qn("w:b")) is not None
        assert inserted_rpr.find(qn("w:i")) is None
        document.revisions.accept_all()
        assert paragraph.text == "Omega"

    def it_tracks_mixed_formatting_without_collapsing_prefix_or_suffix_runs(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("KEEP").underline = True
        paragraph.add_run(" PRE").underline = True
        paragraph.add_run("Al").bold = True
        paragraph.add_run("pha").italic = True
        paragraph.add_run(" END").font.small_caps = True

        result = find_one(document, "KEEP PREAlpha END").replace(
            "KEEP PREOmega END",
            tracked=True,
            author="Carol QA",
            date=FROZEN,
        )

        assert result.deleted_text == "Alpha"
        assert result.inserted_text == "Omega"
        assert paragraph.runs[0].text == "KEEP"
        assert paragraph.runs[0].underline
        assert paragraph.runs[1].text == " PRE"
        assert paragraph.runs[1].underline
        assert paragraph.runs[-1].text == " END"
        assert paragraph.runs[-1].font.small_caps
        document.revisions.accept_all()
        assert paragraph.text == "KEEP PREOmega END"
        assert any(run.bold and run.text == "Omega" for run in paragraph.runs)
        assert paragraph.runs[-1].font.small_caps

    def it_tracks_a_uniform_fragmented_change_and_round_trips(self, tmp_path: Path):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Al").bold = True
        paragraph.add_run("pha").bold = True

        result = find_one(document, "Alpha").replace(
            "Omega", tracked=True, author="Carol QA", date=FROZEN
        )

        assert result.deleted_text == "Alph"
        assert result.inserted_text == "Omeg"
        path = tmp_path / "uniform-fragmented.docx"
        reopened = save_and_reopen(document, path)
        assert [block.text for block in iter_blocks(reopened)] == ["Omega"]
        assert [block.text for block in iter_blocks(reopened, view="original")] == [
            "Alpha"
        ]
        (inserted_rpr,) = reopened.element.body.xpath(  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue, reportUnknownVariableType]
            "//w:ins/w:r/w:rPr"
        )
        assert inserted_rpr.find(qn("w:b")) is not None  # pyright: ignore[reportUnknownMemberType]

        accepted = docx.Document(str(path))
        accepted.revisions.accept_all()
        assert [block.text for block in iter_blocks(accepted)] == ["Omega"]
        assert accepted.paragraphs[0].runs[0].bold
        rejected = docx.Document(str(path))
        rejected.revisions.reject_all()
        assert [block.text for block in iter_blocks(rejected)] == ["Alpha"]
        assert all(run.bold for run in rejected.paragraphs[0].runs if run.text)

    def it_keeps_differently_formatted_exact_affixes_in_place(self, tmp_path: Path):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("pre-").bold = True
        paragraph._p.append(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:proofErr {W} w:type="spellStart"/>')
        )
        paragraph.add_run("old")
        paragraph.add_run("-post").italic = True

        result = find_one(document, "pre-old-post").replace(
            "pre-new-post", tracked=True, author="Carol QA", date=FROZEN
        )

        assert result.deleted_text == "old"
        assert result.inserted_text == "new"
        reopened = save_and_reopen(document, tmp_path / "tracked-affixes.docx")
        assert "".join(
            reopened.element.body.xpath("//w:del//w:delText/text()")  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue, reportUnknownArgumentType]
        ) == "old"
        assert "".join(
            reopened.element.body.xpath("//w:ins//w:t/text()")  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue, reportUnknownArgumentType]
        ) == "new"
        assert reopened.paragraphs[0].runs[0].text == "pre-"
        assert reopened.paragraphs[0].runs[0].bold
        assert reopened.paragraphs[0].runs[-1].text == "-post"
        assert reopened.paragraphs[0].runs[-1].italic

    def it_deletes_mixed_formatting_with_each_source_property(self, tmp_path: Path):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("xAl").bold = True
        paragraph.add_run("phay").italic = True

        result = find_one(document, "Alpha").replace(
            "", tracked=True, author="Carol QA", date=FROZEN
        )

        assert result.deleted_text == "Alpha"
        assert result.inserted_text == ""
        path = tmp_path / "mixed-delete.docx"
        reopened = save_and_reopen(document, path)
        assert not reopened.element.body.xpath("//w:ins")  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
        deleted_runs = reopened.element.body.xpath("//w:del/w:r")  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue, reportUnknownVariableType]
        assert "".join(deleted_runs[0].xpath(".//w:delText/text()")) == "Al"  # pyright: ignore[reportUnknownMemberType, reportUnknownArgumentType]
        assert deleted_runs[0].find("w:rPr/w:b", deleted_runs[0].nsmap) is not None  # pyright: ignore[reportUnknownMemberType]
        assert "".join(deleted_runs[1].xpath(".//w:delText/text()")) == "pha"  # pyright: ignore[reportUnknownMemberType, reportUnknownArgumentType]
        assert deleted_runs[1].find("w:rPr/w:i", deleted_runs[1].nsmap) is not None  # pyright: ignore[reportUnknownMemberType]
        assert [block.text for block in iter_blocks(reopened)] == ["xy"]
        assert [block.text for block in iter_blocks(reopened, view="original")] == [
            "xAlphay"
        ]

        accepted = docx.Document(str(path))
        accepted.revisions.accept_all()
        assert accepted.paragraphs[0].text == "xy"
        rejected = docx.Document(str(path))
        rejected.revisions.reject_all()
        assert rejected.paragraphs[0].text == "xAlphay"
        assert any(run.bold and "Al" in run.text for run in rejected.paragraphs[0].runs)
        assert any(run.italic and "pha" in run.text for run in rejected.paragraphs[0].runs)

    def it_deletes_mixed_formatting_inside_one_wrapper_without_moving_ownership(
        self, tmp_path: Path
    ):
        document = docx.Document()
        paragraph = parse_xml(
            f'<w:p {W}><w:customXml w:uri="urn:paper" w:element="owner">'
            "<w:r><w:rPr><w:b/></w:rPr><w:t>Al</w:t></w:r>"
            "<w:r><w:rPr><w:i/></w:rPr><w:t>pha</w:t></w:r>"
            "</w:customXml></w:p>"
        )
        document.element.body.insert(0, paragraph)  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]

        find_one(document, "Alpha").replace("", tracked=True, author="Carol QA", date=FROZEN)
        path = tmp_path / "same-wrapper-delete.docx"
        document.save(path)
        rejected = docx.Document(path)
        rejected.revisions.reject_all()

        (wrapper,) = rejected.element.body.xpath("//w:customXml")  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
        assert "".join(element.text or "" for element in wrapper.iter(qn("w:t"))) == "Alpha"
        assert not rejected.element.body.xpath("//w:customXml[not(@w:element='owner')]//w:t")  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]

    def it_refuses_a_tracked_deletion_across_separate_identical_wrappers(self):
        document = docx.Document()
        paragraph = parse_xml(
            f"<w:p {W}>"
            '<w:customXml w:uri="urn:paper" w:element="same">'
            "<w:r><w:rPr><w:b/></w:rPr><w:t>Al</w:t></w:r>"
            "</w:customXml>"
            '<w:customXml w:uri="urn:paper" w:element="same">'
            "<w:r><w:rPr><w:i/></w:rPr><w:t>pha</w:t></w:r>"
            "</w:customXml>"
            "</w:p>"
        )
        document.element.body.insert(0, paragraph)  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
        span = find_one(document, "Alpha")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace("", tracked=True, author="Carol QA", date=FROZEN),
            UnsupportedStructureError,
        )

        assert "separate inline wrapper owners" in str(refusal)
        assert span.replace("Alpha!").inserted_text == "Alpha!"

    def it_inserts_only_at_a_proved_tracked_destination(self):
        inside = docx.Document()
        inside_paragraph = inside.add_paragraph()
        inside_paragraph.add_run("AB").bold = True
        find_one(inside, "AB").replace(
            "AXB", tracked=True, author="Carol QA", date=FROZEN
        )
        inside.revisions.accept_all()
        assert inside_paragraph.text == "AXB"
        assert all(run.bold for run in inside_paragraph.runs if run.text)

        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("A").bold = True
        paragraph.add_run("B")
        span = find_one(document, "AB")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace(
                "AXB", tracked=True, author="Carol QA", date=FROZEN
            ),
            UnsupportedStructureError,
        )

        assert "insertion point has competing" in str(refusal)
        assert "explicitly" in str(refusal)
        assert span.replace("AB!").tracked is False

        uniform = docx.Document()
        uniform_paragraph = uniform.add_paragraph()
        uniform_paragraph.add_run("A").bold = True
        uniform_paragraph.add_run("B").bold = True
        find_one(uniform, "AB").replace(
            "AXB", tracked=True, author="Carol QA", date=FROZEN
        )
        uniform.revisions.accept_all()
        assert uniform_paragraph.text == "AXB"

    def it_keeps_multirun_affixes_untouched_around_a_tracked_insertion(
        self, tmp_path: Path
    ):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("A").bold = True
        paragraph.add_run("A")
        paragraph.add_run("B")

        result = find_one(document, "AAB").replace(
            "AAXB", tracked=True, author="Carol QA", date=FROZEN
        )

        assert result.deleted_text == ""
        assert result.inserted_text == "X"
        path = tmp_path / "tracked-insertion-affixes.docx"
        reopened = save_and_reopen(document, path)
        assert [block.text for block in iter_blocks(reopened)] == ["AAXB"]
        assert [block.text for block in iter_blocks(reopened, view="original")] == [
            "AAB"
        ]
        assert [(run.text, run.bold) for run in reopened.paragraphs[0].runs] == [
            ("A", True),
            ("A", None),
            ("B", None),
        ]

        accepted = docx.Document(str(path))
        accepted.revisions.accept_all()
        assert [(run.text, run.bold) for run in accepted.paragraphs[0].runs] == [
            ("A", True),
            ("A", None),
            ("X", None),
            ("B", None),
        ]
        rejected = docx.Document(str(path))
        rejected.revisions.reject_all()
        assert [(run.text, run.bold) for run in rejected.paragraphs[0].runs] == [
            ("A", True),
            ("A", None),
            ("B", None),
        ]

    @pytest.mark.parametrize(
        ("retained_xml", "refusal_type", "message"),
        [
            (
                '<w:fldSimple w:instr=" DATE "><w:r><w:t>B</w:t></w:r></w:fldSimple>',
                UnsupportedStructureError,
                "field result",
            ),
            (
                '<w:hyperlink w:anchor="target"><w:r><w:t>B</w:t></w:r></w:hyperlink>',
                BoundaryViolationError,
                "hyperlink boundary",
            ),
            (
                "<w:sdt><w:sdtPr><w:tag w:val=\"target\"/></w:sdtPr>"
                "<w:sdtContent><w:r><w:t>B</w:t></w:r></w:sdtContent></w:sdt>",
                BoundaryViolationError,
                "content-control boundary",
            ),
        ],
    )
    def it_refuses_narrowing_across_retained_mutation_scopes(
        self,
        retained_xml: str,
        refusal_type: type[BaseException],
        message: str,
    ):
        document = docx.Document()
        paragraph = parse_xml(
            f'<w:p {W}><w:r><w:t>A</w:t><w:tab/></w:r>{retained_xml}</w:p>'
        )
        document.element.body.insert(0, paragraph)  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
        span = find_one(document, "A B", match="normalized")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace(
                "X B", tracked=True, author="Carol QA", date=FROZEN
            ),
            refusal_type,
        )

        assert message in str(refusal)
        assert not paragraph.xpath(".//w:ins | .//w:del")

    def it_proves_complete_tracked_destination_evidence(self):
        compatible = docx.Document()
        compatible_paragraph = parse_xml(
            f"<w:p {W}>"
            '<w:customXml w:uri="urn:paper" w:element="same">'
            "<w:r><w:rPr><w:smallCaps/></w:rPr><w:t>A</w:t></w:r>"
            "<w:r><w:rPr><w:smallCaps/></w:rPr><w:t>B</w:t></w:r>"
            "</w:customXml>"
            "</w:p>"
        )
        compatible.element.body.insert(  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
            0, compatible_paragraph
        )
        find_one(compatible, "AB").replace("AXB", tracked=True, author="Carol QA", date=FROZEN)
        compatible.revisions.accept_all()
        assert [block.text for block in iter_blocks(compatible)] == ["AXB"]

        identical_separate_ancestry = docx.Document()
        identical_separate_ancestry.element.body.insert(  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
            0,
            parse_xml(
                f"<w:p {W}>"
                '<w:customXml w:uri="urn:paper" w:element="same">'
                "<w:r><w:rPr><w:smallCaps/></w:rPr><w:t>A</w:t></w:r>"
                "</w:customXml>"
                '<w:customXml w:uri="urn:paper" w:element="same">'
                "<w:r><w:rPr><w:smallCaps/></w:rPr><w:t>B</w:t></w:r>"
                "</w:customXml>"
                "</w:p>"
            ),
        )
        identical_span = find_one(identical_separate_ancestry, "AB")
        identical_refusal = assert_refusal_atomic(
            identical_separate_ancestry,
            lambda _document: identical_span.replace(
                "AXB", tracked=True, author="Carol QA", date=FROZEN
            ),
            UnsupportedStructureError,
        )
        assert "competing formatting or structural destinations" in str(identical_refusal)

        different_ancestry = docx.Document()
        different_ancestry.element.body.insert(  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
            0,
            parse_xml(
                f"<w:p {W}>"
                '<w:customXml w:uri="urn:paper" w:element="left">'
                "<w:r><w:t>A</w:t></w:r></w:customXml>"
                '<w:customXml w:uri="urn:paper" w:element="right">'
                "<w:r><w:t>B</w:t></w:r></w:customXml>"
                "</w:p>"
            ),
        )
        ancestry_span = find_one(different_ancestry, "AB")
        ancestry_refusal = assert_refusal_atomic(
            different_ancestry,
            lambda _document: ancestry_span.replace(
                "AXB", tracked=True, author="Carol QA", date=FROZEN
            ),
            UnsupportedStructureError,
        )
        assert "competing formatting or structural destinations" in str(ancestry_refusal)

        lexical_rpr = docx.Document()
        lexical_rpr.element.body.insert(  # pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue]
            0,
            parse_xml(
                f"<w:p {W}>"
                "<w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r>"
                '<w:r><w:rPr><w:b w:val="true"/></w:rPr><w:t>B</w:t></w:r>'
                "</w:p>"
            ),
        )
        rpr_span = find_one(lexical_rpr, "AB")
        rpr_refusal = assert_refusal_atomic(
            lexical_rpr,
            lambda _document: rpr_span.replace("AXB", tracked=True, author="Carol QA", date=FROZEN),
            UnsupportedStructureError,
        )
        assert "competing formatting or structural destinations" in str(rpr_refusal)

    def it_refuses_a_tracked_insertion_across_a_positional_marker(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("A")
        paragraph._p.append(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:proofErr {W} w:type="spellStart"/>')
        )
        paragraph.add_run("B")
        span = find_one(document, "AB")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace(
                "AXB", tracked=True, author="Carol QA", date=FROZEN
            ),
            UnsupportedStructureError,
        )

        assert "positional marker" in str(refusal)
        assert not paragraph._p.xpath(  # pyright: ignore[reportPrivateUsage]
            ".//w:ins | .//w:del"
        )

    def it_refuses_a_tracked_deletion_across_a_positional_marker(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Al")
        paragraph._p.append(  # pyright: ignore[reportPrivateUsage]
            parse_xml(f'<w:proofErr {W} w:type="spellStart"/>')
        )
        paragraph.add_run("pha")
        span = find_one(document, "Alpha")

        refusal = assert_refusal_atomic(
            document,
            lambda _document: span.replace(
                "", tracked=True, author="Carol QA", date=FROZEN
            ),
            UnsupportedStructureError,
        )

        assert "positional marker" in str(refusal)
        assert not paragraph._p.xpath(  # pyright: ignore[reportPrivateUsage]
            ".//w:ins | .//w:del"
        )

    def it_marks_only_the_minimal_changed_span(self, tmp_path: Path):
        """The redline marks `75-10 -> 85-11`, not the sentence (pinned)."""
        document = _doc(FRAGMENTED)
        result = find_one(document, "$75–100/hr").replace(
            "$85–110/hr", tracked=True, author="Carol QA", date=FROZEN
        )
        assert result.deleted_text == "75–10"
        assert result.inserted_text == "85–11"
        assert not result.preserved_formatting_regions
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        blocks = list(iter_blocks(reopened))
        assert "$85–110/hr" in blocks[0].text  # current view: change applied
        original = list(iter_blocks(reopened, view="original"))
        assert "$75–100/hr" in original[0].text  # original view: change absent

    def it_keeps_deleted_text_in_delText_never_live_wt(self, tmp_path: Path):
        document = _doc(FRAGMENTED)
        find_one(document, "$75–100/hr").replace(
            "$85–110/hr", tracked=True, author="Carol QA", date=FROZEN
        )
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        body = reopened.element.body
        assert body.xpath("//w:del//w:delText"), "deleted text must be in w:delText"
        assert not body.xpath("//w:del//w:t"), "live w:t inside w:del is corrupt"

    def it_allocates_unique_increasing_revision_ids(self):
        document = _doc(TRACKED)  # fixture already holds ids 11, 12, 21
        first = find_one(document, "Paragraph before").replace(
            "Clause before", tracked=True, author="Carol QA", date=FROZEN
        )
        second = find_one(document, "Paragraph after").replace(
            "Sentence after", tracked=True, author="Carol QA", date=FROZEN
        )
        all_ids = [int(v) for v in document.element.body.xpath(
            "//w:ins/@w:id | //w:del/@w:id"
        )]
        assert len(all_ids) == len(set(all_ids)), "revision ids must be unique"
        assert min(first.revision_ids) > 21
        assert min(second.revision_ids) > max(first.revision_ids)

    def it_stamps_dates_from_the_injectable_clock(
        self, monkeypatch: pytest.MonkeyPatch
    ):
        monkeypatch.setattr(docx._clock, "now", lambda: FROZEN)
        document = _doc(MINIMAL)
        find_one(document, "perfectly ordinary").replace(
            "thoroughly mundane", tracked=True, author="Carol QA"
        )
        (ins,) = document.element.body.xpath("//w:ins")
        assert ins.get(qn("w:date")) == "2026-07-07T12:00:00Z"

    def it_preserves_run_formatting_on_both_sides(self, tmp_path: Path):
        document = _doc(FRAGMENTED)
        find_one(document, "$75–100/hr").replace(
            "$85–110/hr", tracked=True, author="Carol QA", date=FROZEN
        )
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        body = reopened.element.body
        # deleted text keeps each SOURCE run's rPr (one w:r per source run)
        del_rprs = body.xpath("//w:del/w:r/w:rPr")
        assert del_rprs, "deletion lost its runs"
        assert all(r.find(qn("w:b")) is not None for r in del_rprs), (
            "deleted side lost bold"
        )
        (ins_rpr,) = body.xpath("//w:ins/w:r/w:rPr")
        assert ins_rpr.find(qn("w:b")) is not None, "inserted side lost bold"

    def it_requires_an_author(self):
        document = _doc(MINIMAL)
        span = find_one(document, "perfectly ordinary")
        with pytest.raises(ValueError, match="author"):
            span.replace("x", tracked=True)

    def it_refuses_a_replacement_equal_to_the_existing_text_without_consuming(self):
        document = _doc(MINIMAL)
        span = find_one(document, "perfectly ordinary")
        with pytest.raises(TargetNotFoundError, match="nothing to change"):
            span.replace("perfectly ordinary", tracked=True, author="Carol QA")
        result = span.replace("quite ordinary")
        assert result.inserted_text == "quite ordinary"

    def it_refuses_cross_paragraph_tracked_targets(self):
        document = _doc(MINIMAL)
        span = find_one(document, "ordinary text.\nSecond body paragraph")
        with pytest.raises(BoundaryViolationError):
            span.replace("anything", tracked=True, author="Carol QA")

    def it_keeps_the_package_clean_and_budgeted(self, tmp_path: Path):
        from .harness import checks

        source = fixture_path(FRAGMENTED)
        working = tmp_path / "work.docx"
        shutil.copyfile(source, working)
        document = docx.Document(str(working))
        find_one(document, "$75–100/hr").replace(
            "$85–110/hr", tracked=True, author="Carol QA", date=FROZEN
        )
        out = tmp_path / "out.docx"
        docx.package.patch_save(working, document, out)
        assert_changed_parts(working, out, {"word/document.xml"})
        checks.assert_package_facts_clean(out)

    @pytest.mark.lo_smoke
    def it_produces_output_libreoffice_can_open(self, tmp_path: Path):
        from .harness.lo import assert_libreoffice_opens

        document = _doc(FRAGMENTED)
        find_one(document, "$75–100/hr").replace(
            "$85–110/hr", tracked=True, author="Carol QA", date=FROZEN
        )
        out = tmp_path / "out.docx"
        document.save(str(out))
        assert_libreoffice_opens(out)
