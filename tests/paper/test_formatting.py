"""The effective-format resolver.

Read-only, provenance-bearing: every value names the layer it came from,
toggles XOR through style layers (the nested-bold gotcha), and what cannot be
resolved is declared, never guessed.
"""

from __future__ import annotations

import pytest

import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.formatting import format_of, surrounding_format
from docx.search import find_one
from docx.shared import Pt
from docx.story import BlockLocator, iter_blocks

from .harness.paths import fixture_path

MINIMAL = "generated/minimal-clean/minimal.docx"


def _doc():
    return docx.Document(str(fixture_path(MINIMAL)))


class DescribeDirectFormatting:
    def it_resolves_direct_run_formatting_with_provenance(self):
        document = _doc()
        run = document.add_paragraph().add_run("bolded")
        run.bold = True
        run.font.size = Pt(14)
        resolved = format_of(run)
        assert resolved["bold"].value is True
        assert resolved["bold"].source == "direct"
        assert resolved["size_pt"].value == 14
        assert resolved["size_pt"].source == "direct"

    def it_reports_unspecified_toggles_as_false_from_none(self):
        document = _doc()
        run = document.add_paragraph().add_run("plain")
        resolved = format_of(run)
        assert resolved["italic"].value is False
        assert resolved["italic"].source == "none"

    def it_reads_doc_defaults(self):
        document = _doc()
        run = document.add_paragraph().add_run("default-sized")
        resolved = format_of(run)
        assert resolved["size_pt"].value == 11
        assert resolved["size_pt"].source == "doc_defaults"
        # the template's docDefaults use a THEME font reference; the honest
        # answer is the token, never a guessed literal
        assert resolved["font_name"].value == "theme:minorHAnsi"
        assert "theme_font_resolution" in resolved.unresolved

    def it_declares_what_it_cannot_resolve(self):
        document = _doc()
        resolved = format_of(document.paragraphs[1].runs[0])
        assert "table_style_conditional_formatting" in resolved.unresolved
        payload = resolved.to_dict()
        assert payload["schema"] == "paper_effective_format"


class DescribeStyleChainResolution:
    def it_resolves_through_the_paragraph_style_chain(self):
        document = _doc()
        base = document.styles.add_style("ChainBase", WD_STYLE_TYPE.PARAGRAPH)
        base.font.size = Pt(16)
        leaf = document.styles.add_style("ChainLeaf", WD_STYLE_TYPE.PARAGRAPH)
        leaf.base_style = base
        paragraph = document.add_paragraph("chained text", style="ChainLeaf")
        resolved = format_of(paragraph.runs[0])
        assert resolved["size_pt"].value == 16
        assert resolved["size_pt"].source == "paragraph_style:ChainBase"

    def it_lets_the_nearer_layer_win_for_ordinary_properties(self):
        document = _doc()
        base = document.styles.add_style("SizedBase", WD_STYLE_TYPE.PARAGRAPH)
        base.font.size = Pt(16)
        leaf = document.styles.add_style("SizedLeaf", WD_STYLE_TYPE.PARAGRAPH)
        leaf.base_style = base
        leaf.font.size = Pt(9)
        paragraph = document.add_paragraph("near wins", style="SizedLeaf")
        resolved = format_of(paragraph.runs[0])
        assert resolved["size_pt"].value == 9
        assert resolved["size_pt"].source == "paragraph_style:SizedLeaf"

    def it_resolves_character_styles_above_paragraph_styles(self):
        document = _doc()
        para_style = document.styles.add_style("ColoredPara", WD_STYLE_TYPE.PARAGRAPH)
        para_style.font.size = Pt(15)
        char_style = document.styles.add_style("Emph", WD_STYLE_TYPE.CHARACTER)
        char_style.font.size = Pt(8)
        paragraph = document.add_paragraph(style="ColoredPara")
        run = paragraph.add_run("char styled")
        run.style = char_style
        resolved = format_of(run)
        assert resolved["size_pt"].value == 8
        assert resolved["size_pt"].source == "character_style:Emph"

    def it_survives_a_based_on_cycle(self):
        from docx.oxml.ns import nsdecls
        from docx.oxml.parser import parse_xml

        document = _doc()
        styles = document.styles.element
        w = nsdecls("w")
        styles.append(parse_xml(
            f'<w:style {w} w:type="paragraph" w:styleId="CycleA">'
            '<w:name w:val="Cycle A"/><w:basedOn w:val="CycleB"/></w:style>'
        ))
        styles.append(parse_xml(
            f'<w:style {w} w:type="paragraph" w:styleId="CycleB">'
            '<w:name w:val="Cycle B"/><w:basedOn w:val="CycleA"/>'
            "<w:rPr><w:sz w:val=\"40\"/></w:rPr></w:style>"
        ))
        paragraph = document.add_paragraph("cyclic")
        paragraph.style = document.styles["Cycle A"]
        resolved = format_of(paragraph.runs[0])  # must not recurse forever
        assert resolved["size_pt"].value == 20


class DescribeToggleSemantics:
    """The famous gotcha: nested toggles XOR through style layers."""

    def it_xors_bold_when_paragraph_and_character_styles_both_set_it(self):
        document = _doc()
        para_style = document.styles.add_style("BoldPara", WD_STYLE_TYPE.PARAGRAPH)
        para_style.font.bold = True
        char_style = document.styles.add_style("BoldChar", WD_STYLE_TYPE.CHARACTER)
        char_style.font.bold = True
        paragraph = document.add_paragraph(style="BoldPara")
        run = paragraph.add_run("double bold cancels")
        run.style = char_style
        resolved = format_of(run)
        assert resolved["bold"].value is False  # True XOR True
        assert resolved["bold"].source == "toggle_xor"
        assert resolved["bold"].chain == (
            "paragraph_style:BoldPara",
            "character_style:BoldChar",
        )

    def it_keeps_bold_from_a_single_style_layer(self):
        document = _doc()
        para_style = document.styles.add_style("JustBold", WD_STYLE_TYPE.PARAGRAPH)
        para_style.font.bold = True
        paragraph = document.add_paragraph("single layer", style="JustBold")
        resolved = format_of(paragraph.runs[0])
        assert resolved["bold"].value is True
        assert resolved["bold"].source == "paragraph_style:JustBold"

    def it_lets_direct_formatting_override_toggles_absolutely(self):
        document = _doc()
        para_style = document.styles.add_style("BoldBase", WD_STYLE_TYPE.PARAGRAPH)
        para_style.font.bold = True
        paragraph = document.add_paragraph(style="BoldBase")
        run = paragraph.add_run("explicitly unbold")
        run.bold = False
        resolved = format_of(run)
        assert resolved["bold"].value is False
        assert resolved["bold"].source == "direct"


class DescribeParagraphAndSpanTargets:
    def it_resolves_paragraph_alignment_with_provenance(self):
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        document = _doc()
        paragraph = document.add_paragraph("centered")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        resolved = format_of(paragraph)
        assert resolved["alignment"].value == "center"
        assert resolved["alignment"].source == "direct"
        assert resolved["style_name"].value == "Normal"

    def it_resolves_a_span_and_reports_disagreement_as_mixed(self):
        document = _doc()
        paragraph = document.add_paragraph()
        paragraph.add_run("bold half").bold = True
        paragraph.add_run(" plain half")
        span = find_one(document, "bold half plain half")
        resolved = format_of(span)
        assert resolved["bold"].value is None
        assert resolved["bold"].source == "mixed"
        assert resolved["size_pt"].value == 11  # layers agree

    def it_rejects_unsupported_targets(self):
        with pytest.raises(TypeError, match="format_of"):
            format_of("just a string")


class DescribeSurroundingFormat:
    def it_resolves_the_exact_inline_target(self):
        document = _doc()
        resolved = surrounding_format(document, "Minimal Clean Document")
        assert resolved["style_name"].value == "Heading 1"
        # Heading 1 in the default template resolves sz through its chain
        assert resolved["size_pt"].value is not None

    @pytest.mark.parametrize("as_string", [True, False])
    def it_refuses_a_cross_paragraph_target(self, as_string: bool):
        document = docx.Document()
        document.add_paragraph("alpha end")
        document.add_paragraph("beta start")
        span = find_one(document, "alpha end\nbeta start")
        target = span.text if as_string else span

        with pytest.raises(BoundaryViolationError, match="one paragraph"):
            surrounding_format(document, target)

    @pytest.mark.parametrize("as_span", [False, True])
    @pytest.mark.parametrize(
        ("text", "property_name", "expected"),
        [
            ("first target", "bold", True),
            ("middle target", "italic", True),
            ("last target", "size_pt", 18),
        ],
    )
    def it_resolves_first_middle_and_last_run_targets_locally(
        self, text: str, property_name: str, expected: object, as_span: bool
    ):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("first target").bold = True
        paragraph.add_run(" middle target").italic = True
        paragraph.add_run(" last target").font.size = Pt(18)
        target = find_one(document, text) if as_span else text

        resolved = surrounding_format(document, target)

        assert resolved[property_name].value == expected
        assert resolved[property_name].source == "direct"

    def it_keeps_a_supplied_span_identity_when_its_text_is_duplicated(self):
        document = docx.Document()
        document.add_paragraph("duplicate target")
        paragraph = document.add_paragraph()
        paragraph.add_run("duplicate target").bold = True
        second = find_one(document, "duplicate target", nth=2)

        resolved = surrounding_format(document, second)

        assert resolved["bold"].value is True
        assert resolved["bold"].source == "direct"

    def it_accepts_an_explicitly_normalized_span_without_searching_again(self):
        document = docx.Document()
        paragraph = document.add_paragraph()
        paragraph.add_run("Smart—Target").italic = True
        span = find_one(document, "smart-target", match="normalized")
        assert span.text == "Smart—Target"

        resolved = surrounding_format(document, span)

        assert resolved["italic"].value is True
        assert resolved["italic"].source == "direct"

    def it_preserves_mixed_values_and_agreeing_layer_provenance(self):
        document = docx.Document()
        character_style = document.styles.add_style(
            "AgreeingBold", WD_STYLE_TYPE.CHARACTER
        )
        character_style.font.bold = True
        paragraph = document.add_paragraph()
        direct = paragraph.add_run("direct region")
        direct.bold = True
        direct.font.size = Pt(10)
        styled = paragraph.add_run(" styled region")
        styled.style = character_style
        styled.font.size = Pt(16)
        span = find_one(document, "direct region styled region")

        resolved = surrounding_format(document, span)

        assert resolved["size_pt"].value is None
        assert resolved["size_pt"].source == "mixed"
        assert resolved["bold"].value is True
        assert resolved["bold"].source == "agreeing_layers"
        assert "direct" in resolved["bold"].chain
        assert "character_style:AgreeingBold" in resolved["bold"].chain

    @pytest.mark.parametrize("empty", [False, True])
    @pytest.mark.parametrize("as_locator", [False, True])
    def it_resolves_block_targets_to_paragraph_defaults(
        self, empty: bool, as_locator: bool
    ):
        document = docx.Document()
        style = document.styles.add_style("BlockDefaults", WD_STYLE_TYPE.PARAGRAPH)
        style.font.size = Pt(15)
        paragraph = document.add_paragraph(style="BlockDefaults")
        if not empty:
            run = paragraph.add_run("populated block")
            run.bold = True
            run.font.size = Pt(27)
        block = next(
            item for item in iter_blocks(document) if item._element is paragraph._p
        )
        target = block
        if as_locator:
            assert block.locator is not None
            target = BlockLocator.from_dict(block.locator.to_dict())

        resolved = surrounding_format(document, target)

        assert resolved["size_pt"].value == 15
        assert resolved["size_pt"].source == "paragraph_style:BlockDefaults"
        assert resolved["bold"].value is False

    @pytest.mark.parametrize("as_span", [False, True])
    def it_resolves_inline_targets_inside_a_table_cell(self, as_span: bool):
        document = docx.Document()
        paragraph = document.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0]
        paragraph.add_run("cell target").bold = True
        target = find_one(document, "cell target") if as_span else "cell target"

        resolved = surrounding_format(document, target)

        assert resolved["bold"].value is True
        assert resolved["bold"].source == "direct"

    def it_refuses_live_and_portable_table_block_targets(self):
        document = docx.Document()
        document.add_table(rows=1, cols=1).cell(0, 0).text = "cell text"
        table_block = next(block for block in iter_blocks(document) if block.kind == "table")
        assert table_block.locator is not None

        for target in (table_block, table_block.locator):
            with pytest.raises(UnsupportedStructureError, match="table block"):
                surrounding_format(document, target)

    def it_preserves_foreign_and_stale_live_target_refusals(self):
        source = docx.Document()
        source.add_paragraph("foreign span")
        foreign_span = find_one(source, "foreign span")
        foreign_block = next(iter(iter_blocks(source)))
        destination = docx.Document()
        destination.add_paragraph("local")

        for target in (foreign_span, foreign_block):
            with pytest.raises(BoundaryViolationError, match="different document"):
                surrounding_format(destination, target)

        stale_document = docx.Document()
        stale_paragraph = stale_document.add_paragraph("stale target")
        stale_span = find_one(stale_document, "stale target")
        stale_block = next(iter(iter_blocks(stale_document)))
        stale_paragraph._p.getparent().remove(stale_paragraph._p)
        for target in (stale_span, stale_block):
            with pytest.raises(TargetNotFoundError, match="stale"):
                surrounding_format(stale_document, target)

    def it_preserves_missing_and_ambiguous_locator_refusals(self):
        missing_document = docx.Document()
        missing_document.add_paragraph("before")
        paragraph = missing_document.add_paragraph("target")
        missing_document.add_paragraph("after")
        missing_locator = tuple(iter_blocks(missing_document))[1].locator
        assert missing_locator is not None
        paragraph.text = "changed"
        with pytest.raises(TargetNotFoundError, match="stale"):
            surrounding_format(missing_document, missing_locator)

        ambiguous_document = docx.Document()
        for text in ("same", "target", "same", "target", "same"):
            ambiguous_document.add_paragraph(text)
        ambiguous_locator = tuple(iter_blocks(ambiguous_document))[1].locator
        assert ambiguous_locator is not None
        with pytest.raises(AmbiguousTargetError, match="matches 2 blocks"):
            surrounding_format(ambiguous_document, ambiguous_locator)

    def it_refuses_legacy_anchor_evidence(self):
        document = docx.Document()
        document.add_paragraph("target")
        anchor = next(iter(iter_blocks(document))).anchor

        with pytest.raises(UnsupportedStructureError, match="inert location evidence"):
            surrounding_format(document, anchor)

    def it_preserves_missing_and_ambiguous_exact_string_refusals(self):
        document = docx.Document()
        document.add_paragraph("duplicate")
        document.add_paragraph("duplicate")

        with pytest.raises(TargetNotFoundError, match="no match"):
            surrounding_format(document, "missing")
        with pytest.raises(AmbiguousTargetError, match="2 matches"):
            surrounding_format(document, "duplicate")
