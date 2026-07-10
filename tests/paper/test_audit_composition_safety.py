"""Atomicity regressions for cross-document composition."""

from __future__ import annotations

import pytest

import docx
from docx.composition import insert_blocks_from
from docx.enum.style import WD_STYLE_TYPE
from docx.errors import UnsupportedStructureError
from docx.numbering import apply_numbering, ensure_decimal_definition
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement


def _package_state(document) -> tuple:
    parts = []
    for part in document.part.package.iter_parts():
        relationships = tuple(
            sorted(
                (
                    rel.rId,
                    rel.reltype,
                    str(rel.target_ref),
                    rel.is_external,
                )
                for rel in part.rels.values()
            )
        )
        parts.append((str(part.partname), part.blob, relationships))
    return tuple(sorted(parts, key=lambda item: item[0]))


def _append_chart_reference(document, paragraph, r_id: str = "rId99") -> None:
    chart_part = Part(
        PackURI("/word/charts/chart1.xml"),
        CT.DML_CHART,
        b'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
        document.part.package,
    )
    document.part.load_rel(RT.CHART, chart_part, r_id)
    drawing = OxmlElement("w:drawing")
    chart = OxmlElement("c:chart")
    chart.set(qn("r:id"), r_id)
    drawing.append(chart)
    paragraph.add_run()._r.append(drawing)


def _append_style_reference(paragraph, tag: str, style_id: str) -> None:
    properties = (
        paragraph._p.get_or_add_pPr()
        if tag == "w:pStyle"
        else paragraph.runs[0]._r.get_or_add_rPr()
    )
    reference = OxmlElement(tag)
    reference.set(qn("w:val"), style_id)
    properties.append(reference)


def it_refuses_a_chart_relationship_before_copying_stale_rids() -> None:
    source = docx.Document()
    paragraph = source.add_paragraph("Source chart")
    _append_chart_reference(source, paragraph)
    destination = docx.Document()
    destination.add_paragraph("Destination anchor")
    before = _package_state(destination)

    with pytest.raises(UnsupportedStructureError, match="chart relationship"):
        insert_blocks_from(
            destination,
            source,
            "Source chart",
            anchor="Destination anchor",
        )

    assert _package_state(destination) == before
    assert "rId99" not in destination.element.xml


def it_refuses_a_source_num_whose_abstract_definition_is_missing() -> None:
    source = docx.Document()
    num_id = ensure_decimal_definition(source)
    paragraph = source.add_paragraph("Source numbered paragraph")
    apply_numbering(paragraph, num_id=num_id)
    source_numbering = source.part.numbering_part.element
    source_num = next(
        num
        for num in source_numbering.findall(qn("w:num"))
        if num.get(qn("w:numId")) == str(num_id)
    )
    abstract_id = source_num.find(qn("w:abstractNumId")).get(qn("w:val"))
    source_abstract = next(
        abstract
        for abstract in source_numbering.findall(qn("w:abstractNum"))
        if abstract.get(qn("w:abstractNumId")) == abstract_id
    )
    source_numbering.remove(source_abstract)
    destination = docx.Document()
    destination.add_paragraph("Destination anchor")
    before = _package_state(destination)

    with pytest.raises(UnsupportedStructureError, match="missing abstract"):
        insert_blocks_from(
            destination,
            source,
            "Source numbered paragraph",
            anchor="Destination anchor",
        )

    assert _package_state(destination) == before


def it_refuses_malformed_destination_numbering_before_importing_a_style() -> None:
    source = docx.Document()
    style = source.styles.add_style("AuditCompositionStyle", WD_STYLE_TYPE.PARAGRAPH)
    num_id = ensure_decimal_definition(source)
    paragraph = source.add_paragraph("Styled numbered source", style=style)
    apply_numbering(paragraph, num_id=num_id)
    destination = docx.Document()
    destination.add_paragraph("Destination anchor")
    ensure_decimal_definition(destination)
    destination_num = destination.part.numbering_part.element.find(qn("w:num"))
    destination_num.set(qn("w:numId"), "not-a-number")
    before = _package_state(destination)

    with pytest.raises(UnsupportedStructureError, match="destination numbering.*non-numeric"):
        insert_blocks_from(
            destination,
            source,
            "Styled numbered source",
            anchor="Destination anchor",
        )

    assert _package_state(destination) == before
    assert "AuditCompositionStyle" not in {style.name for style in destination.styles}


@pytest.mark.parametrize(
    ("reference_tag", "destination_style_type"),
    [
        pytest.param("w:pStyle", WD_STYLE_TYPE.PARAGRAPH, id="paragraph-style"),
        pytest.param("w:rStyle", WD_STYLE_TYPE.CHARACTER, id="character-style"),
    ],
)
def it_refuses_undefined_source_style_references_before_destination_binding(
    reference_tag: str,
    destination_style_type: WD_STYLE_TYPE,
) -> None:
    source = docx.Document()
    paragraph = source.add_paragraph("Source with an undefined style")
    _append_style_reference(paragraph, reference_tag, "SharedStyleId")
    destination = docx.Document()
    destination.add_paragraph("Destination anchor")
    destination_style = destination.styles.add_style(
        "Unrelated destination style", destination_style_type
    )
    destination_style._element.set(qn("w:styleId"), "SharedStyleId")
    before = _package_state(destination)

    with pytest.raises(
        UnsupportedStructureError, match="undefined source style 'SharedStyleId'"
    ):
        insert_blocks_from(
            destination,
            source,
            "Source with an undefined style",
            anchor="Destination anchor",
        )

    assert _package_state(destination) == before
    assert all(
        paragraph.text != "Source with an undefined style"
        for paragraph in destination.paragraphs
    )


@pytest.mark.parametrize("chain_tag", ["w:basedOn", "w:link", "w:next"])
def it_refuses_undefined_source_style_chain_dependencies_atomically(chain_tag: str) -> None:
    source = docx.Document()
    source_style = source.styles.add_style("Source chain leaf", WD_STYLE_TYPE.PARAGRAPH)
    dependency = OxmlElement(chain_tag)
    dependency.set(qn("w:val"), "MissingDependency")
    source_style._element.append(dependency)
    source.add_paragraph("Source with a broken style chain", style=source_style)
    destination = docx.Document()
    destination.add_paragraph("Destination anchor")
    destination_style = destination.styles.add_style(
        "Unrelated chain destination", WD_STYLE_TYPE.PARAGRAPH
    )
    destination_style._element.set(qn("w:styleId"), "MissingDependency")
    before = _package_state(destination)

    with pytest.raises(
        UnsupportedStructureError, match="undefined source style 'MissingDependency'"
    ):
        insert_blocks_from(
            destination,
            source,
            "Source with a broken style chain",
            anchor="Destination anchor",
        )

    assert _package_state(destination) == before
    assert "Source chain leaf" not in {style.name for style in destination.styles}
