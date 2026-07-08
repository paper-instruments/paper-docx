"""Numbering enumeration and application (paper-docx, narrow and guarded).

v0 applies EXISTING numbering definitions and list styles to paragraphs, and
reports what a document defines and uses. Authoring new `numbering.xml`
definitions (including restart/continue mechanics that require them) is
explicitly out of v0 scope and refuses loudly.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Optional, Tuple

from docx.errors import TargetNotFoundError
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.story import _build_block, _iter_block_elements, _story_elements

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document
    from docx.text.paragraph import Paragraph


@dataclass(frozen=True)
class NumberingLevel:
    level: int
    num_fmt: Optional[str]
    lvl_text: Optional[str]

    def to_dict(self) -> dict:
        return {"level": self.level, "num_fmt": self.num_fmt, "lvl_text": self.lvl_text}


@dataclass(frozen=True)
class NumberingDefinition:
    num_id: int
    abstract_num_id: int
    levels: Tuple[NumberingLevel, ...]

    def to_dict(self) -> dict:
        return {
            "num_id": self.num_id,
            "abstract_num_id": self.abstract_num_id,
            "levels": [level.to_dict() for level in self.levels],
        }


@dataclass(frozen=True)
class NumberedParagraph:
    story: str
    index: int
    num_id: int
    level: int
    text: str

    def to_dict(self) -> dict:
        return {
            "story": self.story,
            "index": self.index,
            "num_id": self.num_id,
            "level": self.level,
            "text": self.text,
        }


@dataclass(frozen=True)
class NumberingReport:
    definitions: Tuple[NumberingDefinition, ...]
    numbered_paragraphs: Tuple[NumberedParagraph, ...]

    def to_dict(self) -> dict:
        return {
            "schema": "paper_numbering",
            "version": 1,
            "definitions": [d.to_dict() for d in self.definitions],
            "numbered_paragraphs": [p.to_dict() for p in self.numbered_paragraphs],
        }


def _numbering_root(document: "Document") -> "Optional[_Element]":
    """The `w:numbering` root element, or None when the part doesn't exist.

    Deliberately avoids `DocumentPart.numbering_part`, whose auto-create path
    hits upstream's NotImplementedError stub when the part is absent.
    """
    try:
        part = document.part.part_related_by(RT.NUMBERING)
    except KeyError:
        return None
    return part._element


def _definitions(numbering: "Optional[_Element]") -> Tuple[NumberingDefinition, ...]:
    if numbering is None:
        return ()
    abstract_levels = {}
    for abstract in numbering.findall(qn("w:abstractNum")):
        abstract_id = int(abstract.get(qn("w:abstractNumId")))
        levels = []
        for lvl in abstract.findall(qn("w:lvl")):
            num_fmt = lvl.find(qn("w:numFmt"))
            lvl_text = lvl.find(qn("w:lvlText"))
            levels.append(
                NumberingLevel(
                    level=int(lvl.get(qn("w:ilvl"))),
                    num_fmt=num_fmt.get(qn("w:val")) if num_fmt is not None else None,
                    lvl_text=lvl_text.get(qn("w:val")) if lvl_text is not None else None,
                )
            )
        abstract_levels[abstract_id] = tuple(levels)
    definitions = []
    for num in numbering.findall(qn("w:num")):
        num_id = int(num.get(qn("w:numId")))
        abstract_ref = num.find(qn("w:abstractNumId"))
        abstract_id = int(abstract_ref.get(qn("w:val"))) if abstract_ref is not None else -1
        definitions.append(
            NumberingDefinition(
                num_id=num_id,
                abstract_num_id=abstract_id,
                levels=abstract_levels.get(abstract_id, ()),
            )
        )
    return tuple(sorted(definitions, key=lambda d: d.num_id))


def list_numbering(document: "Document") -> NumberingReport:
    """Every numbering definition and every numbered paragraph in `document`."""
    numbered = []
    for story, root in _story_elements(document):
        for kind, index, element, in_sdt, in_txbx in _iter_block_elements(story, root):
            if kind != "paragraph":
                continue
            num_prs = element.xpath("./w:pPr/w:numPr")
            if not num_prs:
                continue
            num_pr = num_prs[0]
            num_id_elm = num_pr.find(qn("w:numId"))
            ilvl_elm = num_pr.find(qn("w:ilvl"))
            if num_id_elm is None:
                continue
            block = _build_block(
                story, kind, index, element, "current", in_sdt=in_sdt, in_txbx=in_txbx
            )
            numbered.append(
                NumberedParagraph(
                    story=story,
                    index=index,
                    num_id=int(num_id_elm.get(qn("w:val"))),
                    level=int(ilvl_elm.get(qn("w:val"))) if ilvl_elm is not None else 0,
                    text=block.text,
                )
            )
    return NumberingReport(
        definitions=_definitions(_numbering_root(document)),
        numbered_paragraphs=tuple(numbered),
    )


def _document_of_paragraph(paragraph: "Paragraph") -> "Document":
    part = paragraph.part
    document = getattr(part, "document", None)
    if document is None:
        document = part._document_part.document  # noqa: SLF001 - StoryPart broker
    return document


def apply_numbering(paragraph: "Paragraph", *, num_id: int, level: int = 0) -> None:
    """Give `paragraph` the existing numbering definition `num_id` at `level`.

    The definition must exist in word/numbering.xml and define `level` —
    otherwise |TargetNotFoundError|; this API never fabricates definitions
    (authoring numbering.xml is out of v0).
    """
    document = _document_of_paragraph(paragraph)
    report_definitions = _definitions(_numbering_root(document))
    definition = next((d for d in report_definitions if d.num_id == num_id), None)
    if definition is None:
        defined = [d.num_id for d in report_definitions]
        raise TargetNotFoundError(
            f"numbering definition numId={num_id} does not exist"
            f" (defined: {defined}); authoring new definitions is out of v0"
        )
    if level not in {lvl.level for lvl in definition.levels}:
        raise TargetNotFoundError(
            f"numbering definition numId={num_id} does not define level {level}"
        )

    # -- validated; mutate --
    num_pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    num_pr.get_or_add_ilvl().val = level
    num_pr.get_or_add_numId().val = num_id


def _style_numbering_binding(document: "Document", style_name: str) -> Optional[int]:
    """The numId `style_name` binds (following the basedOn chain), or None.

    numId 0 means "no numbering" per ECMA-376 and reads as no binding.
    """
    style = document.styles[style_name]
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        values = style.element.xpath("./w:pPr/w:numPr/w:numId/@w:val")
        if values:
            num_id = int(values[0])
            return num_id if num_id != 0 else None
        style = style.base_style
    return None


def apply_list_style(paragraph: "Paragraph", style_name: str) -> None:
    """Apply the existing paragraph style named `style_name` (e.g. a list
    style like "List Bullet") — |TargetNotFoundError| when undefined.

    When the style binds a numbering definition, that definition must
    actually resolve in word/numbering.xml; otherwise this call refuses
    rather than producing the classic FAKE bullet (a list-styled paragraph
    that renders with no marker — v0.1 honesty recall, H7). Styles with no
    numbering binding apply as plain styles.
    """
    document = _document_of_paragraph(paragraph)
    defined = {style.name for style in document.styles}
    if style_name not in defined:
        raise TargetNotFoundError(
            f"style {style_name!r} is not defined in this document"
        )
    bound_num_id = _style_numbering_binding(document, style_name)
    if bound_num_id is not None:
        definitions = _definitions(_numbering_root(document))
        if not any(d.num_id == bound_num_id for d in definitions):
            raise TargetNotFoundError(
                f"style {style_name!r} binds numbering definition"
                f" numId={bound_num_id}, which does not resolve in this"
                " document — applying it would render a fake, marker-less"
                " list; add a real definition first"
            )
    paragraph.style = style_name
