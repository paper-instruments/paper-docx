"""Numbering enumeration, application, and minimal authoring (paper-docx).

Applies EXISTING numbering definitions and list styles to paragraphs and
reports what a document defines and uses. As of v0.1 (V2), it also ships
exactly two canonical definitions — one bullet, one decimal —
creatable on demand (`ensure_bullet_definition` / `ensure_decimal_definition`,
idempotent), plus level-0 restarts (`restart_numbering`). Anything more
exotic (custom level text, legal numbering, image bullets) stays out of scope
and refuses loudly.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Optional, Tuple

from docx.errors import TargetNotFoundError
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.protection import _refuse_if_protected
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
    (create one deliberately with `ensure_bullet_definition` /
    `ensure_decimal_definition`).
    """
    document = _document_of_paragraph(paragraph)
    _refuse_if_protected(document, "apply numbering")
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
    _refuse_if_protected(document, "apply a list style")
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


# ---------------------------------------------------------------------------
# v0.1 V2 — minimal, guarded numbering AUTHORING
# ---------------------------------------------------------------------------

#: canonical three-level definitions, shaped like Word's own defaults. The
#: w:name marker makes ensure_* idempotent (one shared definition per kind).
_CANONICAL_ABSTRACT_XML = {
    "bullet": (
        "PaperBulletList",
        '<w:abstractNum {ids}>'
        '<w:name w:val="PaperBulletList"/>'
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="•"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="o"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="▪"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr></w:lvl>'
        "</w:abstractNum>"
    ),
    "decimal": (
        "PaperDecimalList",
        '<w:abstractNum {ids}>'
        '<w:name w:val="PaperDecimalList"/>'
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
        '<w:lvlText w:val="%2."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/>'
        '<w:lvlText w:val="%3."/><w:lvlJc w:val="right"/>'
        '<w:pPr><w:ind w:left="2160" w:hanging="180"/></w:pPr></w:lvl>'
        "</w:abstractNum>"
    ),
}


def _get_or_create_numbering_root(document: "Document") -> "_Element":
    """The `w:numbering` root, creating the part when the document has none.

    Upstream's `NumberingPart.new()` is a NotImplementedError stub, so the
    part is assembled here: root element, part instance, relationship. The
    package writer emits the content-type override automatically on save.
    """
    root = _numbering_root(document)
    if root is not None:
        return root
    from docx.opc.constants import CONTENT_TYPE as CT
    from docx.opc.packuri import PackURI
    from docx.oxml.ns import nsdecls
    from docx.oxml.parser import parse_xml
    from docx.parts.numbering import NumberingPart

    root = parse_xml(f"<w:numbering {nsdecls('w')}/>")
    part = NumberingPart(
        PackURI("/word/numbering.xml"), CT.WML_NUMBERING, root, document.part.package
    )
    document.part.relate_to(part, RT.NUMBERING)
    return root


def _next_free_ids(numbering: "_Element") -> Tuple[int, int]:
    abstract_ids = [
        int(a.get(qn("w:abstractNumId")) or 0)
        for a in numbering.findall(qn("w:abstractNum"))
    ]
    num_ids = [int(n.get(qn("w:numId")) or 0) for n in numbering.findall(qn("w:num"))]
    return (max(abstract_ids, default=-1) + 1, max(num_ids, default=0) + 1)


def _append_definition(numbering: "_Element", abstract: "_Element", num: "_Element") -> None:
    """Insert respecting the schema order: abstractNum* before num*."""
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)
    last = numbering.findall(qn("w:num"))
    if last:
        last[-1].addnext(num)
    else:
        numbering.append(num)


def _ensure_definition(document: "Document", kind: str) -> int:
    from docx.oxml.ns import nsdecls
    from docx.oxml.parser import parse_xml

    name, template = _CANONICAL_ABSTRACT_XML[kind]
    numbering = _get_or_create_numbering_root(document)
    # idempotent: reuse the canonical definition if it is already here
    for abstract in numbering.findall(qn("w:abstractNum")):
        name_elm = abstract.find(qn("w:name"))
        if name_elm is not None and name_elm.get(qn("w:val")) == name:
            abstract_id = abstract.get(qn("w:abstractNumId"))
            for num in numbering.findall(qn("w:num")):
                ref = num.find(qn("w:abstractNumId"))
                if ref is not None and ref.get(qn("w:val")) == abstract_id:
                    return int(num.get(qn("w:numId")))
    abstract_id, num_id = _next_free_ids(numbering)
    ids = f'{nsdecls("w")} w:abstractNumId="{abstract_id}"'
    abstract = parse_xml(template.format(ids=ids))
    num = parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/></w:num>'
    )
    _append_definition(numbering, abstract, num)
    return num_id


def ensure_bullet_definition(document: "Document") -> int:
    """The numId of a real bullet-list definition, creating one when the
    document has none (v0.1 V2 — closes the 'cannot make a real bullet'
    gap). Idempotent: repeated calls return the same definition."""
    _refuse_if_protected(document, "author a numbering definition")
    return _ensure_definition(document, "bullet")


def ensure_decimal_definition(document: "Document") -> int:
    """The numId of a real decimal-list definition, created on demand.
    Idempotent, like `ensure_bullet_definition`."""
    _refuse_if_protected(document, "author a numbering definition")
    return _ensure_definition(document, "decimal")


def restart_numbering(document: "Document", *, num_id: int) -> int:
    """A NEW numId continuing `num_id`'s formatting but restarting at 1.

    Word models restarts as a fresh `w:num` referencing the same abstract
    definition with a level-0 `w:startOverride`; paragraphs re-point at the
    returned numId. Anything fancier (mid-list overrides, custom level text)
    stays out of scope and refuses via the ordinary numId validation.
    """
    _refuse_if_protected(document, "restart numbering")
    numbering = _numbering_root(document)
    definition = None
    if numbering is not None:
        for num in numbering.findall(qn("w:num")):
            if int(num.get(qn("w:numId"))) == num_id:
                definition = num
                break
    if definition is None:
        raise TargetNotFoundError(
            f"numbering definition numId={num_id} does not exist; nothing to restart"
        )
    from docx.oxml.ns import nsdecls
    from docx.oxml.parser import parse_xml

    abstract_id = definition.find(qn("w:abstractNumId")).get(qn("w:val"))
    _, new_num_id = _next_free_ids(numbering)
    new_num = parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{new_num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/>'
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
        "</w:num>"
    )
    last = numbering.findall(qn("w:num"))
    last[-1].addnext(new_num)
    return new_num_id
