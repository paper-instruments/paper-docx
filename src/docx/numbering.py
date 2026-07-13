"""Numbering enumeration, application, and minimal authoring (paper-docx).

Applies EXISTING numbering definitions and list styles to paragraphs and
reports what a document defines and uses. It also ships
exactly two canonical definitions — one bullet, one decimal —
creatable on demand (`ensure_bullet_definition` / `ensure_decimal_definition`,
idempotent), plus level-0 restarts (`restart_numbering`). Anything more
exotic (custom level text, legal numbering, image bullets) stays out of scope
and refuses loudly.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, Optional, Tuple

from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.protection import _refuse_if_protected
from docx.story import (
    _build_block,
    _first_choice_children,
    _iter_block_elements,
    _story_elements,
)

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
    table_cell: "Optional[Tuple[int, int, int]]" = None

    def to_dict(self) -> dict:
        return {
            "story": self.story,
            "index": self.index,
            "num_id": self.num_id,
            "level": self.level,
            "text": self.text,
            "table_cell": self.table_cell,
        }


@dataclass(frozen=True)
class NumberingReport:
    definitions: Tuple[NumberingDefinition, ...]
    numbered_paragraphs: Tuple[NumberedParagraph, ...]

    def to_dict(self) -> dict:
        return {
            "schema": "paper_numbering",
            "version": 2,
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
        levels = {level.level: level for level in abstract_levels.get(abstract_id, ())}
        for override in num.findall(qn("w:lvlOverride")):
            override_level = int(override.get(qn("w:ilvl")))
            level = override.find(qn("w:lvl"))
            if level is None:
                continue
            num_fmt = level.find(qn("w:numFmt"))
            lvl_text = level.find(qn("w:lvlText"))
            levels[override_level] = NumberingLevel(
                level=override_level,
                num_fmt=num_fmt.get(qn("w:val")) if num_fmt is not None else None,
                lvl_text=lvl_text.get(qn("w:val")) if lvl_text is not None else None,
            )
        definitions.append(
            NumberingDefinition(
                num_id=num_id,
                abstract_num_id=abstract_id,
                levels=tuple(levels[index] for index in sorted(levels)),
            )
        )
    return tuple(sorted(definitions, key=lambda d: d.num_id))


def _styles_index(document: "Document") -> "dict[str, _Element]":
    return {
        style_id: style
        for style in document.styles.element.findall(qn("w:style"))
        if (style_id := style.get(qn("w:styleId"))) is not None
    }


def _default_paragraph_style_id(document: "Document") -> Optional[str]:
    default = None
    for style in document.styles.element.findall(qn("w:style")):
        if style.get(qn("w:type")) != "paragraph":
            continue
        if (style.get(qn("w:default")) or "").lower() in ("1", "true", "on"):
            default = style.get(qn("w:styleId"))
    return default


def _paragraph_style_chain(
    styles: "dict[str, _Element]", style_id: Optional[str]
) -> "list[_Element]":
    chain = []
    seen = set()
    current = style_id
    while current is not None and current in styles and current not in seen:
        seen.add(current)
        style = styles[current]
        chain.append(style)
        based_on = style.find(qn("w:basedOn"))
        current = based_on.get(qn("w:val")) if based_on is not None else None
    chain.reverse()
    return chain


def _effective_paragraph_numbering(
    document: "Document", paragraph: "_Element"
) -> "Optional[Tuple[int, int]]":
    """Effective ``(numId, ilvl)`` after paragraph-style inheritance."""
    p_pr = paragraph.find(qn("w:pPr"))
    style_elm = p_pr.find(qn("w:pStyle")) if p_pr is not None else None
    style_id = (
        style_elm.get(qn("w:val"))
        if style_elm is not None
        else _default_paragraph_style_id(document)
    )
    layers = [
        style.find(qn("w:pPr"))
        for style in _paragraph_style_chain(_styles_index(document), style_id)
    ]
    layers.append(p_pr)
    num_id = None
    level = None
    for layer in layers:
        num_pr = layer.find(qn("w:numPr")) if layer is not None else None
        if num_pr is None:
            continue
        num_id_elm = num_pr.find(qn("w:numId"))
        ilvl_elm = num_pr.find(qn("w:ilvl"))
        if num_id_elm is not None and num_id_elm.get(qn("w:val")) is not None:
            num_id = int(num_id_elm.get(qn("w:val")))
        if ilvl_elm is not None and ilvl_elm.get(qn("w:val")) is not None:
            level = int(ilvl_elm.get(qn("w:val")))
    # ECMA-376 reserves numId 0 as an explicit "no numbering" override.
    if num_id in (None, 0):
        return None
    return num_id, level if level is not None else 0


def list_numbering(document: "Document") -> NumberingReport:
    """Every numbering definition and every numbered paragraph in `document`."""
    definitions = _definitions(_numbering_root(document))
    definitions_by_id = {definition.num_id: definition for definition in definitions}
    numbered = []
    for story, root in _story_elements(document):
        for kind, index, element, in_sdt, in_txbx in _iter_block_elements(story, root):
            paragraphs = (
                (element,)
                if kind == "paragraph"
                else tuple(_table_paragraphs(element))
                if kind == "table"
                else ()
            )
            for paragraph in paragraphs:
                effective = _effective_paragraph_numbering(document, paragraph)
                if effective is None:
                    continue
                num_id, level = effective
                definition = definitions_by_id.get(num_id)
                if definition is None:
                    raise UnsupportedStructureError(
                        f"paragraph resolves to missing numbering definition numId={num_id};"
                        " nothing was changed"
                    )
                if level not in {item.level for item in definition.levels}:
                    raise UnsupportedStructureError(
                        f"paragraph resolves to undefined numbering level {level}"
                        f" for numId={num_id}; nothing was changed"
                    )
                block = _build_block(
                    story,
                    "paragraph",
                    index,
                    paragraph,
                    "current",
                    in_sdt=in_sdt,
                    in_txbx=in_txbx,
                )
                numbered.append(
                    NumberedParagraph(
                        story=story,
                        index=index,
                        num_id=num_id,
                        level=level,
                        text=block.text,
                        table_cell=_table_cell_address(paragraph),
                    )
                )
    return NumberingReport(
        definitions=definitions,
        numbered_paragraphs=tuple(numbered),
    )


def _selected_descendants(
    root: "_Element", target_tag: str, *, stop_tags: "Tuple[str, ...]"
) -> "Iterator[_Element]":
    """Yield one compatibility-selected subtree, stopping at nested structures."""
    for child in _first_choice_children(root):
        if child.tag == target_tag:
            yield child
        if child.tag not in stop_tags:
            yield from _selected_descendants(
                child, target_tag, stop_tags=stop_tags
            )


def _table_rows(table: "_Element") -> "Tuple[_Element, ...]":
    return tuple(
        _selected_descendants(table, qn("w:tr"), stop_tags=(qn("w:tbl"),))
    )


def _row_cells(row: "_Element") -> "Tuple[_Element, ...]":
    return tuple(
        _selected_descendants(row, qn("w:tc"), stop_tags=(qn("w:tbl"),))
    )


def _cell_paragraphs(cell: "_Element") -> "Tuple[_Element, ...]":
    return tuple(
        _selected_descendants(cell, qn("w:p"), stop_tags=(qn("w:tbl"),))
    )


def _table_paragraphs(table: "_Element") -> "Iterator[_Element]":
    for row in _table_rows(table):
        for cell in _row_cells(row):
            yield from _cell_paragraphs(cell)


def _table_cell_address(paragraph: "_Element") -> "Optional[Tuple[int, int, int]]":
    tc = paragraph.getparent()
    while tc is not None and tc.tag not in (qn("w:tc"), qn("w:body")):
        tc = tc.getparent()
    if tc is None or tc.tag != qn("w:tc"):
        return None
    tr = tc.getparent()
    while tr is not None and tr.tag not in (qn("w:tr"), qn("w:body")):
        tr = tr.getparent()
    if tr is None or tr.tag != qn("w:tr"):
        return None
    tbl = tr.getparent()
    while tbl is not None and tbl.tag not in (qn("w:tbl"), qn("w:body")):
        tbl = tbl.getparent()
    if tbl is None or tbl.tag != qn("w:tbl"):
        return None
    rows = _table_rows(tbl)
    cells = _row_cells(tr)
    column = 0
    tr_pr = tr.find(qn("w:trPr"))
    grid_before = tr_pr.find(qn("w:gridBefore")) if tr_pr is not None else None
    if grid_before is not None:
        column = int(grid_before.get(qn("w:val")) or 0)
    for cell in cells:
        if cell is tc:
            break
        tc_pr = cell.find(qn("w:tcPr"))
        span = tc_pr.find(qn("w:gridSpan")) if tc_pr is not None else None
        column += int(span.get(qn("w:val")) or 1) if span is not None else 1
    return rows.index(tr), column, _cell_paragraphs(tc).index(paragraph)


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
            f" (defined: {defined}); authoring new definitions is not supported"
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
    that renders with no marker). Styles with no
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
# minimal, guarded numbering AUTHORING
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
        '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"'
        ' w:hint="default"/></w:rPr></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="▪"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"'
        ' w:hint="default"/></w:rPr></w:lvl>'
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
        cleanup = numbering.find(qn("w:numIdMacAtCleanup"))
        if cleanup is not None:
            cleanup.addprevious(num)
        else:
            numbering.append(num)


def _structural_signature(
    element: "_Element", *, ignore_root_abstract_id: bool = False
) -> tuple:
    """Prefix/attribute-order-independent exact XML structure signature."""
    abstract_id_attr = qn("w:abstractNumId")
    attributes = tuple(
        sorted(
            (name, value)
            for name, value in element.attrib.items()
            if not (ignore_root_abstract_id and name == abstract_id_attr)
        )
    )
    text = element.text or ""
    if not text.strip():
        text = ""
    return (
        element.tag,
        attributes,
        text,
        tuple(_structural_signature(child) for child in element),
    )


def _canonical_abstract(kind: str, abstract_id: int) -> "_Element":
    from docx.oxml.ns import nsdecls
    from docx.oxml.parser import parse_xml

    _name, template = _CANONICAL_ABSTRACT_XML[kind]
    ids = f'{nsdecls("w")} w:abstractNumId="{abstract_id}"'
    return parse_xml(template.format(ids=ids))


def _is_canonical_abstract(abstract: "_Element", kind: str) -> bool:
    expected = _canonical_abstract(kind, 0)
    return _structural_signature(
        abstract, ignore_root_abstract_id=True
    ) == _structural_signature(expected, ignore_root_abstract_id=True)


def _is_plain_num_reference(num: "_Element", abstract_id: str) -> bool:
    """Whether ``num`` is exactly the canonical abstract reference shape."""
    children = list(num)
    if len(children) != 1 or children[0].tag != qn("w:abstractNumId"):
        return False
    ref = children[0]
    return (
        tuple(ref.attrib.items()) == ((qn("w:val"), abstract_id),)
        and set(num.attrib) == {qn("w:numId")}
    )


def _ensure_definition(document: "Document", kind: str) -> int:
    from docx.oxml.ns import nsdecls
    from docx.oxml.parser import parse_xml

    name, _template = _CANONICAL_ABSTRACT_XML[kind]
    numbering = _get_or_create_numbering_root(document)
    # A name is only a marker, not proof. Reuse requires the exact canonical
    # abstract shape and a plain w:num reference with no hidden overrides.
    for abstract in numbering.findall(qn("w:abstractNum")):
        name_elm = abstract.find(qn("w:name"))
        if (
            name_elm is None
            or name_elm.get(qn("w:val")) != name
            or not _is_canonical_abstract(abstract, kind)
        ):
            continue
        abstract_id = abstract.get(qn("w:abstractNumId"))
        if abstract_id is None:
            continue
        for num in numbering.findall(qn("w:num")):
            if _is_plain_num_reference(num, abstract_id):
                return int(num.get(qn("w:numId")))
    abstract_id, num_id = _next_free_ids(numbering)
    abstract = _canonical_abstract(kind, abstract_id)
    num = parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/></w:num>'
    )
    _append_definition(numbering, abstract, num)
    return num_id


def ensure_bullet_definition(document: "Document") -> int:
    """The numId of a real bullet-list definition, creating one when the
    document has none. Idempotent: repeated calls return the same
    definition."""
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
            raw_num_id = num.get(qn("w:numId"))
            try:
                candidate_num_id = int(raw_num_id) if raw_num_id is not None else -1
            except ValueError:
                raise UnsupportedStructureError(
                    f"numbering contains a malformed numId {raw_num_id!r};"
                    " nothing was changed"
                ) from None
            if candidate_num_id == num_id:
                definition = num
                break
    if definition is None:
        raise TargetNotFoundError(
            f"numbering definition numId={num_id} does not exist; nothing to restart"
        )
    from docx.oxml.ns import nsdecls
    from docx.oxml.parser import parse_xml

    children = list(definition)
    abstract_ref = definition.find(qn("w:abstractNumId"))
    abstract_id = abstract_ref.get(qn("w:val")) if abstract_ref is not None else None
    if abstract_id is None:
        raise UnsupportedStructureError(
            f"numbering definition numId={num_id} has no abstractNumId;"
            " nothing was changed"
        )
    if len(children) != 1 or children[0] is not abstract_ref:
        raise UnsupportedStructureError(
            f"numbering definition numId={num_id} carries level overrides or"
            " other customizations that restart_numbering cannot preserve;"
            " nothing was changed"
        )
    if not any(
        abstract.get(qn("w:abstractNumId")) == abstract_id
        for abstract in numbering.findall(qn("w:abstractNum"))
    ):
        raise UnsupportedStructureError(
            f"numbering definition numId={num_id} references missing"
            f" abstractNumId={abstract_id!r}; nothing was changed"
        )
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
