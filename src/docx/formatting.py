"""Effective-format resolver — read-only, provenance-bearing.

"What formatting does this text ACTUALLY carry": document defaults →
paragraph-style inheritance chain → character style → direct formatting,
with correct toggle-property semantics. Every resolved value names the
layer it came from; what this resolver cannot resolve is declared in
`EffectiveFormat.unresolved` — "unresolved" is a legal answer, a wrong
guess is not.

Toggle properties (ISO 29500-1 §17.7.3: b, bCs, i, iCs, caps, smallCaps,
strike, emboss, imprint, outline, shadow, vanish) XOR through the style
layers: direct formatting is absolute; otherwise the effective value is
the XOR of every layer (docDefaults, then the paragraph-style chain
root→leaf, then the character-style chain root→leaf) that specifies TRUE —
the famous nested-bold-cancels gotcha, pinned by tests. Ordinary
properties take the nearest specification instead.

Declared out of scope (listed in `unresolved`): table-style
conditional formatting, numbering-mark properties, and East-Asian/
complex-script variants beyond the toggle pair (bCs/iCs).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

_VAL = qn("w:val")
_RPR = qn("w:rPr")
_PPR = qn("w:pPr")

#: the twelve OOXML run toggle properties -> our property-key names
_TOGGLES = {
    "w:b": "bold",
    "w:bCs": "bold_cs",
    "w:i": "italic",
    "w:iCs": "italic_cs",
    "w:caps": "caps",
    "w:smallCaps": "small_caps",
    "w:strike": "strike",
    "w:emboss": "emboss",
    "w:imprint": "imprint",
    "w:outline": "outline",
    "w:shadow": "shadow",
    "w:vanish": "vanish",
}

#: declared-unresolvable areas
_UNRESOLVED = (
    "table_style_conditional_formatting",
    "numbering_mark_properties",
    "east_asian_and_complex_script_variants",
    "theme_font_resolution",  # theme-referenced fonts report "theme:<token>"
)


@dataclass(frozen=True)
class ResolvedValue:
    """One property's effective value plus WHERE it came from.

    `source`: "direct", "character_style:<id>", "paragraph_style:<id>",
    "doc_defaults", "toggle_xor" (multiple toggle layers combined — the
    contributing layers are in `chain`), "mixed" (a span whose runs
    disagree), or "none" (nothing specifies it).
    """

    value: object
    source: str
    chain: Tuple[str, ...] = ()

    def to_dict(self) -> dict:
        return {"value": self.value, "source": self.source, "chain": list(self.chain)}


@dataclass(frozen=True)
class EffectiveFormat:
    """The resolved property map, plus what is declared unresolvable."""

    properties: Dict[str, ResolvedValue]
    unresolved: Tuple[str, ...]

    def __getitem__(self, key: str) -> ResolvedValue:
        return self.properties[key]

    def value_of(self, key: str):
        return self.properties[key].value

    def to_dict(self) -> dict:
        return {
            "schema": "paper_effective_format",
            "version": 1,
            "properties": {
                key: value.to_dict() for key, value in sorted(self.properties.items())
            },
            "unresolved": list(self.unresolved),
        }


def format_of(target) -> EffectiveFormat:
    """The effective formatting of a |Run|, |Paragraph| or |Span|.

    Runs resolve the full run-property set; paragraphs resolve their
    paragraph-level properties (alignment, style) plus the run defaults
    their style chain implies; spans resolve every run they touch and
    report disagreeing properties as "mixed".
    """
    from docx.search import Span
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

    if isinstance(target, Run):
        document = _document_of(target)
        # the run's parent may be a hyperlink/ins wrapper, not the paragraph
        # — walking up matters or the paragraph-style layer silently drops
        return _resolve_run(
            document, target._r, _enclosing_paragraph(target._r)  # noqa: SLF001
        )
    if isinstance(target, Paragraph):
        document = _document_of(target)
        return _resolve_paragraph(document, target._p)  # noqa: SLF001
    if isinstance(target, Span):
        return _resolve_span(target)
    raise TypeError(
        f"format_of resolves Run, Paragraph or Span targets, not"
        f" {type(target).__name__}"
    )


def surrounding_format(document: "Document", anchor) -> EffectiveFormat:
    """The effective format AT an anchor — what inserted content should
    match to adopt its neighbors' look (the insertion helper used when
    adopting a neighbor's formatting).

    Resolves the anchor paragraph's first text run; an empty paragraph
    resolves the paragraph itself.
    """
    from docx.blocks import _locate_anchor_paragraph

    _story, paragraph = _locate_anchor_paragraph(document, anchor)
    for run in paragraph.iter(qn("w:r")):  # incl. runs inside hyperlinks
        inside_textbox = False
        current = run.getparent()
        while current is not None and current is not paragraph:
            if current.tag == qn("w:txbxContent"):
                inside_textbox = True
                break
            current = current.getparent()
        if inside_textbox:
            continue
        if run.find(qn("w:t")) is not None:
            return _resolve_run(document, run, paragraph)
    return _resolve_paragraph(document, paragraph)


# ---------------------------------------------------------------------------
# resolution core
# ---------------------------------------------------------------------------


def _enclosing_paragraph(node: "_Element") -> "Optional[_Element]":
    current = node.getparent()
    while current is not None and current.tag != qn("w:p"):
        current = current.getparent()
    return current


def _document_of(proxy) -> "Document":
    part = proxy.part
    document = getattr(part, "document", None)
    if document is not None:
        return document
    return part.package.main_document_part.document


def _styles_index(document: "Document") -> "Dict[str, _Element]":
    return {
        style.get(qn("w:styleId")): style
        for style in document.styles.element.findall(qn("w:style"))
        if style.get(qn("w:styleId"))
    }


def _doc_defaults_rpr(document: "Document") -> "Optional[_Element]":
    hits = document.styles.element.xpath("w:docDefaults/w:rPrDefault/w:rPr")
    return hits[0] if hits else None


def _doc_defaults_ppr(document: "Document") -> "Optional[_Element]":
    hits = document.styles.element.xpath("w:docDefaults/w:pPrDefault/w:pPr")
    return hits[0] if hits else None


def _style_chain(
    styles: "Dict[str, _Element]", style_id: Optional[str]
) -> "List[Tuple[str, _Element]]":
    """(styleId, w:style element) root-first along w:basedOn, cycle-guarded."""
    chain: "List[Tuple[str, _Element]]" = []
    visited = set()
    current = style_id
    while current and current in styles and current not in visited:
        visited.add(current)
        chain.append((current, styles[current]))
        based_on = styles[current].find(qn("w:basedOn"))
        current = based_on.get(_VAL) if based_on is not None else None
    chain.reverse()
    return chain


def _default_style_id(document: "Document", style_type: str) -> Optional[str]:
    last = None
    for style in document.styles.element.findall(qn("w:style")):
        if style.get(qn("w:type")) == style_type and style.get(qn("w:default")) in (
            "1",
            "true",
            "on",
        ):
            last = style.get(qn("w:styleId"))
    return last


def _tri_state(container: "Optional[_Element]", tag: str) -> Optional[bool]:
    """None = not specified; True/False = the CT_OnOff value."""
    if container is None:
        return None
    node = container.find(qn(tag))
    if node is None:
        return None
    raw = node.get(_VAL)
    if raw is None:
        return True
    return raw.lower() not in ("0", "false", "off")


def _attr_of(container: "Optional[_Element]", tag: str, attr: str) -> Optional[str]:
    if container is None:
        return None
    node = container.find(qn(tag))
    return node.get(qn(attr)) if node is not None else None


def _font_name_of(r_pr: "Optional[_Element]") -> Optional[str]:
    """The ascii font name; theme references report honestly as
    "theme:<token>" rather than a guessed literal (theme_font_resolution is
    declared unresolved)."""
    literal = _attr_of(r_pr, "w:rFonts", "w:ascii")
    if literal is not None:
        return literal
    theme = _attr_of(r_pr, "w:rFonts", "w:asciiTheme")
    return f"theme:{theme}" if theme is not None else None


def _run_layers(
    document: "Document",
    run: "_Element",
    paragraph: "Optional[_Element]",
) -> "List[Tuple[str, Optional[_Element]]]":
    """(layer-name, rPr element) in APPLICATION order: docDefaults, the
    paragraph-style chain, the character-style chain, direct."""
    styles = _styles_index(document)
    layers: "List[Tuple[str, Optional[_Element]]]" = [
        ("doc_defaults", _doc_defaults_rpr(document))
    ]
    paragraph_style = None
    if paragraph is not None:
        p_pr = paragraph.find(_PPR)
        paragraph_style = _attr_of(p_pr, "w:pStyle", "w:val")
    if paragraph_style is None:
        paragraph_style = _default_style_id(document, "paragraph")
    for style_id, style in _style_chain(styles, paragraph_style):
        layers.append((f"paragraph_style:{style_id}", style.find(_RPR)))
    direct_rpr = run.find(_RPR)
    character_style = _attr_of(direct_rpr, "w:rStyle", "w:val")
    for style_id, style in _style_chain(styles, character_style):
        layers.append((f"character_style:{style_id}", style.find(_RPR)))
    layers.append(("direct", direct_rpr))
    return layers


def _resolve_toggle(
    layers: "List[Tuple[str, Optional[_Element]]]", tag: str
) -> ResolvedValue:
    direct_name, direct_rpr = layers[-1]
    direct = _tri_state(direct_rpr, tag)
    if direct is not None:
        return ResolvedValue(value=direct, source="direct", chain=("direct",))
    contributing: "List[str]" = []
    value = False
    for name, r_pr in layers[:-1]:
        specified = _tri_state(r_pr, tag)
        if specified:
            value = not value
            contributing.append(name)
    if not contributing:
        return ResolvedValue(value=False, source="none")
    if len(contributing) == 1:
        return ResolvedValue(
            value=value, source=contributing[0], chain=tuple(contributing)
        )
    return ResolvedValue(value=value, source="toggle_xor", chain=tuple(contributing))


def _resolve_nearest(
    layers: "List[Tuple[str, Optional[_Element]]]",
    reader,
) -> ResolvedValue:
    for name, r_pr in reversed(layers):
        value = reader(r_pr)
        if value is not None:
            return ResolvedValue(value=value, source=name, chain=(name,))
    return ResolvedValue(value=None, source="none")


def _resolve_run(
    document: "Document", run: "_Element", paragraph: "Optional[_Element]"
) -> EffectiveFormat:
    layers = _run_layers(document, run, paragraph)
    properties: "Dict[str, ResolvedValue]" = {}
    for tag, key in _TOGGLES.items():
        properties[key] = _resolve_toggle(layers, tag)
    properties["underline"] = _resolve_nearest(
        layers, lambda r: _attr_of(r, "w:u", "w:val")
    )
    properties["size_pt"] = _resolve_nearest(
        layers,
        lambda r: (
            int(_attr_of(r, "w:sz", "w:val")) / 2
            if _attr_of(r, "w:sz", "w:val")
            else None
        ),
    )
    properties["font_name"] = _resolve_nearest(layers, _font_name_of)
    properties["color_rgb"] = _resolve_nearest(
        layers, lambda r: _attr_of(r, "w:color", "w:val")
    )
    properties["highlight"] = _resolve_nearest(
        layers, lambda r: _attr_of(r, "w:highlight", "w:val")
    )
    properties["vertical_align"] = _resolve_nearest(
        layers, lambda r: _attr_of(r, "w:vertAlign", "w:val")
    )
    if paragraph is not None:
        properties.update(_paragraph_properties(document, paragraph))
    return EffectiveFormat(properties=properties, unresolved=_UNRESOLVED)


def _paragraph_properties(
    document: "Document", paragraph: "_Element"
) -> "Dict[str, ResolvedValue]":
    styles = _styles_index(document)
    p_pr = paragraph.find(_PPR)
    style_id = _attr_of(p_pr, "w:pStyle", "w:val") or _default_style_id(
        document, "paragraph"
    )
    layers: "List[Tuple[str, Optional[_Element]]]" = [
        ("doc_defaults", _doc_defaults_ppr(document))
    ]
    for chain_id, style in _style_chain(styles, style_id):
        layers.append((f"paragraph_style:{chain_id}", style.find(_PPR)))
    layers.append(("direct", p_pr))
    alignment = _resolve_nearest(layers, lambda p: _attr_of(p, "w:jc", "w:val"))
    style_name = None
    if style_id and style_id in styles:
        name_element = styles[style_id].find(qn("w:name"))
        raw_name = name_element.get(_VAL) if name_element is not None else style_id
        from docx.styles import BabelFish

        style_name = BabelFish.internal2ui(raw_name)  # "heading 1" -> "Heading 1"
    return {
        "alignment": alignment,
        "style_name": ResolvedValue(
            value=style_name,
            source=f"paragraph_style:{style_id}" if style_id else "none",
        ),
    }


def _resolve_paragraph(document: "Document", paragraph: "_Element") -> EffectiveFormat:
    properties = _paragraph_properties(document, paragraph)
    # the run defaults the paragraph's own chain implies (for insertions)
    probe_run = paragraph.makeelement(qn("w:r"), {})
    run_level = _resolve_run(document, probe_run, paragraph)
    for key, value in run_level.properties.items():
        properties.setdefault(key, value)
    return EffectiveFormat(properties=properties, unresolved=_UNRESOLVED)


def _resolve_span(span) -> EffectiveFormat:
    from docx.errors import UnsupportedStructureError

    span._validate_fresh()  # noqa: SLF001 - same-package machinery
    document = span._document  # noqa: SLF001
    runs = []
    for atom in span._atoms:  # noqa: SLF001
        if atom.run is None:
            raise UnsupportedStructureError(
                "the span includes content outside ordinary runs; resolve"
                " a plain text range"
            )
        if not any(atom.run is run for run, _p in runs):
            runs.append((atom.run, atom.paragraph))
    resolved = [_resolve_run(document, run, paragraph) for run, paragraph in runs]
    first = resolved[0]
    if len(resolved) == 1:
        return first
    properties: "Dict[str, ResolvedValue]" = {}
    for key, value in first.properties.items():
        if all(other.properties[key].value == value.value for other in resolved[1:]):
            sources = []
            for item in resolved:
                source = item.properties[key].source
                if source not in sources:
                    sources.append(source)
            if len(sources) == 1:
                properties[key] = value
            else:
                # same VALUE from different layers per run: asserting the
                # first run's provenance span-wide would be false provenance
                properties[key] = ResolvedValue(
                    value=value.value,
                    source="agreeing_layers",
                    chain=tuple(sources),
                )
        else:
            properties[key] = ResolvedValue(value=None, source="mixed")
    return EffectiveFormat(properties=properties, unresolved=_UNRESOLVED)
