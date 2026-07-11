"""Cross-document composition.

Copying formatted content between documents is style/numbering/relationship
reconciliation — exactly the package-level, corruption-prone mechanics this
fork exists to own. `insert_blocks_from` copies a block range from a source
document; `append_document` appends a whole source body. Both return a
|CompositionReport| declaring every part the operation may touch
plus the style/numbering/bookmark
maps and report-only findings.

Semantics: styles reconcile by
`match_by_name` (destination definition wins) or `import_renamed`
(colliding-but-different definitions clone under fresh ids/names);
numbering always REMAPS to fresh restarted definitions; images copy as new
parts with fresh rIds; external hyperlinks are recreated; bookmarks rename
on collision with REF instructions inside the range remapped. Source
revisions or comments inside the range refuse (finalize/scrub the source
first); embedded OLE objects and note references refuse (declared).
"""

from __future__ import annotations

import copy
import io
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Dict, List, Optional, Tuple

from docx._ownership import require_anchor_owner
from docx._transaction import rollback_on_error
from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import _refuse_if_protected

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

_P = qn("w:p")
_TBL = qn("w:tbl")
_SDT = qn("w:sdt")
_SECT_PR = qn("w:sectPr")
_CUSTOM_XML = qn("w:customXml")
_DATA_BINDING = qn("w:dataBinding")
_BOOKMARK_START = qn("w:bookmarkStart")
_BOOKMARK_END = qn("w:bookmarkEnd")
_ID = qn("w:id")
_NAME = qn("w:name")
_VAL = qn("w:val")
_INSTR_TEXT = qn("w:instrText")
_FLD_SIMPLE = qn("w:fldSimple")
_INSTR = qn("w:instr")
_BLIP = qn("a:blip")
# VML (legacy image markup); its prefix is not in the upstream nsmap
_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
_R_EMBED = qn("r:embed")
_R_ID = qn("r:id")
_R_LINK = qn("r:link")
_RELATIONSHIP_ATTRIBUTE_PREFIX = (
    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
)
_OFFICE_REL_ID = "{urn:schemas-microsoft-com:office:office}relid"
_HYPERLINK = qn("w:hyperlink")
_SDT_ID = qn("w:id")

_BODY_BLOCK_TAGS = frozenset((_P, _TBL, _SDT))

_STYLE_REF_TAGS = (qn("w:pStyle"), qn("w:rStyle"), qn("w:tblStyle"))
_STYLE_CHAIN_TAGS = (qn("w:basedOn"), qn("w:link"), qn("w:next"))

#: markup that refuses composition outright (declared limits)
_REFUSED_TAGS = {
    qn("w:object"): "an embedded OLE object",
    qn("w:altChunk"): "an altChunk import",
    _CUSTOM_XML: "customXml content whose backing part cannot be carried",
    _DATA_BINDING: ("a data-bound content control whose custom XML binding cannot be carried"),
    qn("w:footnoteReference"): "a footnote reference (its note cannot be carried)",
    qn("w:endnoteReference"): "an endnote reference (its note cannot be carried)",
    qn("w:commentRangeStart"): "a comment anchor (scrub the source's comments first)",
    qn("w:commentReference"): "a comment anchor (scrub the source's comments first)",
    qn("w:subDoc"): "a subdocument reference",
}


@dataclass(frozen=True)
class CompositionFinding:
    kind: str
    detail: str

    def to_dict(self) -> dict:
        return {"kind": self.kind, "detail": self.detail}


@dataclass
class CompositionReport:
    """Everything one composition call did — the declared changed-part
    budget, the reconciliation maps, and what was reported instead of done."""

    inserted_blocks: int = 0
    style_map: Dict[str, str] = field(default_factory=dict)
    imported_styles: List[str] = field(default_factory=list)
    renamed_styles: Dict[str, str] = field(default_factory=dict)
    numbering_map: Dict[int, int] = field(default_factory=dict)
    media_copied: List[str] = field(default_factory=list)
    bookmarks_renamed: Dict[str, str] = field(default_factory=dict)
    findings: List[CompositionFinding] = field(default_factory=list)
    declared_parts: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "schema": "paper_composition",
            "version": 1,
            "inserted_blocks": self.inserted_blocks,
            "style_map": dict(sorted(self.style_map.items())),
            "imported_styles": sorted(self.imported_styles),
            "renamed_styles": dict(sorted(self.renamed_styles.items())),
            "numbering_map": {str(k): v for k, v in sorted(self.numbering_map.items())},
            "media_copied": sorted(self.media_copied),
            "bookmarks_renamed": dict(sorted(self.bookmarks_renamed.items())),
            "findings": [finding.to_dict() for finding in self.findings],
            "declared_parts": sorted(set(self.declared_parts)),
        }


@dataclass(frozen=True)
class _NumberingRemap:
    source_num_id: int
    destination_num_id: int
    destination_abstract_id: int
    source_num: "_Element"
    source_abstract: "_Element"


@dataclass(frozen=True)
class _NumberingPlan:
    remaps: "Tuple[_NumberingRemap, ...]" = ()


def insert_blocks_from(
    document: "Document",
    source: "Document",
    start_anchor,
    *,
    anchor,
    end_anchor=None,
    count: int = 1,
    styles: str = "match_by_name",
) -> CompositionReport:
    """Copy a contiguous block range from `source`'s body after `anchor` in
    `document`, reconciling styles, numbering, media, and bookmarks.

    `start_anchor`/`end_anchor` address SOURCE body paragraphs (str needle,
    Block, Anchor or Span — blocks.py semantics); with no `end_anchor`,
    `count` blocks are taken (tables between the endpoints come along).
    `anchor` addresses the DESTINATION paragraph to insert after.
    """
    _validate_styles_mode(styles)
    require_anchor_owner(source, start_anchor, argument="start_anchor")
    if end_anchor is not None:
        require_anchor_owner(source, end_anchor, argument="end_anchor")
    require_anchor_owner(document, anchor)
    _refuse_if_protected(document, "compose content into the document")
    range_elements = _source_range(source, start_anchor, end_anchor, count)
    with rollback_on_error(document):
        return _compose(document, source, range_elements, anchor, styles)


def append_document(
    document: "Document",
    source: "Document",
    *,
    section: str = "new_page",
    styles: str = "match_by_name",
) -> CompositionReport:
    """Append `source`'s whole body to `document`.

    Keeps the destination's headers/footers and authors no new
    `w:sectPr`: `section="new_page"` prefixes the appended content with a
    page break; `"continuous"` appends flush. (Keeping the source's headers
    is a declared future mode.)
    """
    _validate_styles_mode(styles)
    if section not in ("new_page", "continuous"):
        raise ValueError(f"section must be 'new_page' or 'continuous', got {section!r}")
    _refuse_if_protected(document, "append a document")
    range_elements = [child for child in source.element.body if child.tag != _SECT_PR]
    if not range_elements:
        raise TargetNotFoundError("the source document body has no blocks")
    destination_blocks = [child for child in document.element.body if child.tag in _BODY_BLOCK_TAGS]
    if not destination_blocks:
        raise TargetNotFoundError("the destination document body has no blocks")
    with rollback_on_error(document):
        report = _compose(
            document,
            source,
            range_elements,
            destination_blocks[-1],
            styles,
            anchor_is_element=True,
        )
        if section == "new_page":
            break_paragraph = OxmlElement("w:p")
            run = OxmlElement("w:r")
            page_break = OxmlElement("w:br")
            page_break.set(qn("w:type"), "page")
            run.append(page_break)
            break_paragraph.append(run)
            destination_blocks[-1].addnext(break_paragraph)
        return report


def _validate_styles_mode(styles: str) -> None:
    if styles not in ("match_by_name", "import_renamed"):
        raise ValueError(f"styles must be 'match_by_name' or 'import_renamed', got {styles!r}")


def _source_range(source: "Document", start_anchor, end_anchor, count: int) -> "List[_Element]":
    from docx.blocks import _locate_anchor_paragraph

    if count < 1:
        raise ValueError("count must be >= 1")
    story, start_p = _locate_anchor_paragraph(source, start_anchor)
    if story != "word/document.xml":
        raise UnsupportedStructureError(
            f"composition copies from the main document body only (start anchor is in {story})"
        )
    body = source.element.body
    body_children = [child for child in body if child.tag != _SECT_PR]
    blocks = [child for child in body_children if child.tag in _BODY_BLOCK_TAGS]
    start_block = _body_block_for_paragraph(body, start_p)
    if start_block not in blocks:
        raise UnsupportedStructureError(
            "the start anchor is not a top-level body block (text boxes and"
            " table cells cannot anchor a composition range)"
        )
    start_index = blocks.index(start_block)
    if end_anchor is not None:
        end_story, end_p = _locate_anchor_paragraph(source, end_anchor)
        end_block = _body_block_for_paragraph(body, end_p)
        if end_story != story or end_block not in blocks:
            raise UnsupportedStructureError(
                "the end anchor must be a top-level body block of the same source document"
            )
        end_index = blocks.index(end_block)
        if end_index < start_index:
            raise TargetNotFoundError("end anchor precedes start anchor in the source body")
    else:
        end_index = start_index + count - 1
        if end_index >= len(blocks):
            raise TargetNotFoundError(
                f"the source body has only {len(blocks) - start_index} blocks"
                f" from the start anchor; {count} requested"
            )
    first = body_children.index(blocks[start_index])
    last = body_children.index(blocks[end_index])
    # Keep the physical slice. Unsupported body children between selected
    # blocks must reach preflight and refuse, never disappear by filtering.
    return body_children[first : last + 1]


def _body_block_for_paragraph(body: "_Element", paragraph: "_Element"):
    """Top-level body block containing `paragraph`, or None.

    A paragraph directly under a top-level block content control addresses
    that whole control. Paragraphs in top-level tables and other wrappers do
    not become range anchors merely because they are descendants of `body`.
    """
    current = paragraph
    while current.getparent() is not None and current.getparent() is not body:
        current = current.getparent()
    if current.getparent() is not body:
        return None
    if current.tag == _P:
        return current
    if current.tag == _SDT:
        return current
    return None


def _compose(
    document: "Document",
    source: "Document",
    range_elements: "List[_Element]",
    anchor,
    styles_mode: str,
    *,
    anchor_is_element: bool = False,
) -> CompositionReport:
    from docx.blocks import (
        _insert_after,
        _refuse_cell_anchor,
        _refuse_paragraph_in_open_field,
        _resolve_anchor_paragraph,
    )

    report = CompositionReport()
    # ALL refusal conditions run before any mutation (refusal atomicity):
    # importing styles/numbering/media first would leave orphaned
    # definitions behind when the destination anchor turns out invalid
    _refuse_unsupported_content(range_elements)
    if anchor_is_element:
        story, anchor_p = "word/document.xml", anchor
    else:
        story, anchor_p = _resolve_anchor_paragraph(document, anchor)
        if story != "word/document.xml":
            raise UnsupportedStructureError(
                f"composition inserts into the main document body only (anchor is in {story})"
            )
        _refuse_cell_anchor(anchor_p)
        root = next(r for s, r in _story_elements_of(document) if s == story)
        _refuse_paragraph_in_open_field(story, root, anchor_p, for_insertion=True)
    _refuse_malformed_numeric_ids(document, range_elements)
    _preflight_relationships(source, range_elements)
    chained_definitions = _chained_source_definitions(source, range_elements)
    numbering_plan = _preflight_numbering(document, source, range_elements + chained_definitions)
    _refuse_unloadable_media(source, range_elements)

    clones = [copy.deepcopy(element) for element in range_elements]
    imported_definitions = _reconcile_styles(document, source, clones, styles_mode, report)
    # numbering references live in imported STYLE definitions too — an
    # unmapped one silently binds to unrelated destination numbering
    _remap_numbering(
        document,
        clones + imported_definitions,
        numbering_plan,
        report,
    )
    _copy_media(document, source, clones, report)
    _recreate_hyperlinks(document, source, clones, report)
    _reconcile_bookmarks(document, clones, report)
    _reallocate_sdt_ids(document, clones)

    _pad_adjacent_tables(anchor_p, clones)
    _insert_after(anchor_p, clones)
    report.inserted_blocks = len(clones)
    report.declared_parts = [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
        *report.media_copied,
    ]
    return report


def _story_elements_of(document: "Document"):
    from docx.story import _story_elements

    return _story_elements(document)


def _chained_source_definitions(source: "Document", elements: "List[_Element]") -> "List[_Element]":
    """The source style definitions the copied range pulls in (transitive
    basedOn/link/next chains) — they carry numbering references too."""
    _root, by_id, _by_name = _style_definitions(source)
    return [
        by_id[style_id] for style_id in _expand_style_chain(by_id, _referenced_style_ids(elements))
    ]


def _preflight_relationships(source: "Document", range_elements: "List[_Element]") -> None:
    """Refuse every relationship the composition pipeline cannot remap."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    for element in range_elements:
        for node in element.iter():
            for attribute, r_id in node.attrib.items():
                if not (
                    attribute.startswith(_RELATIONSHIP_ATTRIBUTE_PREFIX)
                    or attribute == _OFFICE_REL_ID
                ):
                    continue
                rel = source.part.rels.get(r_id)
                if node.tag == _BLIP and attribute == _R_EMBED:
                    if rel is not None and rel.reltype == RT.IMAGE and not rel.is_external:
                        continue
                    _refuse_relationship(node, attribute, r_id, rel, "embedded image")
                if node.tag == _IMAGEDATA and attribute == _R_ID:
                    if rel is not None and rel.reltype == RT.IMAGE and not rel.is_external:
                        continue
                    _refuse_relationship(node, attribute, r_id, rel, "embedded image")
                if node.tag == _HYPERLINK and attribute == _R_ID:
                    if rel is not None and rel.reltype == RT.HYPERLINK and rel.is_external:
                        continue
                    _refuse_relationship(node, attribute, r_id, rel, "external hyperlink")
                expected = "linked image" if attribute == _R_LINK else None
                _refuse_relationship(node, attribute, r_id, rel, expected)


def _refuse_relationship(node, attribute, r_id, rel, expected=None) -> None:
    node_name = node.tag.rsplit("}", 1)[-1]
    attribute_name = attribute.rsplit("}", 1)[-1]
    if rel is None:
        kind = expected or "unresolved"
    else:
        kind = rel.reltype.rsplit("/", 1)[-1]
        if expected == "linked image":
            kind = expected
    raise UnsupportedStructureError(
        f"the source range contains an unsupported {kind} relationship"
        f" ({node_name}@{attribute_name}={r_id!r}); composition cannot remap"
        " it. Nothing was changed"
    )


def _refuse_unloadable_media(source: "Document", range_elements: "List[_Element]") -> None:
    """Pre-mutation check: every image in the range must be re-embeddable,
    or _copy_media would raise AFTER styles/numbering were already imported
    (refusal atomicity)."""
    from docx.image.exceptions import UnrecognizedImageError
    from docx.image.image import Image

    for element in range_elements:
        for node in element.iter(_BLIP, _IMAGEDATA):
            attr = _R_EMBED if node.tag == _BLIP else _R_ID
            r_id = node.get(attr)
            if not r_id:
                continue
            rel = source.part.rels.get(r_id)
            if rel is None or rel.is_external:
                raise UnsupportedStructureError(
                    f"image relationship {r_id!r} changed after composition"
                    " preflight. Nothing was changed"
                )
            try:
                Image.from_blob(rel.target_part.blob)
            except UnrecognizedImageError:
                raise UnsupportedStructureError(
                    "the source range contains an image format this package"
                    " cannot re-embed (e.g. EMF/WMF); convert or remove it"
                    " first. Nothing was changed"
                )


def _refuse_unsupported_content(range_elements: "List[_Element]") -> None:
    from docx.revision import _MARKUP_SCAN_TAGS

    for element in range_elements:
        if element.tag not in _BODY_BLOCK_TAGS:
            reason = _REFUSED_TAGS.get(element.tag)
            if reason is None:
                reason = f"unsupported top-level body content {element.tag.rsplit('}', 1)[-1]!r}"
            raise UnsupportedStructureError(
                f"the source range contains {reason}; composition cannot"
                " carry it (a declared limit)"
            )
        for node in element.iter():
            if node.tag in _MARKUP_SCAN_TAGS:
                raise UnsupportedStructureError(
                    "the source range carries tracked-revision markup;"
                    " finalize(revisions=...) the source first, then compose"
                )
            reason = _REFUSED_TAGS.get(node.tag)
            if reason is not None:
                raise UnsupportedStructureError(
                    f"the source range contains {reason}; composition cannot"
                    " carry it (a declared limit)"
                )


def _refuse_malformed_numeric_ids(document: "Document", range_elements: "List[_Element]") -> None:
    """Parse ids used after imports now, while refusal is still atomic."""
    roots = [root for _story, root in _story_elements_of(document)]
    roots.extend(range_elements)
    for root in roots:
        for marker in root.iter(_BOOKMARK_START, _BOOKMARK_END):
            raw = marker.get(_ID)
            try:
                value = int(raw) if raw is not None else -1
            except ValueError:
                value = -1
            if value < 0:
                raise UnsupportedStructureError(
                    "composition found a bookmark marker with a missing or"
                    f" non-numeric w:id {raw!r}; nothing was changed"
                )
        for sdt_pr in root.iter(qn("w:sdtPr")):
            id_element = sdt_pr.find(_SDT_ID)
            raw = id_element.get(_VAL) if id_element is not None else None
            if raw is None:
                continue
            try:
                value = int(raw)
            except ValueError:
                value = -1
            if value < 0:
                raise UnsupportedStructureError(
                    "composition found a content control with a non-numeric"
                    f" w:id {raw!r}; nothing was changed"
                )


# ---------------------------------------------------------------------------
# styles
# ---------------------------------------------------------------------------


def _style_definitions(
    document: "Document",
) -> "Tuple[_Element, Dict[str, _Element], Dict[str, _Element]]":
    root = document.styles.element
    by_id: "Dict[str, _Element]" = {}
    by_name: "Dict[str, _Element]" = {}
    for style in root.findall(qn("w:style")):
        style_id = style.get(qn("w:styleId"))
        name_element = style.find(qn("w:name"))
        name = name_element.get(_VAL) if name_element is not None else None
        if style_id:
            by_id[style_id] = style
        if name:
            by_name[name] = style
    return root, by_id, by_name


def _referenced_style_ids(clones: "List[_Element]") -> "List[str]":
    seen: "List[str]" = []
    for clone in clones:
        for tag in _STYLE_REF_TAGS:
            for node in clone.iter(tag):
                value = node.get(_VAL)
                if value and value not in seen:
                    seen.append(value)
    return seen


def _expand_style_chain(source_by_id: "Dict[str, _Element]", wanted: "List[str]") -> "List[str]":
    ordered: "List[str]" = []
    queue = [(style_id, None, None) for style_id in wanted]
    while queue:
        style_id, parent_style_id, chain_tag = queue.pop(0)
        if style_id in ordered:
            continue
        if style_id not in source_by_id:
            owner = (
                "the source content"
                if parent_style_id is None
                else f"source style {parent_style_id!r} through w:{chain_tag}"
            )
            raise UnsupportedStructureError(
                f"{owner} references undefined source style {style_id!r}; nothing was changed"
            )
        ordered.append(style_id)
        for tag in _STYLE_CHAIN_TAGS:
            for chained in source_by_id[style_id].findall(tag):
                chained_style_id = chained.get(_VAL)
                if chained_style_id:
                    queue.append((chained_style_id, style_id, tag.rsplit("}", 1)[-1]))
    return ordered


def _normalized_style_bytes(style: "_Element") -> bytes:
    from lxml import etree

    clone = copy.deepcopy(style)
    for attr in list(clone.attrib):
        if attr.endswith("}styleId") or "rsid" in attr:
            del clone.attrib[attr]
    return etree.tostring(clone)


def _reconcile_styles(
    document: "Document",
    source: "Document",
    clones: "List[_Element]",
    mode: str,
    report: CompositionReport,
) -> "List[_Element]":
    """Reconcile and remap; returns the IMPORTED definitions (their
    numbering references still need remapping by the caller)."""
    destination_root, destination_by_id, destination_by_name = _style_definitions(document)
    _root, source_by_id, _source_by_name = _style_definitions(source)
    wanted = _expand_style_chain(source_by_id, _referenced_style_ids(clones))
    style_map: "Dict[str, str]" = {}
    to_import: "List[Tuple[str, _Element]]" = []
    taken_ids = set(destination_by_id)  # incl. ids allocated THIS batch
    for style_id in wanted:
        definition = source_by_id[style_id]
        name_element = definition.find(qn("w:name"))
        name = name_element.get(_VAL) if name_element is not None else style_id
        existing = destination_by_name.get(name)
        if existing is not None:
            if mode == "match_by_name":
                style_map[style_id] = existing.get(qn("w:styleId"))
                continue
            if _normalized_style_bytes(existing) == _normalized_style_bytes(definition):
                style_map[style_id] = existing.get(qn("w:styleId"))
                continue
            # import_renamed: clone under a fresh id AND name
            new_id = _fresh_style_id(taken_ids, style_id)
            taken_ids.add(new_id)
            new_name = _fresh_style_name(destination_by_name, name)
            style_map[style_id] = new_id
            report.renamed_styles[name] = new_name
            to_import.append((new_id, _renamed_clone(definition, new_id, new_name)))
            continue
        new_id = style_id if style_id not in taken_ids else _fresh_style_id(taken_ids, style_id)
        taken_ids.add(new_id)
        style_map[style_id] = new_id
        to_import.append((new_id, _renamed_clone(definition, new_id, None)))
    for new_id, definition in to_import:
        # remap chain references to their final destination ids
        for tag in _STYLE_CHAIN_TAGS:
            for chained in definition.findall(tag):
                if chained.get(_VAL) in style_map:
                    chained.set(_VAL, style_map[chained.get(_VAL)])
                else:
                    definition.remove(chained)  # empty/malformed chain link
        destination_root.append(definition)
        report.imported_styles.append(new_id)
        destination_by_id[new_id] = definition
    for clone in clones:
        for tag in _STYLE_REF_TAGS:
            for node in clone.iter(tag):
                value = node.get(_VAL)
                if value in style_map and style_map[value] != value:
                    node.set(_VAL, style_map[value])
    report.style_map = style_map
    return [definition for _new_id, definition in to_import]


def _renamed_clone(definition: "_Element", new_id: str, new_name: Optional[str]) -> "_Element":
    clone = copy.deepcopy(definition)
    clone.set(qn("w:styleId"), new_id)
    if new_name is not None:
        name_element = clone.find(qn("w:name"))
        if name_element is not None:
            name_element.set(_VAL, new_name)
    return clone


def _fresh_style_id(taken_ids, base: str) -> str:
    candidate = f"{base}Imported"
    counter = 1
    while candidate in taken_ids:
        counter += 1
        candidate = f"{base}Imported{counter}"
    return candidate


def _fresh_style_name(destination_by_name: "Dict[str, _Element]", base: str) -> str:
    candidate = f"{base} (imported)"
    counter = 1
    while candidate in destination_by_name:
        counter += 1
        candidate = f"{base} (imported {counter})"
    return candidate


# ---------------------------------------------------------------------------
# numbering
# ---------------------------------------------------------------------------


def _numbering_root_of(document: "Document") -> "Optional[_Element]":
    from docx.numbering import _numbering_root

    return _numbering_root(document)


def _preflight_numbering(
    document: "Document",
    source: "Document",
    elements: "List[_Element]",
) -> _NumberingPlan:
    """Validate both numbering graphs and allocate every remap up front."""
    referenced = _referenced_numbering_ids(elements)
    if not referenced:
        return _NumberingPlan()

    source_root = _numbering_root_of(source)
    if source_root is None:
        raise UnsupportedStructureError(
            "the source content references numbering but the source document"
            " has no numbering part. Nothing was changed"
        )
    destination_root = _numbering_root_of(document)
    if destination_root is None:
        raise UnsupportedStructureError(
            "the destination document has no numbering part; create a list"
            " definition first (ensure_bullet_definition /"
            " ensure_decimal_definition)"
        )

    source_nums, source_abstracts, source_abstract_refs = _validated_numbering_graph(
        source_root, "source numbering"
    )
    destination_nums, destination_abstracts, _destination_abstract_refs = (
        _validated_numbering_graph(destination_root, "destination numbering")
    )

    next_num_id = max(destination_nums, default=0) + 1
    next_abstract_id = max(destination_abstracts, default=-1) + 1
    remaps = []
    for source_num_id in referenced:
        source_num = source_nums.get(source_num_id)
        if source_num is None:
            raise UnsupportedStructureError(
                f"source numbering id {source_num_id} has no w:num definition. Nothing was changed"
            )
        source_abstract_id = source_abstract_refs[source_num_id]
        source_abstract = source_abstracts[source_abstract_id]
        if any(source_abstract.iter(qn("w:lvlPicBulletId"))):
            raise UnsupportedStructureError(
                f"source numbering id {source_num_id} uses a picture bullet;"
                " composition cannot carry its numbering-part relationships."
                " Nothing was changed"
            )
        remaps.append(
            _NumberingRemap(
                source_num_id=source_num_id,
                destination_num_id=next_num_id,
                destination_abstract_id=next_abstract_id,
                source_num=source_num,
                source_abstract=source_abstract,
            )
        )
        next_num_id += 1
        next_abstract_id += 1
    return _NumberingPlan(tuple(remaps))


def _referenced_numbering_ids(elements: "List[_Element]") -> "List[int]":
    referenced = []
    for element in elements:
        for num_id_element in element.iter(qn("w:numId")):
            num_id = _parse_numbering_id(
                num_id_element.get(_VAL),
                "source content has a w:numId with",
            )
            if num_id > 0 and num_id not in referenced:
                referenced.append(num_id)
    return referenced


def _validated_numbering_graph(
    root: "_Element", label: str
) -> "Tuple[Dict[int, _Element], Dict[int, _Element], Dict[int, int]]":
    picture_ids = _index_numbering_elements(
        root.findall(qn("w:numPicBullet")),
        qn("w:numPicBulletId"),
        label,
        "w:numPicBullet",
    )
    abstracts = _index_numbering_elements(
        root.findall(qn("w:abstractNum")),
        qn("w:abstractNumId"),
        label,
        "w:abstractNum",
    )
    nums = _index_numbering_elements(
        root.findall(qn("w:num")),
        qn("w:numId"),
        label,
        "w:num",
    )

    abstract_refs: "Dict[int, int]" = {}
    for num_id, num in nums.items():
        refs = num.findall(qn("w:abstractNumId"))
        if len(refs) != 1:
            raise UnsupportedStructureError(
                f"{label} w:num {num_id} must contain exactly one"
                " w:abstractNumId. Nothing was changed"
            )
        abstract_id = _parse_numbering_id(
            refs[0].get(_VAL),
            f"{label} w:num {num_id} has a w:abstractNumId with",
        )
        if abstract_id not in abstracts:
            raise UnsupportedStructureError(
                f"{label} w:num {num_id} references missing abstract"
                f" numbering definition {abstract_id}. Nothing was changed"
            )
        abstract_refs[num_id] = abstract_id

    for abstract_id, abstract in abstracts.items():
        for picture_ref in abstract.iter(qn("w:lvlPicBulletId")):
            picture_id = _parse_numbering_id(
                picture_ref.get(_VAL),
                f"{label} w:abstractNum {abstract_id} has a w:lvlPicBulletId with",
            )
            if picture_id not in picture_ids:
                raise UnsupportedStructureError(
                    f"{label} w:abstractNum {abstract_id} references missing"
                    f" picture-bullet definition {picture_id}. Nothing was"
                    " changed"
                )
    return nums, abstracts, abstract_refs


def _index_numbering_elements(
    elements, attribute, label: str, element_name: str
) -> "Dict[int, _Element]":
    indexed = {}
    for element in elements:
        value = _parse_numbering_id(
            element.get(attribute),
            f"{label} has a {element_name} with",
        )
        if value in indexed:
            raise UnsupportedStructureError(
                f"{label} contains duplicate {element_name} id {value}. Nothing was changed"
            )
        indexed[value] = element
    return indexed


def _parse_numbering_id(raw: Optional[str], prefix: str) -> int:
    try:
        value = int(raw) if raw is not None else -1
    except ValueError:
        value = -1
    if value < 0:
        raise UnsupportedStructureError(
            f"{prefix} missing or non-numeric id {raw!r}. Nothing was changed"
        )
    return value


def _remap_numbering(
    document: "Document",
    clones: "List[_Element]",
    plan: _NumberingPlan,
    report: CompositionReport,
) -> None:
    referenced = []
    for clone in clones:
        for num_id_element in clone.iter(qn("w:numId")):
            value = num_id_element.get(_VAL)
            if value and int(value) > 0 and int(value) not in referenced:
                referenced.append(int(value))
    if not referenced:
        return
    destination_root = _numbering_root_of(document)
    numbering_map: "Dict[int, int]" = {}
    for remap in plan.remaps:
        if remap.source_num_id not in referenced:
            continue
        abstract_clone = copy.deepcopy(remap.source_abstract)
        abstract_clone.set(qn("w:abstractNumId"), str(remap.destination_abstract_id))
        # nsid/tmpl uniqueness is advisory; leaving them is Word-tolerated
        first_num = destination_root.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(abstract_clone)
        else:
            destination_root.append(abstract_clone)
        num_clone = copy.deepcopy(remap.source_num)
        num_clone.set(qn("w:numId"), str(remap.destination_num_id))
        new_ref = num_clone.find(qn("w:abstractNumId"))
        new_ref.set(_VAL, str(remap.destination_abstract_id))
        destination_root.append(num_clone)
        numbering_map[remap.source_num_id] = remap.destination_num_id
    for clone in clones:
        for num_id_element in clone.iter(qn("w:numId")):
            value = num_id_element.get(_VAL)
            if value and int(value) in numbering_map:
                num_id_element.set(_VAL, str(numbering_map[int(value)]))
    report.numbering_map = numbering_map


# ---------------------------------------------------------------------------
# media and hyperlinks
# ---------------------------------------------------------------------------


def _copy_media(
    document: "Document",
    source: "Document",
    clones: "List[_Element]",
    report: CompositionReport,
) -> None:
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    rel_map: "Dict[str, str]" = {}
    for clone in clones:
        for node in clone.iter(_BLIP, _IMAGEDATA):
            attr = _R_EMBED if node.tag == _BLIP else _R_ID
            r_id = node.get(attr)
            if not r_id:
                continue
            if r_id not in rel_map:
                rel = source.part.rels.get(r_id)
                if rel is None or rel.is_external or rel.reltype != RT.IMAGE:
                    raise UnsupportedStructureError(
                        f"image relationship {r_id!r} changed after composition"
                        " preflight. Nothing was changed"
                    )
                blob = rel.target_part.blob
                new_r_id, _image = document.part.get_or_add_image(io.BytesIO(blob))
                rel_map[r_id] = new_r_id
                report.media_copied.append(
                    str(document.part.rels[new_r_id].target_part.partname).lstrip("/")
                )
            node.set(attr, rel_map[r_id])


def _recreate_hyperlinks(
    document: "Document",
    source: "Document",
    clones: "List[_Element]",
    report: CompositionReport,
) -> None:
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    for clone in clones:
        for hyperlink in clone.iter(_HYPERLINK):
            r_id = hyperlink.get(_R_ID)
            if not r_id:
                continue  # internal anchor link: carried as-is
            rel = source.part.rels.get(r_id)
            if rel is None or not rel.is_external or rel.reltype != RT.HYPERLINK:
                raise UnsupportedStructureError(
                    f"hyperlink relationship {r_id!r} changed after composition"
                    " preflight. Nothing was changed"
                )
            new_r_id = document.part.relate_to(rel.target_ref, RT.HYPERLINK, is_external=True)
            hyperlink.set(_R_ID, new_r_id)


# ---------------------------------------------------------------------------
# bookmarks and control ids
# ---------------------------------------------------------------------------


def _reconcile_bookmarks(
    document: "Document", clones: "List[_Element]", report: CompositionReport
) -> None:
    from docx.story import _story_elements

    existing_names = set()
    max_id = 0
    for _story, root in _story_elements(document):
        for start in root.iter(_BOOKMARK_START):
            if start.get(_NAME):
                existing_names.add(start.get(_NAME))
            max_id = max(max_id, int(start.get(_ID) or 0))
        for end in root.iter(_BOOKMARK_END):
            max_id = max(max_id, int(end.get(_ID) or 0))
    renames: "Dict[str, str]" = {}
    id_map: "Dict[str, str]" = {}
    next_id = max_id + 1
    for clone in clones:
        for start in list(clone.iter(_BOOKMARK_START)):
            name = start.get(_NAME) or ""
            if name == "_GoBack":  # Word cursor noise; never carried
                old_id = start.get(_ID)
                for end in clone.iter(_BOOKMARK_END):
                    if end.get(_ID) == old_id:
                        end.getparent().remove(end)
                        break
                start.getparent().remove(start)
                continue
            old_id = start.get(_ID)
            id_map[old_id] = str(next_id)
            start.set(_ID, str(next_id))
            next_id += 1
            if name in existing_names:
                new_name = _fresh_bookmark_name(existing_names, name)
                renames[name] = new_name
                start.set(_NAME, new_name)
                existing_names.add(new_name)
            else:
                existing_names.add(name)
        for end in list(clone.iter(_BOOKMARK_END)):
            old_id = end.get(_ID)
            if old_id in id_map:
                end.set(_ID, id_map[old_id])
            else:
                # its start lies OUTSIDE the copied range: keeping the source
                # id would terminate an unrelated destination bookmark
                end.getparent().remove(end)
                report.findings.append(
                    CompositionFinding(
                        kind="bookmark_partially_in_range",
                        detail=(
                            "a bookmark end whose start lies outside the copied range was dropped"
                        ),
                    )
                )
    # starts whose end never appeared in the range are half-pairs too
    ends_present = {end.get(_ID) for clone in clones for end in clone.iter(_BOOKMARK_END)}
    for clone in clones:
        for start in list(clone.iter(_BOOKMARK_START)):
            if start.get(_ID) not in ends_present:
                name = start.get(_NAME) or ""
                start.getparent().remove(start)
                renames.pop(name, None)
                report.findings.append(
                    CompositionFinding(
                        kind="bookmark_partially_in_range",
                        detail=(
                            f"bookmark {name!r} starts in the copied range"
                            " but ends outside it; the start marker was"
                            " dropped"
                        ),
                    )
                )
    if renames:
        _remap_field_refs(clones, renames)
    report.bookmarks_renamed = renames


def _fresh_bookmark_name(existing: set, base: str) -> str:
    # Imported names are authored by this package, so they must obey Word's
    # public bookmark grammar even when the source name did not.
    stem = "".join(char if char == "_" or char.isalnum() else "_" for char in base)
    if not stem or not stem[0].isalpha():
        stem = f"B_{stem}"
    counter = 1
    while True:
        suffix = "_imported" if counter == 1 else f"_imported{counter}"
        candidate = f"{stem[: 40 - len(suffix)]}{suffix}"
        if candidate not in existing:
            return candidate
        counter += 1


def _remap_field_refs(clones: "List[_Element]", renames: "Dict[str, str]") -> None:
    """Rewrite REF/PAGEREF/NOTEREF instructions and hyperlink anchors inside
    the copied range that point at renamed bookmarks.

    All renames apply SIMULTANEOUSLY (one alternation pass — sequential
    substitution chains A->B then B->C), and complex-field instructions are
    matched on their CONCATENATION across split w:instrText runs."""
    from docx.bookmarks import _iter_field_instructions

    lookup = {old.casefold(): new for old, new in renames.items()}
    alternatives = "|".join(re.escape(old) for old in sorted(renames, key=len, reverse=True))
    reference_operand = re.compile(
        rf"(?P<prefix>\b(?:PAGEREF|NOTEREF|REF)\s+)"
        rf'(?:"(?P<quoted>{alternatives})"|(?P<bare>{alternatives}))'
        r"(?=\s|$)",
        re.IGNORECASE,
    )

    def rewrite(text: str) -> str:
        def replacement(match) -> str:
            old = match.group("quoted") or match.group("bare")
            new = lookup[old.casefold()]
            operand = f'"{new}"' if match.group("quoted") is not None else new
            return f"{match.group('prefix')}{operand}"

        return reference_operand.sub(replacement, text)

    def redistribute(nodes: "List[_Element]", instruction: str) -> None:
        """Write a rewritten concatenation back across its existing nodes."""
        offset = 0
        for position, node in enumerate(nodes):
            if position == len(nodes) - 1:
                node.text = instruction[offset:]
                break
            width = len(node.text or "")
            node.text = instruction[offset : offset + width]
            offset += width

    for clone in clones:
        for instruction, nodes in _iter_field_instructions(clone):
            rewritten = rewrite(instruction)
            if rewritten != instruction:
                redistribute(nodes, rewritten)
        for node in clone.iter(_FLD_SIMPLE):
            instr = node.get(_INSTR)
            if instr:
                rewritten = rewrite(instr)
                if rewritten != instr:
                    node.set(_INSTR, rewritten)
        for link in clone.iter(_HYPERLINK):
            anchor = link.get(qn("w:anchor"))
            if anchor and anchor.casefold() in lookup:
                link.set(qn("w:anchor"), lookup[anchor.casefold()])


def _reallocate_sdt_ids(document: "Document", clones: "List[_Element]") -> None:
    from docx.story import _story_elements

    existing = set()
    for _story, root in _story_elements(document):
        for sdt_pr in root.iter(qn("w:sdtPr")):
            id_element = sdt_pr.find(_SDT_ID)
            if id_element is not None and id_element.get(_VAL):
                existing.add(int(id_element.get(_VAL)))
    next_id = max(existing, default=0) + 1
    for clone in clones:
        for sdt_pr in clone.iter(qn("w:sdtPr")):
            id_element = sdt_pr.find(_SDT_ID)
            if id_element is None:
                continue
            value = id_element.get(_VAL)
            if value and int(value) in existing:
                id_element.set(_VAL, str(next_id))
                existing.add(next_id)
                next_id += 1
            elif value:
                kept_id = int(value)
                existing.add(kept_id)
                next_id = max(next_id, kept_id + 1)


def _pad_adjacent_tables(anchor_p: "_Element", clones: "List[_Element]") -> None:
    """Two adjacent tables fuse visually in Word; keep an empty paragraph
    between a cloned table and a neighboring destination table."""
    if clones and clones[0].tag == _TBL and anchor_p.tag == _TBL:
        clones.insert(0, OxmlElement("w:p"))
    if clones and clones[-1].tag == _TBL:
        following = anchor_p.getnext()
        if following is not None and following.tag == _TBL:
            clones.append(OxmlElement("w:p"))
