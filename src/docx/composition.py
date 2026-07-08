"""Cross-document composition (v0.11 Phase 5, API-PROPOSAL §11).

Copying formatted content between documents is style/numbering/relationship
reconciliation — exactly the package-level, corruption-prone mechanics this
fork exists to own. `insert_blocks_from` copies a block range from a source
document; `append_document` appends a whole source body. Both return a
|CompositionReport| declaring every part the operation may touch
(report-matches-diff, never small-diff) plus the style/numbering/bookmark
maps and report-only findings.

Pinned semantics (see the proposal): styles reconcile by
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

from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import _refuse_if_protected

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

_P = qn("w:p")
_TBL = qn("w:tbl")
_SECT_PR = qn("w:sectPr")
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
_HYPERLINK = qn("w:hyperlink")
_SDT_ID = qn("w:id")

_STYLE_REF_TAGS = (qn("w:pStyle"), qn("w:rStyle"), qn("w:tblStyle"))
_STYLE_CHAIN_TAGS = (qn("w:basedOn"), qn("w:link"), qn("w:next"))

#: markup that refuses composition outright (declared limits)
_REFUSED_TAGS = {
    qn("w:object"): "an embedded OLE object",
    qn("w:altChunk"): "an altChunk import",
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
            "numbering_map": {
                str(k): v for k, v in sorted(self.numbering_map.items())
            },
            "media_copied": sorted(self.media_copied),
            "bookmarks_renamed": dict(sorted(self.bookmarks_renamed.items())),
            "findings": [finding.to_dict() for finding in self.findings],
            "declared_parts": sorted(set(self.declared_parts)),
        }


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
    _refuse_if_protected(document, "compose content into the document")
    range_elements = _source_range(source, start_anchor, end_anchor, count)
    return _compose(document, source, range_elements, anchor, styles)


def append_document(
    document: "Document",
    source: "Document",
    *,
    section: str = "new_page",
    styles: str = "match_by_name",
) -> CompositionReport:
    """Append `source`'s whole body to `document`.

    v0.11 keeps the destination's headers/footers and authors no new
    `w:sectPr`: `section="new_page"` prefixes the appended content with a
    page break; `"continuous"` appends flush. (Keeping the source's headers
    is a declared future mode.)
    """
    _validate_styles_mode(styles)
    if section not in ("new_page", "continuous"):
        raise ValueError(
            f"section must be 'new_page' or 'continuous', got {section!r}"
        )
    _refuse_if_protected(document, "append a document")
    range_elements = [
        child
        for child in source.element.body
        if child.tag in (_P, _TBL)
    ]
    if not range_elements:
        raise TargetNotFoundError("the source document body has no blocks")
    destination_blocks = [
        child for child in document.element.body if child.tag in (_P, _TBL)
    ]
    if not destination_blocks:
        raise TargetNotFoundError("the destination document body has no blocks")
    report = _compose(
        document, source, range_elements, destination_blocks[-1], styles,
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
        raise ValueError(
            f"styles must be 'match_by_name' or 'import_renamed', got {styles!r}"
        )


def _source_range(
    source: "Document", start_anchor, end_anchor, count: int
) -> "List[_Element]":
    from docx.blocks import _locate_anchor_paragraph

    if count < 1:
        raise ValueError("count must be >= 1")
    story, start_p = _locate_anchor_paragraph(source, start_anchor)
    if story != "word/document.xml":
        raise UnsupportedStructureError(
            f"composition copies from the main document body only"
            f" (start anchor is in {story})"
        )
    body = source.element.body
    blocks = [child for child in body if child.tag in (_P, _TBL)]
    if start_p not in blocks:
        raise UnsupportedStructureError(
            "the start anchor is not a top-level body block (text boxes and"
            " table cells cannot anchor a composition range)"
        )
    start_index = blocks.index(start_p)
    if end_anchor is not None:
        end_story, end_p = _locate_anchor_paragraph(source, end_anchor)
        if end_story != story or end_p not in blocks:
            raise UnsupportedStructureError(
                "the end anchor must be a top-level body block of the same"
                " source document"
            )
        end_index = blocks.index(end_p)
        if end_index < start_index:
            raise TargetNotFoundError(
                "end anchor precedes start anchor in the source body"
            )
    else:
        end_index = start_index + count - 1
        if end_index >= len(blocks):
            raise TargetNotFoundError(
                f"the source body has only {len(blocks) - start_index} blocks"
                f" from the start anchor; {count} requested"
            )
    return blocks[start_index : end_index + 1]


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
                f"composition inserts into the main document body only"
                f" (anchor is in {story})"
            )
        _refuse_cell_anchor(anchor_p)
        root = next(
            r for s, r in _story_elements_of(document) if s == story
        )
        _refuse_paragraph_in_open_field(story, root, anchor_p, for_insertion=True)
    _refuse_missing_numbering_part(document, range_elements)

    clones = [copy.deepcopy(element) for element in range_elements]
    _reconcile_styles(document, source, clones, styles_mode, report)
    _remap_numbering(document, source, clones, report)
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


def _refuse_missing_numbering_part(
    document: "Document", range_elements: "List[_Element]"
) -> None:
    """Pre-mutation check for the numbering remap's only failure mode."""
    needs_numbering = any(
        int(node.get(_VAL) or 0) > 0
        for element in range_elements
        for node in element.iter(qn("w:numId"))
    )
    if needs_numbering and _numbering_root_of(document) is None:
        raise UnsupportedStructureError(
            "the destination document has no numbering part; create a list"
            " definition first (ensure_bullet_definition /"
            " ensure_decimal_definition)"
        )


def _refuse_unsupported_content(range_elements: "List[_Element]") -> None:
    from docx.revision import _MARKUP_SCAN_TAGS

    for element in range_elements:
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
                    " carry it in v0.11 (declared limit)"
                )


# ---------------------------------------------------------------------------
# styles
# ---------------------------------------------------------------------------


def _style_definitions(document: "Document") -> "Tuple[_Element, Dict[str, _Element], Dict[str, _Element]]":
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


def _expand_style_chain(
    source_by_id: "Dict[str, _Element]", wanted: "List[str]"
) -> "List[str]":
    ordered: "List[str]" = []
    queue = list(wanted)
    while queue:
        style_id = queue.pop(0)
        if style_id in ordered or style_id not in source_by_id:
            continue
        ordered.append(style_id)
        for tag in _STYLE_CHAIN_TAGS:
            chained = source_by_id[style_id].find(tag)
            if chained is not None and chained.get(_VAL):
                queue.append(chained.get(_VAL))
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
) -> None:
    destination_root, destination_by_id, destination_by_name = _style_definitions(
        document
    )
    _root, source_by_id, _source_by_name = _style_definitions(source)
    wanted = _expand_style_chain(source_by_id, _referenced_style_ids(clones))
    style_map: "Dict[str, str]" = {}
    to_import: "List[Tuple[str, _Element]]" = []
    for style_id in wanted:
        definition = source_by_id[style_id]
        name_element = definition.find(qn("w:name"))
        name = name_element.get(_VAL) if name_element is not None else style_id
        existing = destination_by_name.get(name)
        if existing is not None:
            if mode == "match_by_name":
                style_map[style_id] = existing.get(qn("w:styleId"))
                continue
            if _normalized_style_bytes(existing) == _normalized_style_bytes(
                definition
            ):
                style_map[style_id] = existing.get(qn("w:styleId"))
                continue
            # import_renamed: clone under a fresh id AND name
            new_id = _fresh_style_id(destination_by_id, style_id)
            new_name = _fresh_style_name(destination_by_name, name)
            style_map[style_id] = new_id
            report.renamed_styles[name] = new_name
            to_import.append((new_id, _renamed_clone(definition, new_id, new_name)))
            continue
        new_id = (
            style_id
            if style_id not in destination_by_id
            else _fresh_style_id(destination_by_id, style_id)
        )
        style_map[style_id] = new_id
        to_import.append((new_id, _renamed_clone(definition, new_id, None)))
    for new_id, definition in to_import:
        # remap chain references to their final destination ids
        for tag in _STYLE_CHAIN_TAGS:
            chained = definition.find(tag)
            if chained is not None and chained.get(_VAL) in style_map:
                chained.set(_VAL, style_map[chained.get(_VAL)])
            elif chained is not None:
                definition.remove(chained)  # dangling chain link
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


def _renamed_clone(
    definition: "_Element", new_id: str, new_name: Optional[str]
) -> "_Element":
    clone = copy.deepcopy(definition)
    clone.set(qn("w:styleId"), new_id)
    if new_name is not None:
        name_element = clone.find(qn("w:name"))
        if name_element is not None:
            name_element.set(_VAL, new_name)
    return clone


def _fresh_style_id(destination_by_id: "Dict[str, _Element]", base: str) -> str:
    candidate = f"{base}Imported"
    counter = 1
    while candidate in destination_by_id:
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


def _remap_numbering(
    document: "Document",
    source: "Document",
    clones: "List[_Element]",
    report: CompositionReport,
) -> None:
    referenced: "List[int]" = []
    for clone in clones:
        for num_id_element in clone.iter(qn("w:numId")):
            value = num_id_element.get(_VAL)
            if value and int(value) > 0 and int(value) not in referenced:
                referenced.append(int(value))
    if not referenced:
        return
    source_root = _numbering_root_of(source)
    destination_root = _numbering_root_of(document)
    numbering_map: "Dict[int, int]" = {}
    for num_id in referenced:
        source_num = _find_num(source_root, num_id)
        if source_num is None:
            for clone in clones:
                _strip_num_refs(clone, num_id)
            report.findings.append(
                CompositionFinding(
                    kind="numbering_reference_stripped",
                    detail=(
                        f"source numbering id {num_id} has no definition;"
                        " the reference was stripped"
                    ),
                )
            )
            continue
        abstract_ref = source_num.find(qn("w:abstractNumId"))
        source_abstract = (
            _find_abstract(source_root, abstract_ref.get(_VAL))
            if abstract_ref is not None
            else None
        )
        new_num_id = _max_attr(destination_root, qn("w:num"), qn("w:numId")) + 1
        new_abstract_id = (
            _max_attr(destination_root, qn("w:abstractNum"), qn("w:abstractNumId")) + 1
        )
        if source_abstract is not None:
            abstract_clone = copy.deepcopy(source_abstract)
            abstract_clone.set(qn("w:abstractNumId"), str(new_abstract_id))
            # nsid/tmpl uniqueness is advisory; leaving them is Word-tolerated
            first_num = destination_root.find(qn("w:num"))
            if first_num is not None:
                first_num.addprevious(abstract_clone)
            else:
                destination_root.append(abstract_clone)
        num_clone = copy.deepcopy(source_num)
        num_clone.set(qn("w:numId"), str(new_num_id))
        new_ref = num_clone.find(qn("w:abstractNumId"))
        if new_ref is not None:
            new_ref.set(_VAL, str(new_abstract_id))
        destination_root.append(num_clone)
        numbering_map[num_id] = new_num_id
    for clone in clones:
        for num_id_element in clone.iter(qn("w:numId")):
            value = num_id_element.get(_VAL)
            if value and int(value) in numbering_map:
                num_id_element.set(_VAL, str(numbering_map[int(value)]))
    report.numbering_map = numbering_map


def _find_num(root: "Optional[_Element]", num_id: int) -> "Optional[_Element]":
    if root is None:
        return None
    for num in root.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(num_id):
            return num
    return None


def _find_abstract(root: "_Element", abstract_id: Optional[str]) -> "Optional[_Element]":
    if abstract_id is None:
        return None
    for abstract in root.findall(qn("w:abstractNum")):
        if abstract.get(qn("w:abstractNumId")) == abstract_id:
            return abstract
    return None


def _max_attr(root: "_Element", element_tag, attr) -> int:
    values = [
        int(element.get(attr))
        for element in root.findall(element_tag)
        if element.get(attr) is not None
    ]
    return max(values, default=0)


def _strip_num_refs(clone: "_Element", num_id: int) -> None:
    for num_pr in list(clone.iter(qn("w:numPr"))):
        num_id_element = num_pr.find(qn("w:numId"))
        if num_id_element is not None and num_id_element.get(_VAL) == str(num_id):
            num_pr.getparent().remove(num_pr)


# ---------------------------------------------------------------------------
# media and hyperlinks
# ---------------------------------------------------------------------------


def _copy_media(
    document: "Document",
    source: "Document",
    clones: "List[_Element]",
    report: CompositionReport,
) -> None:
    rel_map: "Dict[str, str]" = {}
    for clone in clones:
        for node in list(clone.iter(_BLIP, _IMAGEDATA)):
            attr = _R_EMBED if node.tag == _BLIP else _R_ID
            r_id = node.get(attr)
            if not r_id:
                continue
            if r_id not in rel_map:
                rel = source.part.rels.get(r_id)
                if rel is None or rel.is_external:
                    _remove_holder_run(node)
                    report.findings.append(
                        CompositionFinding(
                            kind="image_reference_stripped",
                            detail=(
                                f"image relationship {r_id} could not be"
                                " resolved in the source; the image was"
                                " dropped"
                            ),
                        )
                    )
                    continue
                blob = rel.target_part.blob
                new_r_id, _image = document.part.get_or_add_image(io.BytesIO(blob))
                rel_map[r_id] = new_r_id
                report.media_copied.append(
                    str(document.part.rels[new_r_id].target_part.partname).lstrip("/")
                )
            node.set(attr, rel_map[r_id])


def _remove_holder_run(node: "_Element") -> None:
    holder = node
    while holder.getparent() is not None and holder.tag != qn("w:r"):
        holder = holder.getparent()
    if holder.getparent() is not None:
        holder.getparent().remove(holder)


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
            if rel is None or not rel.is_external:
                # flatten a dangling/internal-target link to its text
                parent = hyperlink.getparent()
                for child in list(hyperlink):
                    hyperlink.addprevious(child)
                parent.remove(hyperlink)
                report.findings.append(
                    CompositionFinding(
                        kind="hyperlink_flattened",
                        detail=(
                            f"hyperlink relationship {r_id} could not be"
                            " recreated; the link was flattened to text"
                        ),
                    )
                )
                continue
            new_r_id = document.part.relate_to(
                rel.target_ref, RT.HYPERLINK, is_external=True
            )
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
        for end in clone.iter(_BOOKMARK_END):
            old_id = end.get(_ID)
            if old_id in id_map:
                end.set(_ID, id_map[old_id])
    if renames:
        _remap_field_refs(clones, renames)
    report.bookmarks_renamed = renames


def _fresh_bookmark_name(existing: set, base: str) -> str:
    candidate = f"{base}_imported"
    counter = 1
    while candidate in existing:
        counter += 1
        candidate = f"{base}_imported{counter}"
    return candidate


def _remap_field_refs(clones: "List[_Element]", renames: "Dict[str, str]") -> None:
    """Rewrite REF/PAGEREF instructions inside the copied range that point at
    renamed bookmarks."""

    def rewrite(instr: str) -> str:
        for old, new in renames.items():
            instr = re.sub(rf"(?<=\s){re.escape(old)}(?=\s|$)", new, instr)
        return instr

    for clone in clones:
        for node in clone.iter(_INSTR_TEXT):
            if node.text and ("REF" in node.text):
                node.text = rewrite(node.text)
        for node in clone.iter(_FLD_SIMPLE):
            instr = node.get(_INSTR)
            if instr and "REF" in instr:
                node.set(_INSTR, rewrite(instr))


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


def _pad_adjacent_tables(anchor_p: "_Element", clones: "List[_Element]") -> None:
    """Two adjacent tables fuse visually in Word; keep an empty paragraph
    between a cloned table and a neighboring destination table."""
    if clones and clones[0].tag == _TBL and anchor_p.tag == _TBL:
        clones.insert(0, OxmlElement("w:p"))
    if clones and clones[-1].tag == _TBL:
        following = anchor_p.getnext()
        if following is not None and following.tag == _TBL:
            clones.append(OxmlElement("w:p"))
