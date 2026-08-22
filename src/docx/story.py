"""Visibility-complete story traversal and inspection (paper-docx, opt-in).

Standard python-docx traversal (`Document.paragraphs`, `.tables`,
`iter_inner_content()`) is blind to text inside tracked insertions and
deletions, content controls, text boxes, and to entire story parts
(footnotes, endnotes). This module is the *new, explicitly named* perception
layer: existing traversal semantics are untouched.

Traversal rules:

* Every story part is walked: body, headers, footers, footnotes, endnotes,
  comments. Separator/continuation-separator footnotes and endnotes are
  plumbing, not content, and are skipped.
* Paragraphs inside table cells are not emitted as separate blocks — their
  text belongs to the table block. Paragraphs inside content controls and
  text boxes ARE emitted, flagged.
* `mc:AlternateContent` contributes its first supported `mc:Choice`, or its
  `mc:Fallback` when none of the choices' required namespaces are supported.
  Exactly one branch is traversed, so duplicated compatibility content is
  never counted twice.
* Empty paragraphs are emitted — block indices must be stable, and an empty
  paragraph is a real edit target.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from typing import (
    TYPE_CHECKING,
    AbstractSet,
    ClassVar,
    Dict,
    Iterator,
    List,
    Optional,
    Tuple,
    cast,
)

from docx import _textatoms
from docx._guard import check_install
from docx._normalize import normalize_text
from docx.oxml.ns import qn

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

check_install()

VIEWS = ("current", "original", "all")

_T = _textatoms.T
_DEL_TEXT = _textatoms.DEL_TEXT
_FLD_CHAR = _textatoms.FLD_CHAR
_P = qn("w:p")
_TBL = qn("w:tbl")
_SDT = qn("w:sdt")
_SDT_CONTENT = qn("w:sdtContent")
_INS = qn("w:ins")
_DEL = qn("w:del")
_MOVE_FROM = qn("w:moveFrom")
_MOVE_TO = qn("w:moveTo")
_TXBX = qn("w:txbxContent")
_FLD_SIMPLE = qn("w:fldSimple")
_FLD_CHAR_TYPE = qn("w:fldCharType")

#: tracked property-change vocabulary — enumerable, countable, not resolvable
_FORMAT_CHANGE_TAGS = frozenset(
    qn(tag)
    for tag in (
        "w:rPrChange", "w:pPrChange", "w:tblPrChange", "w:tcPrChange",
        "w:trPrChange", "w:sectPrChange", "w:numberingChange",
        "w:cellIns", "w:cellDel", "w:cellMerge",
    )
)
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_MATH_TAGS = (f"{{{_M_NS}}}oMath", f"{{{_M_NS}}}oMathPara")
_OBJECT = qn("w:object")
_ALT_CHUNK = qn("w:altChunk")
_VANISH = qn("w:vanish")
_P_STYLE_XPATH = "./w:pPr/w:pStyle/@w:val"
_FOOTNOTE_TYPE = qn("w:type")

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_ALTERNATE = f"{{{_MC_NS}}}AlternateContent"
_MC_CHOICE = f"{{{_MC_NS}}}Choice"
_MC_FALLBACK = f"{{{_MC_NS}}}Fallback"

# Namespace capabilities this traversal actually understands. ``Requires``
# names prefixes, but support is a property of their namespace URIs; merely
# declaring an unknown prefix does not make its choice processable.
_SUPPORTED_MC_NAMESPACES = frozenset(
    (
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "http://schemas.microsoft.com/office/word/2012/wordml",
        "urn:schemas-microsoft-com:office:office",
        "urn:schemas-microsoft-com:office:word",
        "urn:schemas-microsoft-com:vml",
    )
)


def _story_sort_key(name: str) -> Tuple[int, str]:
    """Traversal order: body, headers, footers, footnotes, endnotes, comments."""
    if name == "word/document.xml":
        return (0, name)
    if name.startswith("word/header"):
        return (1, name)
    if name.startswith("word/footer"):
        return (2, name)
    if name == "word/footnotes.xml":
        return (3, name)
    if name == "word/endnotes.xml":
        return (4, name)
    return (5, name)  # word/comments.xml


@dataclass(frozen=True)
class Anchor:
    """Legacy, inert location evidence: story part + index + content hash.

    ``Anchor`` remains serializable for historical search and revision result
    data. It is not a mutation-capable block target; reacquire a live |Block|
    or |Span|, or persist a ``BlockLocator`` instead.
    """

    story: str
    index: int
    content_hash: str

    def to_dict(self) -> "Dict[str, object]":
        return {"story": self.story, "index": self.index, "content_hash": self.content_hash}


@dataclass(frozen=True)
class TableShape:
    rows: int
    columns: int
    has_merges: bool
    has_nested_table: bool

    def to_dict(self) -> "Dict[str, object]":
        return {
            "rows": self.rows,
            "columns": self.columns,
            "has_merges": self.has_merges,
            "has_nested_table": self.has_nested_table,
        }


def _strict_keys(
    value: object, expected: "AbstractSet[str]", *, label: str
) -> "Dict[str, object]":
    if not isinstance(value, dict):
        raise ValueError(f"{label} must contain exactly {sorted(expected)!r}")
    data = cast("Dict[object, object]", value)
    if not all(isinstance(key, str) for key in data) or set(data) != expected:
        raise ValueError(f"{label} must contain exactly {sorted(expected)!r}")
    return cast("Dict[str, object]", data)


def _strict_str(value: object, *, label: str) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{label} must be a string")
    return value


def _strict_optional_str(value: object, *, label: str) -> Optional[str]:
    if value is not None and not isinstance(value, str):
        raise ValueError(f"{label} must be a string or null")
    return value


def _strict_bool(value: object, *, label: str) -> bool:
    if not isinstance(value, bool):
        raise ValueError(f"{label} must be a boolean")
    return value


def _strict_int(value: object, *, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValueError(f"{label} must be an integer")
    return value


@dataclass(frozen=True)
class _TableCellEvidence:
    text: str
    grid_span: int
    vertical_merge: Optional[str]
    nested_tables: "Tuple[_TableEvidence, ...]"

    def to_dict(self) -> "Dict[str, object]":
        return {
            "text": self.text,
            "grid_span": self.grid_span,
            "vertical_merge": self.vertical_merge,
            "nested_tables": [table.to_dict() for table in self.nested_tables],
        }

    @classmethod
    def from_dict(cls, value: object) -> "_TableCellEvidence":
        data = _strict_keys(
            value,
            {"text", "grid_span", "vertical_merge", "nested_tables"},
            label="table cell evidence",
        )
        nested = data["nested_tables"]
        if not isinstance(nested, list):
            raise ValueError("table cell nested_tables must be a list")
        nested_items = cast("List[object]", nested)
        grid_span = _strict_int(data["grid_span"], label="table cell grid_span")
        if grid_span < 1:
            raise ValueError("table cell grid_span must be >= 1")
        return cls(
            text=_strict_str(data["text"], label="table cell text"),
            grid_span=grid_span,
            vertical_merge=_strict_optional_str(
                data["vertical_merge"], label="table cell vertical_merge"
            ),
            nested_tables=tuple(
                _TableEvidence.from_dict(item) for item in nested_items
            ),
        )


@dataclass(frozen=True)
class _TableEvidence:
    shape: TableShape
    rows: "Tuple[Tuple[_TableCellEvidence, ...], ...]"

    def to_dict(self) -> "Dict[str, object]":
        return {
            "shape": self.shape.to_dict(),
            "rows": [[cell.to_dict() for cell in row] for row in self.rows],
        }

    @classmethod
    def from_dict(cls, value: object) -> "_TableEvidence":
        data = _strict_keys(value, {"shape", "rows"}, label="table evidence")
        shape_data = _strict_keys(
            data["shape"],
            {"rows", "columns", "has_merges", "has_nested_table"},
            label="table shape",
        )
        shape = TableShape(
            rows=_strict_int(shape_data["rows"], label="table shape rows"),
            columns=_strict_int(shape_data["columns"], label="table shape columns"),
            has_merges=_strict_bool(
                shape_data["has_merges"], label="table shape has_merges"
            ),
            has_nested_table=_strict_bool(
                shape_data["has_nested_table"],
                label="table shape has_nested_table",
            ),
        )
        rows = data["rows"]
        if not isinstance(rows, list):
            raise ValueError("table evidence rows must be a list of lists")
        row_items = cast("List[object]", rows)
        if not all(isinstance(row, list) for row in row_items):
            raise ValueError("table evidence rows must be a list of lists")
        parsed_rows = tuple(
            tuple(
                _TableCellEvidence.from_dict(cell)
                for cell in cast("List[object]", row)
            )
            for row in row_items
        )
        if shape.rows != len(parsed_rows) or shape.columns != max(
            (len(row) for row in parsed_rows), default=0
        ):
            raise ValueError("table evidence rows do not match its declared shape")
        return cls(shape=shape, rows=parsed_rows)


@dataclass(frozen=True)
class _BlockEvidence:
    kind: str
    text: str
    style_id: Optional[str]
    in_content_control: bool
    in_text_box: bool
    container_path: "Tuple[Tuple[str, Optional[str]], ...]"
    table: Optional[_TableEvidence]

    def to_dict(self) -> "Dict[str, object]":
        return {
            "kind": self.kind,
            "text": self.text,
            "style_id": self.style_id,
            "in_content_control": self.in_content_control,
            "in_text_box": self.in_text_box,
            "container_path": [
                {"tag": tag, "id": identifier}
                for tag, identifier in self.container_path
            ],
            "table": self.table.to_dict() if self.table else None,
        }

    @classmethod
    def from_dict(cls, value: object) -> "_BlockEvidence":
        data = _strict_keys(
            value,
            {
                "kind",
                "text",
                "style_id",
                "in_content_control",
                "in_text_box",
                "container_path",
                "table",
            },
            label="block evidence",
        )
        kind = _strict_str(data["kind"], label="block evidence kind")
        if kind not in ("paragraph", "table"):
            raise ValueError("block evidence kind must be 'paragraph' or 'table'")
        path_data = data["container_path"]
        if not isinstance(path_data, list) or not path_data:
            raise ValueError("block evidence container_path must be a non-empty list")
        path: "List[Tuple[str, Optional[str]]]" = []
        for item in cast("List[object]", path_data):
            entry = _strict_keys(item, {"tag", "id"}, label="container path item")
            path.append(
                (
                    _strict_str(entry["tag"], label="container path tag"),
                    _strict_optional_str(entry["id"], label="container path id"),
                )
            )
        table_data = data["table"]
        table = None if table_data is None else _TableEvidence.from_dict(table_data)
        if (kind == "table") != (table is not None):
            raise ValueError("block evidence table details must match its kind")
        style_id = _strict_optional_str(data["style_id"], label="block style_id")
        if kind == "table" and style_id is not None:
            raise ValueError("table block evidence cannot carry a paragraph style_id")
        return cls(
            kind=kind,
            text=_strict_str(data["text"], label="block evidence text"),
            style_id=style_id,
            in_content_control=_strict_bool(
                data["in_content_control"], label="block in_content_control"
            ),
            in_text_box=_strict_bool(data["in_text_box"], label="block in_text_box"),
            container_path=tuple(path),
            table=table,
        )


@dataclass(frozen=True)
class _NeighborEvidence:
    boundary: Optional[str] = None
    block: Optional[_BlockEvidence] = None

    def to_dict(self) -> "Dict[str, object]":
        if self.boundary is not None:
            return {"boundary": self.boundary}
        assert self.block is not None
        return {"block": self.block.to_dict()}

    @classmethod
    def from_dict(cls, value: object, *, side: str) -> "_NeighborEvidence":
        if not isinstance(value, dict):
            raise ValueError(f"{side} context must be an object")
        data = cast("Dict[object, object]", value)
        if set(data) == {"boundary"}:
            boundary = _strict_str(data["boundary"], label=f"{side} boundary")
            expected = "start" if side == "previous" else "end"
            if boundary != expected:
                raise ValueError(f"{side} boundary must be {expected!r}")
            return cls(boundary=boundary)
        if set(data) == {"block"}:
            return cls(block=_BlockEvidence.from_dict(data["block"]))
        raise ValueError(
            f"{side} context must contain exactly 'boundary' or exactly 'block'"
        )


@dataclass(frozen=True)
class BlockLocator:
    """Versioned, portable block evidence resolved fail-closed.

    A locator is inert data until an operation evaluates all of its exact
    story/view/kind/content/topology and adjacent-block evidence against a
    document. Its positional hint and optional Word paragraph ID never select
    a candidate on their own.
    """

    SCHEMA: ClassVar[str] = "paper_block_locator"
    VERSION: ClassVar[int] = 1

    story: str
    view: str
    kind: str
    evidence: _BlockEvidence
    previous: _NeighborEvidence
    next: _NeighborEvidence
    position_hint: int
    paragraph_id: Optional[str]

    def __post_init__(self) -> None:
        _strict_str(self.story, label="block locator story")
        view = _strict_str(self.view, label="block locator view")
        if view not in VIEWS:
            raise ValueError(f"block locator view must be one of {VIEWS!r}")
        kind = _strict_str(self.kind, label="block locator kind")
        if kind not in ("paragraph", "table"):
            raise ValueError("block locator kind must be 'paragraph' or 'table'")
        evidence_value = cast("object", self.evidence)
        previous_value = cast("object", self.previous)
        next_value = cast("object", self.next)
        if not isinstance(evidence_value, _BlockEvidence):
            raise ValueError("block locator evidence must contain block evidence")
        if evidence_value.kind != kind:
            raise ValueError("block locator kind contradicts its evidence")
        if not isinstance(previous_value, _NeighborEvidence):
            raise ValueError("block locator previous context is invalid")
        if not isinstance(next_value, _NeighborEvidence):
            raise ValueError("block locator next context is invalid")
        try:
            evidence = _BlockEvidence.from_dict(evidence_value.to_dict())
            previous = _NeighborEvidence.from_dict(
                previous_value.to_dict(), side="previous"
            )
            next_evidence = _NeighborEvidence.from_dict(
                next_value.to_dict(), side="next"
            )
        except (AttributeError, AssertionError, TypeError) as exc:
            raise ValueError("block locator contains malformed evidence") from exc
        if (
            evidence != evidence_value
            or previous != previous_value
            or next_evidence != next_value
        ):
            raise ValueError("block locator evidence must use canonical field types")
        position_hint = _strict_int(
            self.position_hint, label="block locator position_hint"
        )
        if position_hint < 0:
            raise ValueError("block locator position_hint must be >= 0")
        paragraph_id = _strict_optional_str(
            self.paragraph_id, label="block locator paragraph_id"
        )
        if kind == "table" and paragraph_id is not None:
            raise ValueError("table block locators cannot carry a paragraph_id")

    def to_dict(self) -> "Dict[str, object]":
        return {
            "schema": self.SCHEMA,
            "version": self.VERSION,
            "story": self.story,
            "view": self.view,
            "kind": self.kind,
            "evidence": self.evidence.to_dict(),
            "context": {
                "previous": self.previous.to_dict(),
                "next": self.next.to_dict(),
            },
            "position_hint": self.position_hint,
            "paragraph_id": self.paragraph_id,
        }

    @classmethod
    def from_dict(cls, value: object) -> "BlockLocator":
        data = _strict_keys(
            value,
            {
                "schema",
                "version",
                "story",
                "view",
                "kind",
                "evidence",
                "context",
                "position_hint",
                "paragraph_id",
            },
            label="block locator",
        )
        if data["schema"] != cls.SCHEMA:
            raise ValueError(f"block locator schema must be {cls.SCHEMA!r}")
        if _strict_int(data["version"], label="block locator version") != cls.VERSION:
            raise ValueError(f"unsupported block locator version {data['version']!r}")
        story = _strict_str(data["story"], label="block locator story")
        view = _strict_str(data["view"], label="block locator view")
        kind = _strict_str(data["kind"], label="block locator kind")
        evidence = _BlockEvidence.from_dict(data["evidence"])
        context = _strict_keys(
            data["context"], {"previous", "next"}, label="block locator context"
        )
        position_hint = _strict_int(
            data["position_hint"], label="block locator position_hint"
        )
        paragraph_id = _strict_optional_str(
            data["paragraph_id"], label="block locator paragraph_id"
        )
        return cls(
            story=story,
            view=view,
            kind=kind,
            evidence=evidence,
            previous=_NeighborEvidence.from_dict(
                context["previous"], side="previous"
            ),
            next=_NeighborEvidence.from_dict(context["next"], side="next"),
            position_hint=position_hint,
            paragraph_id=paragraph_id,
        )


@dataclass(frozen=True)
class Block:
    """One live, owner-bound paragraph or table observed during traversal."""

    story: str
    kind: str  # "paragraph" | "table"
    index: int
    anchor: Anchor
    text: str
    style_id: Optional[str]
    in_insert: bool
    in_delete: bool
    in_content_control: bool
    in_text_box: bool
    has_field: bool
    table: Optional[TableShape]
    locator: Optional[BlockLocator] = None
    _document: "Optional[Document]" = field(default=None, repr=False, compare=False)
    _element: "Optional[_Element]" = field(default=None, repr=False, compare=False)
    _story_root: "Optional[_Element]" = field(default=None, repr=False, compare=False)
    _parent: "Optional[_Element]" = field(default=None, repr=False, compare=False)
    _container_elements: "Tuple[_Element, ...]" = field(
        default=(), repr=False, compare=False
    )
    _view: str = field(default="current", repr=False, compare=False)

    def to_dict(self) -> "Dict[str, object]":
        return {
            "story": self.story,
            "kind": self.kind,
            "index": self.index,
            "anchor": self.anchor.to_dict(),
            "anchor_role": "legacy_inert_location_evidence",
            "text": self.text,
            "style_id": self.style_id,
            "in_insert": self.in_insert,
            "in_delete": self.in_delete,
            "in_content_control": self.in_content_control,
            "in_text_box": self.in_text_box,
            "has_field": self.has_field,
            "table": self.table.to_dict() if self.table else None,
            "locator": self.locator.to_dict() if self.locator else None,
        }


@dataclass(frozen=True)
class Outline:
    """Inspection snapshot of a document: every block in every story part."""

    story_parts: Tuple[str, ...]
    blocks: Tuple[Block, ...]
    blind_region_counts: Dict[str, int]

    def to_dict(self) -> "Dict[str, object]":
        return {
            "schema": "paper_outline",
            "version": 3,  # v3: inert Anchor evidence + exact BlockLocator data
            "story_parts": list(self.story_parts),
            "blind_region_counts": dict(sorted(self.blind_region_counts.items())),
            "blocks": [block.to_dict() for block in self.blocks],
        }


def content_hash(text: str) -> str:
    """First 8 hex chars of SHA-256 over the block's normalized text."""
    return hashlib.sha256(normalize_text(text).encode("utf-8")).hexdigest()[:8]


def _story_elements(document: "Document") -> "List[Tuple[str, _Element]]":
    """(story-name, root element) for every story part, deterministic order."""
    package = document.part.package
    assert package is not None
    found: "List[Tuple[str, _Element]]" = []
    for part in package.iter_parts():
        name = str(part.partname).lstrip("/").casefold()
        element = getattr(part, "_element", None)
        if element is None:
            continue
        if name == "word/document.xml" or (
            name.startswith(("word/header", "word/footer")) and name.endswith(".xml")
        ) or name in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"):
            found.append((name, element))
    return sorted(found, key=lambda item: _story_sort_key(item[0]))


def story_parts(document: "Document") -> Tuple[str, ...]:
    """Every story part present in `document`, in traversal order."""
    return tuple(name for name, _ in _story_elements(document))


def _first_choice_children(element: "_Element") -> "List[_Element]":
    """Children with each ``mc:AlternateContent`` collapsed to one branch.

    Choices are considered in document order and selected only when every
    prefix in their required namespace list is one this traversal supports.
    If no choice qualifies, the fallback branch is used. Missing fallback
    content honestly contributes nothing.
    """

    def choice_is_supported(choice: "_Element") -> bool:
        requires = (choice.get("Requires") or "").split()
        return bool(requires) and all(
            choice.nsmap.get(prefix) in _SUPPORTED_MC_NAMESPACES
            for prefix in requires
        )

    result: "List[_Element]" = []
    for child in element:
        if child.tag == _MC_ALTERNATE:
            selected = None
            fallback = None
            for alt_child in child:
                if alt_child.tag == _MC_CHOICE and choice_is_supported(alt_child):
                    selected = alt_child
                    break
                if alt_child.tag == _MC_FALLBACK and fallback is None:
                    fallback = alt_child
            selected = selected if selected is not None else fallback
            if selected is not None:
                result.extend(_first_choice_children(selected))
        else:
            result.append(child)
    return result


class _TextVisitor:
    """Accumulates view-filtered text and region flags over a subtree.

    Tracked MOVES participate in the views: `w:moveFrom` content is
    deletion-like (excluded from "current", present in "original") and
    `w:moveTo` content is insertion-like — so moved text appears exactly once
    per view instead of doubling. Resolution of
    moves is a separate, refused concern.
    """

    def __init__(self, view: str) -> None:
        self.view = view
        self.pieces: List[str] = []
        self.in_insert = False
        self.in_delete = False
        self.in_content_control = False
        self.in_text_box = False
        self.has_field = False

    def visit(self, element: "_Element", *, in_ins: bool, in_del: bool,
              in_sdt: bool, in_txbx: bool, skip_text_boxes: bool) -> None:
        tag = element.tag
        if tag == _TXBX and skip_text_boxes:
            return
        if tag in (_INS, _MOVE_TO):
            in_ins = True
        elif tag in (_DEL, _MOVE_FROM):
            in_del = True
        elif tag == _SDT:
            in_sdt = True
        elif tag == _TXBX:
            in_txbx = True
        elif tag == _FLD_SIMPLE or (
            tag == _FLD_CHAR and element.get(_FLD_CHAR_TYPE) == "begin"
        ):
            self.has_field = True

        if tag == _T:
            if self.view == "current" and in_del:
                return  # moveFrom source text: gone once changes are accepted
            if self.view == "original" and in_ins:
                return
            self._emit(element.text or "", in_ins, in_del, in_sdt, in_txbx)
            return
        if tag == _DEL_TEXT:
            if self.view == "current":
                return
            if self.view == "original" and in_ins:
                # a deletion nested inside a pending insertion never existed
                # in the original document
                return
            self._emit(element.text or "", in_ins, True, in_sdt, in_txbx)
            return
        if tag == _textatoms.INSTR_TEXT:
            # Field instructions are searchable so edits can detect and
            # refuse them, but they are not visible document text.
            return
        if _textatoms.is_direct_run_child(element):
            projection = _textatoms.project_run_child(element)
            if projection.barrier:
                return
            if not projection.text:
                return
            if self.view == "current" and in_del:
                return
            if self.view == "original" and in_ins:
                return
            self._emit(projection.text, in_ins, in_del, in_sdt, in_txbx)
            return
        for child in _first_choice_children(element):
            self.visit(child, in_ins=in_ins, in_del=in_del, in_sdt=in_sdt,
                       in_txbx=in_txbx, skip_text_boxes=skip_text_boxes)

    def _emit(self, text: str, in_ins: bool, in_del: bool, in_sdt: bool,
              in_txbx: bool) -> None:
        self.pieces.append(text)
        self.in_insert = self.in_insert or in_ins
        self.in_delete = self.in_delete or in_del
        self.in_content_control = self.in_content_control or in_sdt
        self.in_text_box = self.in_text_box or in_txbx

    @property
    def text(self) -> str:
        return "".join(self.pieces)


def _subtree_text(element: "_Element", view: str, *, skip_text_boxes: bool,
                  in_sdt: bool = False, in_txbx: bool = False) -> _TextVisitor:
    visitor = _TextVisitor(view)
    for child in _first_choice_children(element):
        visitor.visit(child, in_ins=False, in_del=False, in_sdt=in_sdt,
                      in_txbx=in_txbx, skip_text_boxes=skip_text_boxes)
    return visitor


def _table_shape(table: "_Element") -> TableShape:
    rows = table.findall(qn("w:tr"))
    columns = max((len(row.findall(qn("w:tc"))) for row in rows), default=0)
    has_merges = bool(
        table.findall(f".//{qn('w:vMerge')}") or table.findall(f".//{qn('w:gridSpan')}")
    )
    has_nested = any(t is not table for t in table.iter(_TBL))
    return TableShape(
        rows=len(rows), columns=columns, has_merges=has_merges, has_nested_table=has_nested
    )


def _text_box_contents(paragraph: "_Element") -> "List[_Element]":
    """w:txbxContent elements reachable from `paragraph`, fallbacks excluded."""
    found: "List[_Element]" = []

    def walk(element: "_Element") -> None:
        for child in _first_choice_children(element):
            if child.tag == _TXBX:
                found.append(child)
            else:
                walk(child)

    walk(paragraph)
    return found


def _block_containers(story: str, root: "_Element") -> "Iterator[Tuple[_Element, bool]]":
    """(container, in_content_control) holders of block-level content."""
    tag = root.tag
    if tag == qn("w:document"):
        body = root.find(qn("w:body"))
        if body is not None:
            yield body, False
    elif tag in (qn("w:hdr"), qn("w:ftr")):
        yield root, False
    elif tag in (qn("w:footnotes"), qn("w:endnotes")):
        for note in root:
            if note.tag not in (qn("w:footnote"), qn("w:endnote")):
                continue
            note_type = note.get(_FOOTNOTE_TYPE)
            if note_type in ("separator", "continuationSeparator"):
                continue  # plumbing, not content
            yield note, False
    elif tag == qn("w:comments"):
        for comment in root:
            if comment.tag == qn("w:comment"):
                yield comment, False
    else:  # pragma: no cover - unknown story roots are a programming error
        raise ValueError(f"unrecognized story root {tag!r} in {story}")


def _walk_container(
    container: "_Element",
    counter: List[int],
    *,
    in_sdt: bool,
    in_txbx: bool,
) -> "Iterator[Tuple[str, int, _Element, bool, bool]]":
    """(kind, block-index, element, in_sdt, in_txbx) for each block, in order.

    THE single definition of block identity and indexing — `iter_blocks` and
    `docx.search` both ride on it, so a span's block anchor can never disagree
    with the outline's.
    """
    for child in _first_choice_children(container):
        if child.tag == _P:
            index = counter[0]
            counter[0] += 1
            yield ("paragraph", index, child, in_sdt, in_txbx)
            for txbx in _text_box_contents(child):
                yield from _walk_container(txbx, counter, in_sdt=in_sdt, in_txbx=True)
        elif child.tag == _TBL:
            index = counter[0]
            counter[0] += 1
            yield ("table", index, child, in_sdt, in_txbx)
        elif child.tag == _SDT:
            content = child.find(_SDT_CONTENT)
            if content is not None:
                yield from _walk_container(content, counter, in_sdt=True, in_txbx=in_txbx)


def _iter_block_elements(
    story: str, root: "_Element"
) -> "Iterator[Tuple[str, int, _Element, bool, bool]]":
    counter = [0]
    for container, in_sdt in _block_containers(story, root):
        yield from _walk_container(container, counter, in_sdt=in_sdt, in_txbx=False)


def _count_fldchar_delta(element: "_Element") -> int:
    """Net change in open complex-field depth across `element`: begins minus ends.

    Goes negative when a block closes more fields than it opens; `iter_blocks` clamps at the
    call site.
    """
    delta = 0

    def walk(node: "_Element") -> None:
        nonlocal delta
        for child in _first_choice_children(node):
            if child.tag == _FLD_CHAR:
                fld_type = child.get(_FLD_CHAR_TYPE)
                if fld_type == "begin":
                    delta += 1
                elif fld_type == "end":
                    delta -= 1
            walk(child)

    walk(element)
    return delta


def _container_elements(element: "_Element", root: "_Element") -> "Tuple[_Element, ...]":
    elements: "List[_Element]" = []
    current = element.getparent()
    while current is not None:
        elements.append(current)
        if current is root:
            return tuple(elements)
        current = current.getparent()
    return tuple(elements)


def _container_path(
    element: "_Element", root: "_Element"
) -> "Tuple[Tuple[str, Optional[str]], ...]":
    identifying_tags = {
        qn("w:footnote"),
        qn("w:endnote"),
        qn("w:comment"),
        qn("w:sdt"),
    }
    return tuple(
        (
            str(container.tag),
            container.get(qn("w:id")) if container.tag in identifying_tags else None,
        )
        for container in _container_elements(element, root)
    )


def _nested_tables(cell: "_Element") -> "Tuple[_Element, ...]":
    return tuple(
        element
        for kind, _index, element, _in_sdt, _in_txbx in _walk_container(
            cell, [0], in_sdt=False, in_txbx=False
        )
        if kind == "table"
    )


def _table_evidence(table: "_Element", view: str) -> _TableEvidence:
    rows: "List[Tuple[_TableCellEvidence, ...]]" = []
    for row in table.findall(qn("w:tr")):
        cells: "List[_TableCellEvidence]" = []
        for cell in row.findall(qn("w:tc")):
            tc_pr = cell.find(qn("w:tcPr"))
            grid_span = tc_pr.find(qn("w:gridSpan")) if tc_pr is not None else None
            vertical_merge = tc_pr.find(qn("w:vMerge")) if tc_pr is not None else None
            cells.append(
                _TableCellEvidence(
                    text=_subtree_text(cell, view, skip_text_boxes=False).text,
                    grid_span=(
                        int(grid_span.get(qn("w:val")) or 1)
                        if grid_span is not None
                        else 1
                    ),
                    vertical_merge=(
                        vertical_merge.get(qn("w:val")) or "continue"
                        if vertical_merge is not None
                        else None
                    ),
                    nested_tables=tuple(
                        _table_evidence(nested, view) for nested in _nested_tables(cell)
                    ),
                )
            )
        rows.append(tuple(cells))
    return _TableEvidence(shape=_table_shape(table), rows=tuple(rows))


def _block_snapshot(
    kind: str,
    element: "_Element",
    view: str,
    root: "_Element",
    *,
    in_sdt: bool,
    in_txbx: bool,
) -> "Tuple[str, Optional[str], Optional[TableShape], _TextVisitor, _BlockEvidence]":
    if kind == "table":
        visitor = _subtree_text(element, view, skip_text_boxes=False,
                                in_sdt=in_sdt, in_txbx=in_txbx)
        text = _table_text(element, view)
        style_id = None
        table = _table_shape(element)
        table_evidence = _table_evidence(element, view)
    else:
        visitor = _subtree_text(element, view, skip_text_boxes=True,
                                in_sdt=in_sdt, in_txbx=in_txbx)
        text = visitor.text
        style_values = element.xpath(_P_STYLE_XPATH)
        style_id = str(style_values[0]) if style_values else None
        table = None
        table_evidence = None
    evidence = _BlockEvidence(
        kind=kind,
        text=text,
        style_id=style_id,
        in_content_control=in_sdt or visitor.in_content_control,
        in_text_box=in_txbx or visitor.in_text_box,
        container_path=_container_path(element, root),
        table=table_evidence,
    )
    return text, style_id, table, visitor, evidence


@dataclass(frozen=True)
class _BlockRecord:
    kind: str
    index: int
    element: "_Element"
    text: str
    style_id: Optional[str]
    table: Optional[TableShape]
    visitor: _TextVisitor
    evidence: _BlockEvidence
    in_open_field: bool


def _build_story_blocks(
    document: "Document", story: str, root: "_Element", view: str
) -> "Tuple[Block, ...]":
    """Canonical live-block and locator-evidence builder for one story."""
    records: "List[_BlockRecord]" = []
    open_field_depth = 0
    for kind, index, element, in_sdt, in_txbx in _iter_block_elements(story, root):
        text, style_id, table, visitor, evidence = _block_snapshot(
            kind, element, view, root, in_sdt=in_sdt, in_txbx=in_txbx
        )
        records.append(
            _BlockRecord(
                kind=kind,
                index=index,
                element=element,
                text=text,
                style_id=style_id,
                table=table,
                visitor=visitor,
                evidence=evidence,
                in_open_field=open_field_depth > 0,
            )
        )
        open_field_depth = max(0, open_field_depth + _count_fldchar_delta(element))

    blocks: "List[Block]" = []
    for position, record in enumerate(records):
        previous = (
            _NeighborEvidence(boundary="start")
            if position == 0
            else _NeighborEvidence(block=records[position - 1].evidence)
        )
        next_evidence = (
            _NeighborEvidence(boundary="end")
            if position == len(records) - 1
            else _NeighborEvidence(block=records[position + 1].evidence)
        )
        paragraph_id = (
            record.element.get(qn("w14:paraId"))
            if record.kind == "paragraph"
            else None
        )
        locator = BlockLocator(
            story=story,
            view=view,
            kind=record.kind,
            evidence=record.evidence,
            previous=previous,
            next=next_evidence,
            position_hint=record.index,
            paragraph_id=paragraph_id,
        )
        containers = _container_elements(record.element, root)
        blocks.append(
            Block(
                story=story,
                kind=record.kind,
                index=record.index,
                anchor=Anchor(
                    story=story,
                    index=record.index,
                    content_hash=content_hash(record.text),
                ),
                text=record.text,
                style_id=record.style_id,
                in_insert=record.visitor.in_insert,
                in_delete=record.visitor.in_delete,
                in_content_control=record.evidence.in_content_control,
                in_text_box=record.evidence.in_text_box,
                # a block BETWEEN a field's begin and end (TOC entry paragraphs) is
                # field content even though neither marker lives in it
                has_field=record.visitor.has_field or record.in_open_field,
                table=record.table,
                locator=locator,
                _document=document,
                _element=record.element,
                _story_root=root,
                _parent=record.element.getparent(),
                _container_elements=containers,
                _view=view,
            )
        )
    return tuple(blocks)


def _table_text(table: "_Element", view: str) -> str:
    """Cell texts in row-major order, newline-joined (the table block owns
    all text inside it, nested content included)."""
    pieces: List[str] = []
    for row in table.findall(qn("w:tr")):
        for cell in row.findall(qn("w:tc")):
            visitor = _subtree_text(cell, view, skip_text_boxes=False)
            pieces.append(visitor.text)
    return "\n".join(pieces)


#: every key blind_region_counts reports, in payload order
BLIND_REGION_KEYS = (
    "tracked_insertions",
    "tracked_deletions",
    "moves",
    "format_changes",
    "content_controls",
    "text_boxes",
    "fields",
    "math",
    "embedded_objects",
    "alt_chunks",
    "hidden_text",
)


def _count_blind_regions(root: "_Element") -> Dict[str, int]:
    """Occurrences of each region the traversal flags — or CANNOT read — in
    traversal space (mc:Fallback duplicates excluded).

    The last four keys are the honesty confession: math, embedded
    objects (OLE/charts/SmartArt), altChunk imports and hidden (`w:vanish`)
    text hold content this package does not surface; a non-zero count says
    "there is more here than the outline shows".
    """
    counts = dict.fromkeys(BLIND_REGION_KEYS, 0)
    tag_keys = {
        _INS: "tracked_insertions",
        _DEL: "tracked_deletions",
        _MOVE_FROM: "moves",
        _MOVE_TO: "moves",
        _SDT: "content_controls",
        _TXBX: "text_boxes",
        _FLD_SIMPLE: "fields",
        _MATH_TAGS[0]: "math",
        _MATH_TAGS[1]: "math",
        _OBJECT: "embedded_objects",
        _ALT_CHUNK: "alt_chunks",
        _VANISH: "hidden_text",
    }

    graphic_data = (
        "{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData"
    )
    _SURFACED_GRAPHIC_URIS = (
        "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    )

    def walk(element: "_Element") -> None:
        for child in _first_choice_children(element):
            key = tag_keys.get(child.tag)
            if key is not None:
                counts[key] += 1
            elif child.tag in _FORMAT_CHANGE_TAGS:
                counts["format_changes"] += 1
            elif child.tag == _FLD_CHAR and child.get(_FLD_CHAR_TYPE) == "begin":
                counts["fields"] += 1
            elif child.tag == graphic_data and (
                child.get("uri") not in _SURFACED_GRAPHIC_URIS
            ):
                # charts, SmartArt, OLE previews — content we cannot read
                counts["embedded_objects"] += 1
            walk(child)

    walk(root)
    return counts


def iter_blocks(document: "Document", *, view: str = "current") -> Iterator[Block]:
    """Every block in every story part of `document`, in document order.

    `view` selects the text layer: "current" (insertions in, deletions out —
    the document if all changes were accepted), "original" (deletions in,
    insertions out), or "all" (everything).
    """
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}, got {view!r}")
    for story, root in _story_elements(document):
        yield from _build_story_blocks(document, story, root, view)


def outline(document: "Document", *, view: str = "current") -> Outline:
    """Inspection snapshot: story parts, all blocks, blind-region counts.

    Deterministic: the same document yields byte-identical `to_dict()` output
    on every call (inspection determinism).
    """
    blocks = tuple(iter_blocks(document, view=view))
    totals = dict.fromkeys(BLIND_REGION_KEYS, 0)
    for _, root in _story_elements(document):
        for key, value in _count_blind_regions(root).items():
            totals[key] += value
    return Outline(
        story_parts=story_parts(document),
        blocks=blocks,
        blind_region_counts=totals,
    )
