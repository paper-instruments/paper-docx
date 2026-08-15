"""Content-control (structured document tag) surface — the templating
primitive (paper-docx).

Enumerate a document's controls by tag/alias, read their values, and set
them type-correctly: text controls get their runs replaced (placeholder
state cleared), checkboxes flip `w14:checked` AND their
glyph, dropdowns/combos validate against their `w:listItem` choices, dates
set `w:fullDate` alongside the display text. Data-bound controls
(`w:dataBinding`) write the custom XML store Word re-syncs on open, then
update the visible surface.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, List, Optional, Tuple, Union

from lxml import etree

from docx._guard import check_install
from docx._transaction import rollback_on_error
from docx.errors import (
    AmbiguousTargetError,
    BoundaryViolationError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.opc.constants import CONTENT_TYPE as CT
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import _refuse_if_protected
from docx.search import _validate_writable_text
from docx.story import _first_choice_children, _story_elements

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

check_install()

_SDT = qn("w:sdt")
_SDT_PR = qn("w:sdtPr")
_SDT_CONTENT = qn("w:sdtContent")
_ALIAS = qn("w:alias")
_TAG = qn("w:tag")
_W_VAL = qn("w:val")
_T = qn("w:t")
_R = qn("w:r")
_P = qn("w:p")
_RPR = qn("w:rPr")
_BLOCK_CONTROL_PARENTS = frozenset(
    qn(tag)
    for tag in (
        "w:body",
        "w:tc",
        "w:hdr",
        "w:ftr",
        "w:footnote",
        "w:endnote",
        "w:comment",
        "w:txbxContent",
    )
)
_SHOWING_PLC_HDR = qn("w:showingPlcHdr")
_DATA_BINDING = qn("w:dataBinding")
_STORE_ITEM_ID = qn("w:storeItemID")
_XPATH = qn("w:xpath")
_PREFIX_MAPPINGS = qn("w:prefixMappings")
_DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
_DS_ITEM_ID = f"{{{_DS_NS}}}itemID"
_XSI_NIL = "{http://www.w3.org/2001/XMLSchema-instance}nil"
_LOCK = qn("w:lock")
_DROPDOWN = qn("w:dropDownList")
_COMBO = qn("w:comboBox")
_DATE = qn("w:date")
_FULL_DATE = qn("w:fullDate")
_DATE_FORMAT = qn("w:dateFormat")
_LID = qn("w:lid")
_CALENDAR = qn("w:calendar")
_LIST_ITEM = qn("w:listItem")
_DISPLAY_TEXT = qn("w:displayText")
_LIST_VALUE = qn("w:value")
_TEXT_CTRL = qn("w:text")
_PICTURE = qn("w:picture")
_DOC_PART_OBJ = qn("w:docPartObj")
_GROUP = qn("w:group")
_CHECKBOX = qn("w14:checkbox")
_W14_CHECKED = qn("w14:checked")
_W14_CHECKED_STATE = qn("w14:checkedState")
_W14_UNCHECKED_STATE = qn("w14:uncheckedState")
_W14_VAL = qn("w14:val")
_W14_FONT = qn("w14:font")
_R_FONTS = qn("w:rFonts")

_CHECKED_GLYPH = "☒"
_UNCHECKED_GLYPH = "☐"


def _control_level(sdt: "_Element", content: "Optional[_Element]") -> str:
    """Return ``block`` or ``inline`` when the control shape is unambiguous."""
    parent = sdt.getparent()
    parent_level = None
    if parent is not None:
        if parent.tag == _P:
            parent_level = "inline"
        elif parent.tag in _BLOCK_CONTROL_PARENTS:
            parent_level = "block"

    if content is not None and len(content):
        has_paragraph = any(child.tag == _P for child in content)
        has_run = any(child.tag == _R for child in content)
        if has_paragraph == has_run:
            raise UnsupportedStructureError(
                "control content has an ambiguous block/inline shape; nothing was changed"
            )
        content_level = "block" if has_paragraph else "inline"
        if parent_level is not None and content_level != parent_level:
            raise UnsupportedStructureError(
                "control content does not match its block/inline parent; nothing was changed"
            )
        return content_level

    if parent_level is None:
        raise UnsupportedStructureError(
            "empty control has an ambiguous block/inline parent; nothing was changed"
        )
    return parent_level

_MONTH_NAMES = (
    "",
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)
_WEEKDAY_NAMES = (
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
)


def _declared_checkbox_state(
    checkbox: "_Element", checked: bool
) -> "Tuple[str, Optional[str]]":
    """The state glyph/font declared by the control, with Word defaults."""
    state_tag = _W14_CHECKED_STATE if checked else _W14_UNCHECKED_STATE
    fallback = _CHECKED_GLYPH if checked else _UNCHECKED_GLYPH
    state = checkbox.find(state_tag)
    if state is None:
        return fallback, None
    raw = state.get(_W14_VAL)
    codepoint = -1
    try:
        codepoint = int(raw, 16) if raw is not None else -1
        glyph = chr(codepoint)
    except (ValueError, OverflowError):
        glyph = ""
    if not glyph or 0xD800 <= codepoint <= 0xDFFF:
        raise UnsupportedStructureError(
            f"checkbox declares invalid hexadecimal glyph {raw!r}; nothing was changed"
        )
    return glyph, state.get(_W14_FONT)


def _utc_date_stamp(value: "Union[dt.date, dt.datetime]") -> dt.datetime:
    if isinstance(value, dt.datetime):
        stamp = value
        if stamp.utcoffset() is not None:
            stamp = stamp.astimezone(dt.timezone.utc)
    else:
        stamp = dt.datetime(value.year, value.month, value.day)
    return stamp.replace(microsecond=0)


def _format_declared_date(date_pr: "_Element", stamp: dt.datetime) -> str:
    """Format ``stamp`` using the supported subset of Word date patterns."""
    calendar_elm = date_pr.find(_CALENDAR)
    calendar = calendar_elm.get(_W_VAL) if calendar_elm is not None else None
    if calendar not in (None, "gregorian"):
        raise UnsupportedStructureError(
            f"date control declares unsupported calendar {calendar!r};"
            " nothing was changed"
        )
    format_elm = date_pr.find(_DATE_FORMAT)
    if format_elm is None:
        raise UnsupportedStructureError(
            "date control has no w:dateFormat; nothing was changed"
        )
    pattern = format_elm.get(_W_VAL) or ""
    if not pattern:
        raise UnsupportedStructureError(
            "date control declares an empty w:dateFormat; nothing was changed"
        )

    lid_elm = date_pr.find(_LID)
    language = lid_elm.get(_W_VAL) if lid_elm is not None else None
    pieces: List[str] = []
    uses_names = False
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char in ("'", '"'):
            quote = char
            index += 1
            literal: List[str] = []
            while index < len(pattern) and pattern[index] != quote:
                literal.append(pattern[index])
                index += 1
            if index >= len(pattern):
                raise UnsupportedStructureError(
                    f"unsupported date format {pattern!r}: unterminated literal;"
                    " nothing was changed"
                )
            pieces.append("".join(literal))
            index += 1
            continue
        if char == "\\":
            if index + 1 >= len(pattern):
                raise UnsupportedStructureError(
                    f"unsupported date format {pattern!r}: trailing escape;"
                    " nothing was changed"
                )
            pieces.append(pattern[index + 1])
            index += 2
            continue
        if not char.isalpha():
            pieces.append(char)
            index += 1
            continue
        end = index + 1
        while end < len(pattern) and pattern[end] == char:
            end += 1
        token = pattern[index:end]
        hour_12 = stamp.hour % 12 or 12
        replacements = {
            "d": str(stamp.day),
            "dd": f"{stamp.day:02d}",
            "ddd": _WEEKDAY_NAMES[stamp.weekday()][:3],
            "dddd": _WEEKDAY_NAMES[stamp.weekday()],
            "M": str(stamp.month),
            "MM": f"{stamp.month:02d}",
            "MMM": _MONTH_NAMES[stamp.month][:3],
            "MMMM": _MONTH_NAMES[stamp.month],
            "y": str(stamp.year),
            "yy": f"{stamp.year % 100:02d}",
            "yyyy": f"{stamp.year:04d}",
            "H": str(stamp.hour),
            "HH": f"{stamp.hour:02d}",
            "h": str(hour_12),
            "hh": f"{hour_12:02d}",
            "m": str(stamp.minute),
            "mm": f"{stamp.minute:02d}",
            "s": str(stamp.second),
            "ss": f"{stamp.second:02d}",
            "t": ("A" if stamp.hour < 12 else "P"),
            "tt": ("AM" if stamp.hour < 12 else "PM"),
        }
        replacement = replacements.get(token)
        if replacement is None:
            raise UnsupportedStructureError(
                f"unsupported date format {pattern!r} (token {token!r});"
                " nothing was changed"
            )
        uses_names = uses_names or token in ("ddd", "dddd", "MMM", "MMMM", "t", "tt")
        pieces.append(replacement)
        index = end

    if uses_names:
        if language is None:
            raise UnsupportedStructureError(
                f"date format {pattern!r} uses locale-sensitive names, but the"
                " date control has no w:lid; nothing was changed"
            )
        normalized_language = language.lower()
        if normalized_language != "en" and not normalized_language.startswith("en-"):
            raise UnsupportedStructureError(
                f"date format {pattern!r} requires locale {language!r}, but only"
                " English date names are supported; nothing was changed"
            )
    return "".join(pieces)


@dataclass(frozen=True)
class ControlInfo:
    """Identity and current state of one content control."""

    tag: Optional[str]
    alias: Optional[str]
    control_type: str  # text|rich_text|checkbox|dropdown|combo|date|picture|group|building_block
    value: str
    story: str
    showing_placeholder: bool
    is_data_bound: bool
    choices: Tuple[str, ...]  # dropdown/combo display texts

    def to_dict(self) -> dict:
        return {
            "tag": self.tag,
            "alias": self.alias,
            "control_type": self.control_type,
            "value": self.value,
            "story": self.story,
            "showing_placeholder": self.showing_placeholder,
            "is_data_bound": self.is_data_bound,
            "choices": list(self.choices),
        }


def _control_type(sdt_pr: "Optional[_Element]") -> str:
    if sdt_pr is None:
        return "rich_text"
    if sdt_pr.find(_CHECKBOX) is not None:
        return "checkbox"
    if sdt_pr.find(_DROPDOWN) is not None:
        return "dropdown"
    if sdt_pr.find(_COMBO) is not None:
        return "combo"
    if sdt_pr.find(_DATE) is not None:
        return "date"
    if sdt_pr.find(_PICTURE) is not None:
        return "picture"
    if sdt_pr.find(_DOC_PART_OBJ) is not None:
        return "building_block"
    if sdt_pr.find(_GROUP) is not None:
        return "group"
    if sdt_pr.find(_TEXT_CTRL) is not None:
        return "text"
    return "rich_text"


def _choices(sdt_pr: "Optional[_Element]") -> Tuple[str, ...]:
    if sdt_pr is None:
        return ()
    for holder_tag in (_DROPDOWN, _COMBO):
        holder = sdt_pr.find(holder_tag)
        if holder is not None:
            return tuple(
                item.get(_DISPLAY_TEXT) or item.get(_LIST_VALUE) or ""
                for item in holder.findall(_LIST_ITEM)
            )
    return ()


def _content_text(sdt: "_Element") -> str:
    content = sdt.find(_SDT_CONTENT)
    if content is None:
        return ""
    return "".join(
        node.text or "" for node in content.iter(_T)
    )


class Control:
    """Live proxy for one content control."""

    def __init__(self, sdt: "_Element", story: str, document: "Document") -> None:
        self._sdt = sdt
        self._story = story
        self._document = document

    @property
    def _sdt_pr(self) -> "Optional[_Element]":
        return self._sdt.find(_SDT_PR)

    def _pr_value(self, tag) -> Optional[str]:
        sdt_pr = self._sdt_pr
        node = sdt_pr.find(tag) if sdt_pr is not None else None
        return node.get(_W_VAL) if node is not None else None

    @property
    def tag(self) -> Optional[str]:
        return self._pr_value(_TAG)

    @property
    def alias(self) -> Optional[str]:
        return self._pr_value(_ALIAS)

    @property
    def control_type(self) -> str:
        return _control_type(self._sdt_pr)

    @property
    def is_data_bound(self) -> bool:
        sdt_pr = self._sdt_pr
        return sdt_pr is not None and sdt_pr.find(_DATA_BINDING) is not None

    @property
    def showing_placeholder(self) -> bool:
        sdt_pr = self._sdt_pr
        return sdt_pr is not None and sdt_pr.find(_SHOWING_PLC_HDR) is not None

    @property
    def value(self) -> "Union[str, bool]":
        """The control's current value (bool for checkboxes, text otherwise)."""
        if self.control_type == "checkbox":
            sdt_pr = self._sdt_pr
            checkbox = sdt_pr.find(_CHECKBOX) if sdt_pr is not None else None
            checked = checkbox.find(_W14_CHECKED) if checkbox is not None else None
            return bool(checked is not None and checked.get(_W14_VAL) in ("1", "true"))
        return _content_text(self._sdt)

    def info(self) -> ControlInfo:
        value = self.value
        return ControlInfo(
            tag=self.tag,
            alias=self.alias,
            control_type=self.control_type,
            value=str(value),
            story=self._story,
            showing_placeholder=self.showing_placeholder,
            is_data_bound=self.is_data_bound,
            choices=_choices(self._sdt_pr),
        )

    def _validate_ownership(self) -> None:
        """Refuse a control no longer active in the document that created it."""
        story_roots = tuple(root for _story, root in _story_elements(self._document))
        if any(
            candidate is self._sdt
            for root in story_roots
            for candidate in _iter_sdts(root)
        ):
            return

        tree_root = self._sdt
        while tree_root.getparent() is not None:
            tree_root = tree_root.getparent()
        if any(tree_root is root for root in story_roots):
            raise TargetNotFoundError(
                "control is stale: it is no longer in the document's active"
                " content-control traversal"
            )
        if tree_root.tag in {root.tag for root in story_roots}:
            raise BoundaryViolationError(
                "control belongs to a different document; reacquire it from"
                " the document that now contains it"
            )
        raise TargetNotFoundError(
            "control is stale: it was removed from its document"
        )

    # -- writing ----------------------------------------------------------

    def set_value(self, value: "Union[str, bool, dt.date, dt.datetime]") -> None:
        """Set the control's value type-correctly; validate-fully-then-mutate.

        Text/rich-text: replaces the content with one run (first run's
        formatting kept, placeholder state cleared). Checkbox: bool.
        Dropdown/combo: must match a listItem (dropdown) — combos accept free
        text. Date: date/datetime (stamps `w:fullDate`) or display string.
        """
        self._validate_ownership()
        _refuse_if_protected(self._document, "set a control value")
        if self.is_data_bound:
            control_type = self.control_type
            if control_type not in ("text", "rich_text"):
                raise UnsupportedStructureError(
                    f"data-bound {control_type} controls are not settable;"
                    " nothing was changed"
                )
            if not isinstance(value, str):
                raise ValueError("data-bound controls take a string value")
            _validate_writable_text(value, argument="value")
            _refuse_locked_control_content(self._sdt)
            binding = self._sdt_pr.find(_DATA_BINDING)
            with rollback_on_error(self._document):
                _write_bound_store(self._document, binding, value)
                self._set_text(value)
            return
        _refuse_control_write_restrictions(self._sdt)
        control_type = self.control_type
        if control_type == "checkbox":
            if not isinstance(value, bool):
                raise ValueError("checkbox controls take a bool value")
            sdt_pr = self._sdt_pr
            checkbox = sdt_pr.find(_CHECKBOX) if sdt_pr is not None else None
            if checkbox is None:
                raise UnsupportedStructureError(
                    "checkbox control has no w14:checkbox properties; nothing was changed"
                )
            glyph, font = _declared_checkbox_state(checkbox, value)
            _validate_writable_text(glyph, argument="declared checkbox glyph")
            # _set_text performs every structure check before mutation. Only
            # after it succeeds is the machine-readable state changed.
            self._set_text(glyph, font=font)
            self._set_checkbox_state(checkbox, value)
            return
        if control_type in ("picture", "group", "building_block"):
            raise UnsupportedStructureError(
                f"{control_type} controls are not settable"
            )
        if isinstance(value, bool):
            raise ValueError(f"{control_type} controls take text, not bool")
        if control_type == "date" and isinstance(value, (dt.date, dt.datetime)):
            sdt_pr = self._sdt_pr
            date_pr = sdt_pr.find(_DATE) if sdt_pr is not None else None
            if date_pr is None:
                raise UnsupportedStructureError(
                    "date control has no w:date properties; nothing was changed"
                )
            stamp = _utc_date_stamp(value)
            display = _format_declared_date(date_pr, stamp)
            _validate_writable_text(display, argument="formatted date value")
            self._set_text(display)
            date_pr.set(
                _FULL_DATE,
                stamp.strftime("%Y-%m-%dT%H:%M:%SZ"),
            )
            return
        if not isinstance(value, str):
            raise ValueError(f"{control_type} controls take a string value")
        _validate_writable_text(value, argument="value")
        if control_type == "dropdown":
            choices = _choices(self._sdt_pr)
            if value not in choices:
                raise TargetNotFoundError(
                    f"{value!r} is not one of this dropdown's choices {list(choices)}"
                )
        self._set_text(value)
        if control_type == "date":
            # A plain display string has no trustworthy machine-readable date.
            # Drop fullDate only after the text replacement has succeeded.
            sdt_pr = self._sdt_pr
            date_pr = sdt_pr.find(_DATE) if sdt_pr is not None else None
            if date_pr is not None and date_pr.get(_FULL_DATE):
                del date_pr.attrib[_FULL_DATE]

    def _set_checkbox_state(self, checkbox: "_Element", checked: bool) -> None:
        state = checkbox.find(_W14_CHECKED)
        if state is None:
            state = OxmlElement("w14:checked")
            checkbox.insert(0, state)
        state.set(_W14_VAL, "1" if checked else "0")

    def _set_text(self, text: str, *, font: Optional[str] = None) -> None:
        from docx.search import _clear_placeholder_state

        content = self._sdt.find(_SDT_CONTENT)
        # structure guards: never destroy nested controls or tables, and never
        # write a bare run into block-level content that has no paragraph
        if content is not None:
            if next(iter(content.iter(_SDT)), None) is not None:
                raise UnsupportedStructureError(
                    "control content holds nested content controls; set the inner"
                    " controls individually instead of overwriting the group"
                )
            if next(iter(content.iter(qn("w:tbl"))), None) is not None:
                raise UnsupportedStructureError(
                    "control content holds a table; refusing to replace structure"
                    " with plain text"
                )
            has_block_content = content.find(_P) is not None
            has_inline_content = any(child.tag == _R for child in content) or not len(
                content
            )
            if not has_block_content and not has_inline_content:
                raise UnsupportedStructureError(
                    "control content is not simple text (no paragraph or runs to"
                    " replace); refusing to destroy its structure"
                )
        control_level = _control_level(self._sdt, content)

        # Everything is validated. Keep the first run's formatting; block
        # controls keep one paragraph.
        first_run = next(iter(content.iter(_R)), None) if content is not None else None
        template_rpr = None
        if first_run is not None:
            rpr = first_run.find(_RPR)
            if rpr is not None:
                import copy

                template_rpr = copy.deepcopy(rpr)
        if font:
            if template_rpr is None:
                template_rpr = OxmlElement("w:rPr")
            fonts = template_rpr.find(_R_FONTS)
            if fonts is None:
                fonts = OxmlElement("w:rFonts")
                template_rpr.insert(0, fonts)
            for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                fonts.set(qn(attribute), font)
        if content is None:
            content = OxmlElement("w:sdtContent")
            self._sdt.append(content)
        run = OxmlElement("w:r")
        if template_rpr is not None:
            run.append(template_rpr)
        run.add_t(text)
        first_paragraph = content.find(_P)
        if control_level == "block":
            if first_paragraph is None:
                first_paragraph = OxmlElement("w:p")
                content.append(first_paragraph)
            for child in list(first_paragraph):
                if child.tag not in (qn("w:pPr"),):
                    first_paragraph.remove(child)
            first_paragraph.append(run)
            for extra in list(content):
                if extra is not first_paragraph:
                    content.remove(extra)
        else:
            for child in list(content):
                content.remove(child)
            content.append(run)
        _clear_placeholder_state(self._sdt)


def _validate_span_surface_edit(sdt: "_Element") -> None:
    """Allow generic edits only on plain or rich-text control surfaces."""
    sdt_pr = sdt.find(_SDT_PR)
    if sdt_pr is not None and sdt_pr.find(_DATA_BINDING) is not None:
        raise UnsupportedStructureError(
            "span lies in a data-bound control whose surface text Word"
            " re-syncs from custom XML; use the binding source instead"
        )
    control_type = _control_type(sdt_pr)
    if control_type not in ("text", "rich_text"):
        raise UnsupportedStructureError(
            f"span lies in a {control_type} control whose display text is tied"
            " to machine-readable control state; use Control.set_value()"
        )


def _prefix_map(raw: Optional[str]) -> dict:
    if not raw:
        return {}
    return {
        match.group(1): match.group(2)
        for match in re.finditer(
            r"xmlns:([A-Za-z0-9_]+)\s*=\s*['\"]([^'\"]+)['\"]", raw
        )
    }


def _part_root(part):
    element = getattr(part, "_element", None)
    if element is not None:
        return element
    if "xml" not in (part.content_type or "").lower():
        return None
    blob = getattr(part, "blob", None)
    if not blob:
        return None
    return etree.fromstring(blob)


def _write_bound_node(node, value: str) -> None:
    if isinstance(node, etree._Element):
        if _XSI_NIL in node.attrib:
            del node.attrib[_XSI_NIL]
        node.text = value
        return
    parent = node.getparent() if hasattr(node, "getparent") else None
    attrname = getattr(node, "attrname", None)
    if parent is not None and attrname and getattr(node, "is_attribute", False):
        parent.set(attrname, value)
        return
    raise UnsupportedStructureError(
        "xpath matched a non-element store target; nothing was changed"
    )


def _write_bound_store(document: "Document", binding: "_Element", value: str) -> None:
    store_id = binding.get(_STORE_ITEM_ID)
    xpath = binding.get(_XPATH)
    if not store_id or not xpath:
        raise UnsupportedStructureError(
            "data-bound control is missing storeItemID or xpath; nothing was changed"
        )
    nsmap = _prefix_map(binding.get(_PREFIX_MAPPINGS))
    props_part = None
    for part in document.part.package.iter_parts():
        if part.content_type != CT.OFC_CUSTOM_XML_PROPERTIES:
            continue
        root = _part_root(part)
        if root is None:
            continue
        if root.get(_DS_ITEM_ID) == store_id:
            props_part = part
            break
    if props_part is None:
        raise TargetNotFoundError(
            f"no custom XML store with itemID {store_id!r}; nothing was changed"
        )
    item_part = None
    for owner in document.part.package.iter_parts():
        rels = getattr(owner, "rels", None)
        if not rels:
            continue
        for rel in rels.values():
            if rel.is_external:
                continue
            if rel.target_part is props_part:
                item_part = owner
                break
        if item_part is not None:
            break
    if item_part is None or item_part is props_part:
        raise TargetNotFoundError(
            "custom XML properties have no item payload; nothing was changed"
        )
    item_root = _part_root(item_part)
    if item_root is None:
        raise UnsupportedStructureError(
            "custom XML item is not writable XML; nothing was changed"
        )
    nodes = item_root.xpath(xpath, namespaces=nsmap or None)
    if not nodes:
        raise TargetNotFoundError(
            f"xpath {xpath!r} matched no node in the custom XML store"
        )
    _write_bound_node(nodes[0], value)
    if getattr(item_part, "_element", None) is None:
        item_part._blob = etree.tostring(
            item_root, xml_declaration=True, encoding="UTF-8"
        )


def _refuse_locked_control_content(sdt: "_Element") -> None:
    sdt_pr = sdt.find(_SDT_PR)
    if sdt_pr is None:
        return
    locks = sdt_pr.findall(_LOCK)
    if len(locks) > 1:
        raise UnsupportedStructureError(
            "control has multiple w:lock declarations; nothing was changed"
        )
    if locks and locks[0].get(_W_VAL) in ("contentLocked", "sdtContentLocked"):
        raise UnsupportedStructureError(
            "control content is locked; nothing was changed"
        )


def _refuse_control_write_restrictions(sdt: "_Element") -> None:
    """Honor binding and content locks without rejecting typed controls."""
    sdt_pr = sdt.find(_SDT_PR)
    if sdt_pr is None:
        return
    if sdt_pr.find(_DATA_BINDING) is not None:
        raise UnsupportedStructureError(
            "control is data-bound; edits could be discarded on sync"
        )
    _refuse_locked_control_content(sdt)


def _iter_sdts(element: "_Element") -> "Iterator[_Element]":
    for child in _first_choice_children(element):
        if child.tag == _SDT:
            yield child
        yield from _iter_sdts(child)


def iter_controls(document: "Document") -> "Iterator[Control]":
    """Every content control in every story part, document order."""
    for story, root in _story_elements(document):
        for sdt in _iter_sdts(root):
            yield Control(sdt, story, document)


def list_controls(document: "Document") -> Tuple[ControlInfo, ...]:
    """Identity and state of every control (schema-stable via .to_dict())."""
    return tuple(control.info() for control in iter_controls(document))


def get_control(
    document: "Document",
    *,
    tag: Optional[str] = None,
    alias: Optional[str] = None,
) -> Control:
    """The single control matching `tag` and/or `alias`, or a typed refusal."""
    if tag is None and alias is None:
        raise ValueError("provide tag= and/or alias=")
    matches = [
        control
        for control in iter_controls(document)
        if (tag is None or control.tag == tag)
        and (alias is None or control.alias == alias)
    ]
    wanted = f"tag={tag!r}" if tag is not None else f"alias={alias!r}"
    if not matches:
        raise TargetNotFoundError(f"no content control with {wanted}")
    if len(matches) > 1:
        raise AmbiguousTargetError(
            f"{len(matches)} content controls match {wanted}; give controls"
            " unique tags in the template"
        )
    return matches[0]


def set_control_value(
    document: "Document",
    value: "Union[str, bool, dt.date, dt.datetime]",
    *,
    tag: Optional[str] = None,
    alias: Optional[str] = None,
) -> None:
    """Find one control by tag/alias and set its value (see Control.set_value)."""
    get_control(document, tag=tag, alias=alias).set_value(value)
