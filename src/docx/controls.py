"""Content-control (structured document tag) surface — the templating
primitive (paper-docx).

Enumerate a document's controls by tag/alias, read their values, and set
them type-correctly: text controls get their runs replaced (placeholder
state cleared), checkboxes flip `w14:checked` AND their
glyph, dropdowns/combos validate against their `w:listItem` choices, dates
set `w:fullDate` alongside the display text. Data-bound controls
(`w:dataBinding`) refuse: their value lives in a custom XML part Word
re-syncs on open, so editing the surface text would silently vanish.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, List, Optional, Tuple, Union

from docx.errors import (
    AmbiguousTargetError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import _refuse_if_protected
from docx.search import _validate_writable_text
from docx.story import _first_choice_children, _story_elements

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.document import Document

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
_SHOWING_PLC_HDR = qn("w:showingPlcHdr")
_DATA_BINDING = qn("w:dataBinding")
_DROPDOWN = qn("w:dropDownList")
_COMBO = qn("w:comboBox")
_DATE = qn("w:date")
_FULL_DATE = qn("w:fullDate")
_LIST_ITEM = qn("w:listItem")
_DISPLAY_TEXT = qn("w:displayText")
_LIST_VALUE = qn("w:value")
_TEXT_CTRL = qn("w:text")
_PICTURE = qn("w:picture")
_DOC_PART_OBJ = qn("w:docPartObj")
_GROUP = qn("w:group")
_CHECKBOX = qn("w14:checkbox")
_W14_CHECKED = qn("w14:checked")
_W14_VAL = qn("w14:val")

_CHECKED_GLYPH = "☒"
_UNCHECKED_GLYPH = "☐"


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

    # -- writing ----------------------------------------------------------

    def set_value(self, value: "Union[str, bool, dt.date, dt.datetime]") -> None:
        """Set the control's value type-correctly; validate-fully-then-mutate.

        Text/rich-text: replaces the content with one run (first run's
        formatting kept, placeholder state cleared). Checkbox: bool.
        Dropdown/combo: must match a listItem (dropdown) — combos accept free
        text. Date: date/datetime (stamps `w:fullDate`) or display string.
        """
        _refuse_if_protected(self._document, "set a control value")
        if self.is_data_bound:
            raise UnsupportedStructureError(
                "control is data-bound (w:dataBinding): its value lives in a"
                " custom XML part Word re-syncs on open; editing the surface"
                " text would silently vanish"
            )
        control_type = self.control_type
        if control_type == "checkbox":
            if not isinstance(value, bool):
                raise ValueError("checkbox controls take a bool value")
            self._set_checkbox(value)
            return
        if control_type in ("picture", "group", "building_block"):
            raise UnsupportedStructureError(
                f"{control_type} controls are not settable"
            )
        if isinstance(value, bool):
            raise ValueError(f"{control_type} controls take text, not bool")
        if control_type == "date" and isinstance(value, (dt.date, dt.datetime)):
            display = value.strftime("%Y-%m-%d")
            self._set_text(display)
            self._set_full_date(value)
            return
        if control_type == "date" and isinstance(value, str):
            # a plain display string: drop any machine-readable w:fullDate so
            # the control never shows one date while claiming another
            sdt_pr = self._sdt_pr
            date_pr = sdt_pr.find(_DATE) if sdt_pr is not None else None
            if date_pr is not None and date_pr.get(_FULL_DATE):
                del date_pr.attrib[_FULL_DATE]
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

    def _set_checkbox(self, checked: bool) -> None:
        sdt_pr = self._sdt_pr
        checkbox = sdt_pr.find(_CHECKBOX)
        state = checkbox.find(_W14_CHECKED)
        if state is None:
            state = OxmlElement("w14:checked")
            checkbox.insert(0, state)
        state.set(_W14_VAL, "1" if checked else "0")
        self._set_text(_CHECKED_GLYPH if checked else _UNCHECKED_GLYPH)

    def _set_full_date(self, value: "Union[dt.date, dt.datetime]") -> None:
        sdt_pr = self._sdt_pr
        date_pr = sdt_pr.find(_DATE)
        if isinstance(value, dt.datetime):
            stamp = value.replace(microsecond=0)
        else:
            stamp = dt.datetime(value.year, value.month, value.day)
        date_pr.set(_FULL_DATE, stamp.strftime("%Y-%m-%dT%H:%M:%SZ"))

    def _set_text(self, text: str) -> None:
        from docx.search import _clear_placeholder_state

        content = self._sdt.find(_SDT_CONTENT)
        if content is None:
            content = OxmlElement("w:sdtContent")
            self._sdt.append(content)
        # structure guards: never destroy nested controls or tables, and never
        # write a bare run into block-level content that has no paragraph
        if content.find(qn("w:sdt")) is not None or any(
            node.tag == qn("w:sdt") for node in content.iter(qn("w:sdt"))
        ):
            raise UnsupportedStructureError(
                "control content holds nested content controls; set the inner"
                " controls individually instead of overwriting the group"
            )
        if content.find(qn("w:tbl")) is not None:
            raise UnsupportedStructureError(
                "control content holds a table; refusing to replace structure"
                " with plain text"
            )
        has_block_content = content.find(_P) is not None
        has_inline_content = any(child.tag == _R for child in content) or not len(content)
        if not has_block_content and not has_inline_content:
            raise UnsupportedStructureError(
                "control content is not simple text (no paragraph or runs to"
                " replace); refusing to destroy its structure"
            )
        # keep the first run's formatting; block controls keep one paragraph
        first_run = next(iter(content.iter(_R)), None)
        template_rpr = None
        if first_run is not None:
            rpr = first_run.find(_RPR)
            if rpr is not None:
                import copy

                template_rpr = copy.deepcopy(rpr)
        run = OxmlElement("w:r")
        if template_rpr is not None:
            run.append(template_rpr)
        run.add_t(text)
        first_paragraph = content.find(_P)
        if first_paragraph is not None:
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
