"""Custom element classes for the tracked-change (revision) vocabulary.

`w:ins` and `w:del` wrap runs whose content was inserted or deleted with
change tracking on. Deleted text lives in `w:delText` (which shares the
`CT_Text` shape), never in a live `w:t` — Word treats a `w:t` inside `w:del`
as corrupt.

This is paper-docx's first new XML vocabulary; it follows the comments
feature (v1.2.0) structurally: declarative descriptors here, registration in
`docx/oxml/__init__.py`, emission through these classes only (never
hand-assembled lxml in proxy code).
"""

from __future__ import annotations

import copy
import datetime as dt
from typing import TYPE_CHECKING, Optional

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_DateTime, ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
)

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.oxml.text.run import CT_R


class CT_RunTrackChange(BaseOxmlElement):
    """`w:ins` and `w:del` elements: a tracked run-level change.

    Both tags share this shape (ECMA-376 CT_RunTrackChange): required unique
    `w:id`, required `w:author`, optional `w:date`, containing the changed
    runs. `w:del` also appears childless inside `w:pPr/w:rPr` to mark a
    paragraph mark as deleted; ``ZeroOrMore`` runs covers that use too.
    """

    id = RequiredAttribute("w:id", ST_DecimalNumber)  # noqa: A003
    author = RequiredAttribute("w:author", ST_String)
    date = OptionalAttribute("w:date", ST_DateTime)
    r = ZeroOrMore("w:r")

    @classmethod
    def new(
        cls,
        tag: str,
        revision_id: int,
        author: str,
        date: Optional[dt.datetime],
    ) -> "CT_RunTrackChange":
        """A new `w:ins` or `w:del` element with its identity attributes set."""
        from docx.oxml.parser import OxmlElement

        if tag not in ("w:ins", "w:del"):
            raise ValueError(f"tag must be 'w:ins' or 'w:del', got {tag!r}")
        element = OxmlElement(tag)
        element.id = revision_id
        element.author = author
        if date is not None:
            element.date = date
        return element

    def add_tracked_run(
        self, text: str, rpr: "Optional[_Element]", *, deleted: bool
    ) -> "CT_R":
        """Append a run holding `text`, with `rpr` cloned in when given.

        Deleted text goes into `w:delText` (never live `w:t` inside `w:del`);
        inserted text into `w:t`. Edge whitespace gets `xml:space="preserve"`.
        """
        from docx.oxml.parser import OxmlElement

        run = OxmlElement("w:r")
        if rpr is not None:
            run.insert(0, copy.deepcopy(rpr))
        if deleted:
            del_text = OxmlElement("w:delText")
            del_text.text = text
            if text[:1].isspace() or text[-1:].isspace():
                del_text.set(qn("xml:space"), "preserve")
            run.append(del_text)
        else:
            run.add_t(text)
        self.append(run)
        return run
