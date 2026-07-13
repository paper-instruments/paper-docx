"""Shared projection rules for text-bearing WordprocessingML run children.

The inspection and editing surfaces must agree on which run children produce
plain text. Unknown run content is not silently transparent: it contributes a
private object-replacement barrier to search so a match cannot jump across a
visible symbol, reference, drawing, or extension that we do not model.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from lxml.etree import _Element


BARRIER_TEXT = "\ufffc"

T = qn("w:t")
DEL_TEXT = qn("w:delText")
INSTR_TEXT = qn("w:instrText")
TEXT_TAGS = (T, DEL_TEXT, INSTR_TEXT)

R = qn("w:r")
RPR = qn("w:rPr")
TAB = qn("w:tab")
PTAB = qn("w:ptab")
BR = qn("w:br")
CR = qn("w:cr")
NO_BREAK_HYPHEN = qn("w:noBreakHyphen")
FLD_CHAR = qn("w:fldChar")
LAST_RENDERED_PAGE_BREAK = qn("w:lastRenderedPageBreak")

_W_TYPE = qn("w:type")
_TRANSPARENT_RUN_CHILDREN = frozenset((RPR, FLD_CHAR, LAST_RENDERED_PAGE_BREAK))


@dataclass(frozen=True)
class RunChildProjection:
    """Plain-text contribution of one direct ``w:r`` child.

    ``barrier`` means the child is visible but has no safe plain-text model.
    Search uses :data:`BARRIER_TEXT` to split otherwise adjacent text; outline
    traversal omits the sentinel while still refusing to recurse into content
    it cannot honestly interpret.
    """

    text: str
    barrier: bool = False


def project_run_child(element: "_Element") -> RunChildProjection:
    """Return the conservative text projection for a direct ``w:r`` child."""
    tag = element.tag
    if tag in TEXT_TAGS:
        return RunChildProjection(element.text or "")
    if tag in (TAB, PTAB):
        return RunChildProjection("\t")
    if tag == CR:
        return RunChildProjection("\n")
    if tag == BR:
        # Only a text-wrapping break has a newline text equivalent. A page or
        # column break remains a visible boundary but contributes no character.
        if (element.get(_W_TYPE) or "textWrapping") == "textWrapping":
            return RunChildProjection("\n")
        return RunChildProjection(BARRIER_TEXT, barrier=True)
    if tag == NO_BREAK_HYPHEN:
        # Match python-docx's public Run.text representation.
        return RunChildProjection("-")
    if tag in _TRANSPARENT_RUN_CHILDREN:
        return RunChildProjection("")
    return RunChildProjection(BARRIER_TEXT, barrier=True)


def is_direct_run_child(element: "_Element") -> bool:
    """Whether ``element`` is a direct child of a ``w:r`` element."""
    parent = element.getparent()
    return parent is not None and parent.tag == R
