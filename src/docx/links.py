"""Create and retarget hyperlinks on existing text."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx._guard import check_install
from docx._transaction import rollback_on_error
from docx.errors import BoundaryViolationError, UnsupportedStructureError
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.protection import _refuse_if_protected
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    from docx.document import Document
    from docx.search import Span

check_install()

_P = qn("w:p")
_R = qn("w:r")
_RPR = qn("w:rPr")
_RSTYLE = qn("w:rStyle")
_HYPERLINK = qn("w:hyperlink")


class _PartHost:
    """Parent whose `.part` is the story that owns the hyperlink XML."""

    def __init__(self, part):
        self.part = part


def _part_for_span(document: "Document", span: "Span"):
    wanted = span.story.lstrip("/").casefold()
    package = document.part.package
    if package is None:
        raise UnsupportedStructureError(
            f"cannot resolve story part for {span.story}; nothing was changed"
        )
    for part in package.iter_parts():
        if str(part.partname).lstrip("/").casefold() == wanted:
            return part
    raise UnsupportedStructureError(
        f"span story {span.story} has no package part; nothing was changed"
    )


def add_hyperlink(document: "Document", span: "Span", address: str) -> Hyperlink:
    """Wrap `span`'s runs in a hyperlink pointing at `address`.

    The span must lie in a single paragraph. Visible text stays; Word
    shows it as a link after the file opens.
    """
    if not address:
        raise ValueError("address must be a non-empty URL")
    _refuse_if_protected(document, "add a hyperlink")
    with rollback_on_error(document, span):
        span._isolate_edge_runs()  # noqa: SLF001
        runs = []
        for atom in span._atoms:  # noqa: SLF001
            run = atom.run
            if run is not None and not any(existing is run for existing in runs):
                runs.append(run)
        if not runs:
            raise BoundaryViolationError(
                "span has no runs to wrap in a hyperlink; nothing was changed"
            )
        parents = {run.getparent() for run in runs}
        if len(parents) != 1:
            raise BoundaryViolationError(
                "cannot wrap a hyperlink across a paragraph boundary; nothing was changed"
            )
        parent = runs[0].getparent()
        while parent is not None and parent.tag != _P:
            parent = parent.getparent()
        if parent is None:
            raise BoundaryViolationError(
                "hyperlink runs are not inside a paragraph; nothing was changed"
            )
        for run in runs:
            ancestor = run.getparent()
            while ancestor is not None and ancestor.tag != _P:
                if ancestor.tag == _HYPERLINK:
                    raise UnsupportedStructureError(
                        "span is already inside a hyperlink; retarget it"
                        " with Hyperlink.address. Nothing was changed"
                    )
                ancestor = ancestor.getparent()
        story_part = _part_for_span(document, span)
        r_id = story_part.relate_to(address, RT.HYPERLINK, is_external=True)
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)
        hyperlink.set(qn("w:history"), "1")
        first = runs[0]
        first.addprevious(hyperlink)
        for run in runs:
            hyperlink.append(run)
            r_pr = run.find(_RPR)
            if r_pr is None:
                r_pr = OxmlElement("w:rPr")
                run.insert(0, r_pr)
            style = r_pr.find(_RSTYLE)
            if style is None:
                style = OxmlElement("w:rStyle")
                r_pr.insert(0, style)
            style.set(qn("w:val"), "Hyperlink")
        paragraph = Paragraph(parent, _PartHost(story_part))
        return Hyperlink(hyperlink, paragraph)
