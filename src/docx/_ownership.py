"""Ownership checks for live document proxies used by paper-docx verbs.

Value anchors are re-resolved against the document passed to an operation.
Live proxies such as ``Span`` and ``Comment`` are different: they retain XML
elements from the package that created them, so accepting one from another
document can mutate that other package.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.errors import BoundaryViolationError, TargetNotFoundError

if TYPE_CHECKING:
    from docx.comments import Comment
    from docx.document import Document
    from docx.search import Span


def require_span_owner(
    document: "Document", span: "Span", *, argument: str = "span"
) -> None:
    """Refuse a live span captured from a different document package."""
    owner = getattr(span, "_document", None)
    owner_part = getattr(owner, "part", None)
    if owner_part is not document.part:
        raise BoundaryViolationError(
            f"{argument} belongs to a different document; re-find the text"
            " in the document passed to this operation"
        )


def require_anchor_owner(
    document: "Document", anchor: object, *, argument: str = "anchor"
) -> None:
    """Check ownership for live anchor forms; value/string anchors are inert."""
    if getattr(anchor, "_document", None) is not None:
        require_span_owner(document, cast("Span", anchor), argument=argument)


def require_comment_owner(
    document: "Document", comment: "Comment", *, argument: str = "comment"
) -> None:
    """Refuse a comment proxy from another part and detect detached proxies."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        comment_part = comment.part
    except (AttributeError, ValueError):
        comment_part = None
    try:
        expected_part = document.part.part_related_by(RT.COMMENTS)
    except KeyError:
        expected_part = None

    if comment_part is not expected_part or expected_part is None:
        raise BoundaryViolationError(
            f"{argument} belongs to a different document; select the comment"
            " from the document passed to this operation"
        )

    element = getattr(comment, "_comment_elm", None)
    expected_element = getattr(expected_part, "_element", None)
    if element is None or element.getparent() is not expected_element:
        raise TargetNotFoundError(
            f"{argument} is stale: its comment was removed from the document"
        )
