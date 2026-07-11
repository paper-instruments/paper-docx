"""Regressions for comment ownership and anchor-id preflights."""

from __future__ import annotations

import pytest

import docx
from docx.commentops import (
    anchored_text,
    comment_thread,
    parent_of,
    reply,
    resolve,
)
from docx.errors import PaperRefusal, UnsupportedStructureError
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml.ns import qn
from docx.search import find_one

from .harness.contract import assert_refusal_atomic


def _commented_document():
    document = docx.Document()
    document.add_paragraph("comment target")
    comment = find_one(document, "comment target").comment(
        "Parent", author="Reviewer"
    )
    return document, comment


class DescribeCommentRelationshipOwnership:
    @pytest.mark.parametrize("operation", ["resolve", "thread"])
    def it_refuses_multiple_comments_relationships_atomically(
        self, operation: str
    ):
        document, comment = _commented_document()
        package = document.part.package
        assert package is not None
        duplicate = Part(
            PackURI("/word/comments-duplicate.xml"),
            CT.WML_COMMENTS,
            b"<comments/>",
            package,
        )
        document.part.relate_to(duplicate, RT.COMMENTS)

        error = assert_refusal_atomic(
            document,
            lambda candidate: (
                resolve(candidate, comment)
                if operation == "resolve"
                else comment_thread(candidate)
            ),
            PaperRefusal,
        )

        assert isinstance(error, UnsupportedStructureError)
        assert "multiple comments relationships" in str(error)


class DescribeCommentAnchorIds:
    def it_compares_anchor_marker_ids_numerically(self):
        document, comment = _commented_document()
        comment._comment_elm.set(qn("w:id"), "1")  # noqa: SLF001
        for marker in document.element.body.iter(
            qn("w:commentRangeStart"),
            qn("w:commentRangeEnd"),
            qn("w:commentReference"),
        ):
            marker.set(qn("w:id"), "01")

        assert anchored_text(document, comment) == "comment target"
        child = reply(document, comment, "Reply", author="Second Reviewer")
        assert parent_of(document, child) == 1

    @pytest.mark.parametrize(
        "marker_name",
        ["commentRangeStart", "commentRangeEnd", "commentReference"],
    )
    @pytest.mark.parametrize("bad_id", [None, "broken"])
    def it_refuses_invalid_anchor_marker_ids_atomically(
        self, marker_name: str, bad_id: str | None
    ):
        document, comment = _commented_document()
        marker = next(document.element.body.iter(qn(f"w:{marker_name}")))
        if bad_id is None:
            del marker.attrib[qn("w:id")]
        else:
            marker.set(qn("w:id"), bad_id)

        error = assert_refusal_atomic(
            document,
            lambda candidate: reply(
                candidate, comment, "Reply", author="Second Reviewer"
            ),
            PaperRefusal,
        )

        assert isinstance(error, UnsupportedStructureError)
        assert f"malformed {marker_name} comment marker" in str(error)
