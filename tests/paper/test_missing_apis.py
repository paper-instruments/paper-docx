"""Missing-table APIs: comments, lock, pictures, links, captions, notes."""

from __future__ import annotations

from pathlib import Path

import docx
from docx.commentops import delete_comment
from docx.search import find_one

from .harness.contract import save_and_reopen
from .harness.paths import fixture_path

MINIMAL = "generated/minimal-clean/minimal.docx"


def _doc(relpath: str = MINIMAL):
    return docx.Document(str(fixture_path(relpath)))


class DescribeCommentDeleteAndIdentity:
    def it_writes_modern_comment_identity_parts(self):
        document = _doc()
        find_one(document, "perfectly ordinary").comment(text="note", author="Ada")
        names = {str(part.partname) for part in document.part.package.iter_parts()}
        assert "/word/commentsIds.xml" in names
        assert "/word/commentsExtensible.xml" in names

    def it_deletes_one_comment_and_leaves_the_rest(self, tmp_path: Path):
        document = _doc()
        first = find_one(document, "perfectly ordinary").comment(text="keep", author="Ada")
        second = find_one(document, "First body paragraph").comment(
            text="drop", author="Ada"
        )
        delete_comment(document, second)
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        texts = [comment.text for comment in reopened.comments]
        assert "keep" in texts
        assert "drop" not in texts
        assert first.comment_id in {comment.comment_id for comment in reopened.comments}
