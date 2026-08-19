"""Missing-table APIs: comments, lock, pictures, links, captions, notes."""

from __future__ import annotations

from pathlib import Path

import docx
from docx.commentops import COMMENTS_IDS_RELATIONSHIP_TYPE, delete_comment
from docx.oxml.ns import qn
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

    def it_retargets_identity_when_the_last_paragraph_moves(self):
        document = _doc()
        comment = find_one(document, "perfectly ordinary").comment(
            text="note", author="Ada"
        )
        old_para = comment.paragraphs[-1]._p.get(qn("w14:paraId"))
        comment.add_paragraph("more")
        new_para = comment.paragraphs[-1]._p.get(qn("w14:paraId"))
        assert new_para and new_para != old_para
        ids_root = document.part.part_related_by(COMMENTS_IDS_RELATIONSHIP_TYPE)._element
        para_ids = [
            entry.get(
                "{http://schemas.microsoft.com/office/word/2016/wordml/cid}paraId"
            )
            for entry in ids_root
        ]
        assert new_para in para_ids
        assert old_para not in para_ids
        delete_comment(document, comment)
        assert list(ids_root) == []
