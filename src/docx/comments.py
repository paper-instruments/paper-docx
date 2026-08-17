# pyright: reportPrivateUsage=false

"""Collection providing access to comments added to this document."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Callable, Iterator, Optional, cast

from docx._transaction import rollback_on_error, rollback_xml_on_error
from docx.blkcntnr import BlockItemContainer
from docx.errors import TargetNotFoundError, UnsupportedStructureError

if TYPE_CHECKING:
    from docx.document import Document
    from docx.oxml.comments import CT_Comment, CT_Comments
    from docx.parts.comments import CommentsPart
    from docx.shared import Length
    from docx.styles.style import ParagraphStyle
    from docx.table import Table
    from docx.text.paragraph import Paragraph


def _document_for_comments_part(comments_part: "CommentsPart") -> "Optional[Document]":
    """Return the live owning document, or None for a standalone test part."""
    package = comments_part.package
    if package is None:
        return None
    try:
        document_part = package.main_document_part
    except KeyError:
        return None
    except ValueError:
        raise UnsupportedStructureError(
            "multiple main-document relationships make comment mutation"
            " ambiguous; nothing was changed"
        ) from None
    document = getattr(document_part, "document", None)
    from docx.document import Document
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    if not isinstance(document, Document):
        return None
    try:
        live_part = document.part.part_related_by(RT.COMMENTS)
    except KeyError:
        raise TargetNotFoundError(
            "comment proxy is stale: its part was removed from the document"
        ) from None
    except ValueError:
        raise UnsupportedStructureError(
            "multiple comments relationships make comment mutation ambiguous;"
            " nothing was changed"
        ) from None
    if live_part is not comments_part:
        raise TargetNotFoundError(
            "comment proxy is stale: its part was replaced in the document"
        )
    return document


def _refuse_document_protection(document: "Document", operation: str) -> None:
    from docx.protection import _refuse_if_protected  # pyright: ignore[reportUnknownVariableType]

    refuse = cast("Callable[[Document, str], None]", _refuse_if_protected)
    refuse(document, operation)


class Comments:
    """Collection containing the comments added to this document."""

    def __init__(self, comments_elm: CT_Comments, comments_part: CommentsPart):
        self._comments_elm = comments_elm
        self._comments_part = comments_part

    def __iter__(self) -> Iterator[Comment]:
        """Iterator over the comments in this collection."""
        return (
            Comment(comment_elm, self._comments_part)
            for comment_elm in self._comments_elm.comment_lst
        )

    def __len__(self) -> int:
        """The number of comments in this collection."""
        return len(self._comments_elm.comment_lst)

    def add_comment(self, text: str = "", author: str = "", initials: str | None = "") -> Comment:
        """Add a new comment to the document and return it.

        The comment is added to the end of the comments collection and is assigned a unique
        comment-id.

        If `text` is provided, it is added to the comment. This option provides for the common
        case where a comment contains a modest passage of plain text. Multiple paragraphs can be
        added using the `text` argument by separating their text with newlines (`"\\\\n"`).
        Between newlines, text is interpreted as it is in `Document.add_paragraph(text=...)`.

        The default is to place a single empty paragraph in the comment, which is the same
        behavior as the Word UI when you add a comment. New runs can be added to the first
        paragraph in the empty comment with `comments.paragraphs[0].add_run()` to adding more
        complex text with emphasis or images. Additional paragraphs can be added using
        `.add_paragraph()`.

        `author` is a required attribute, set to the empty string by default.

        `initials` is an optional attribute, set to the empty string by default. Passing `None`
        for the `initials` parameter causes that attribute to be omitted from the XML.
        """
        if self._comments_elm is not self._comments_part._element:
            raise TargetNotFoundError(
                "comments collection is stale: reacquire it from the document"
            )
        document = _document_for_comments_part(self._comments_part)
        if document is not None:
            from docx.commentops import _preflight_comment_add

            _refuse_document_protection(document, "add a comment")
            _preflight_comment_add(document)

        def add() -> Comment:
            comment_elm = self._comments_elm.add_comment()
            comment_elm.author = author
            comment_elm.initials = initials
            comment_elm.date = dt.datetime.now(dt.timezone.utc)
            comment = Comment(comment_elm, self._comments_part)
            if text == "":
                return comment
            para_text_iter = iter(text.split("\n"))
            comment.paragraphs[0].add_run(next(para_text_iter))
            for paragraph_text in para_text_iter:
                comment.add_paragraph(text=paragraph_text)
            return comment

        if document is not None:
            with rollback_on_error(document, self):
                return add()
        with rollback_xml_on_error(self._comments_elm):
            return add()

    def get(self, comment_id: int) -> Comment | None:
        """Return the comment identified by `comment_id`, or `None` if not found."""
        comment_elm = self._comments_elm.get_comment_by_id(comment_id)
        return Comment(comment_elm, self._comments_part) if comment_elm is not None else None


class Comment(BlockItemContainer):
    """Proxy for a single comment in the document.

    Provides methods to access comment metadata such as author, initials, and date.

    A comment is also a block-item container, similar to a table cell, so it can contain both
    paragraphs and tables and its paragraphs can contain rich text, hyperlinks and images,
    although the common case is that a comment contains a single paragraph of plain text like a
    sentence or phrase.

    Note that certain content like tables may not be displayed in the Word comment sidebar due to
    space limitations. Such "over-sized" content can still be viewed in the review pane.
    """

    def __init__(self, comment_elm: CT_Comment, comments_part: CommentsPart):
        super().__init__(comment_elm, comments_part)
        self._comment_elm = comment_elm

    def _validate_live(self, operation: str):
        comments_part = cast("CommentsPart", self._parent)
        document = _document_for_comments_part(comments_part)
        if document is None:
            return None
        if self._comment_elm.getparent() is not comments_part._element:
            raise TargetNotFoundError(
                "comment proxy is stale: reacquire it from document.comments"
            )
        _refuse_document_protection(document, operation)
        return document

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the content in this container.

        The paragraph has `text` in a single run if present, and is given paragraph style `style`.
        When `style` is `None` or ommitted, the "CommentText" paragraph style is applied, which is
        the default style for comments.
        """
        document = self._validate_live("edit a comment")
        previous_last = None
        if document is not None:
            from docx.commentops import (
                _last_paragraph,
                _preflight_comment_add,
                _preflight_comments_extended_write,
            )

            _preflight_comment_add(document)
            _preflight_comments_extended_write(document)
            previous_last = _last_paragraph(self._comment_elm)

        def add() -> Paragraph:
            paragraph = super(Comment, self).add_paragraph(text, style)
            # Assign directly because paragraph.style raises when the style is absent.
            if style is None:
                paragraph._p.style = "CommentText"
            if document is not None:
                from docx.commentops import _migrate_comment_extension

                assert previous_last is not None
                _migrate_comment_extension(
                    document, self._comment_elm, previous_last
                )
            return paragraph

        if document is not None:
            with rollback_on_error(document, self):
                return add()
        with rollback_xml_on_error(self._comment_elm):
            return add()

    def add_table(self, rows: int, cols: int, width: Length) -> Table:
        """Return a table appended to this comment."""
        document = self._validate_live("edit a comment")
        previous_last = None
        if document is not None:
            from docx.commentops import (
                _last_paragraph,
                _preflight_comment_add,
                _preflight_comments_extended_write,
            )

            _preflight_comment_add(document)
            _preflight_comments_extended_write(document)
            previous_last = _last_paragraph(self._comment_elm)

        def add() -> Table:
            table = super(Comment, self).add_table(rows, cols, width)
            if document is not None:
                from docx.commentops import _migrate_comment_extension

                assert previous_last is not None
                _migrate_comment_extension(
                    document, self._comment_elm, previous_last
                )
            return table

        if document is not None:
            with rollback_on_error(document, self):
                return add()
        with rollback_xml_on_error(self._comment_elm):
            return add()

    @property
    def author(self) -> str:
        """Read/write. The recorded author of this comment.

        This field is required but can be set to the empty string.
        """
        return self._comment_elm.author

    @author.setter
    def author(self, value: str):
        self._validate_live("edit a comment")
        with rollback_xml_on_error(self._comment_elm):
            self._comment_elm.author = value

    @property
    def comment_id(self) -> int:
        """The unique identifier of this comment."""
        return self._comment_elm.id

    @property
    def initials(self) -> str | None:
        """Read/write. The recorded initials of the comment author.

        This attribute is optional in the XML, returns `None` if not set. Assigning `None` removes
        any existing initials from the XML.
        """
        return self._comment_elm.initials

    @initials.setter
    def initials(self, value: str | None):
        self._validate_live("edit a comment")
        with rollback_xml_on_error(self._comment_elm):
            self._comment_elm.initials = value

    @property
    def text(self) -> str:
        """The text content of this comment as a string.

        Only content in paragraphs is included and of course all emphasis and styling is stripped.

        Paragraph boundaries are indicated with a newline (`"\\\\n"`)
        """
        return "\n".join(p.text for p in self.paragraphs)

    @property
    def timestamp(self) -> dt.datetime | None:
        """The date and time this comment was authored.

        This attribute is optional in the XML, returns `None` if not set.
        """
        return self._comment_elm.date
