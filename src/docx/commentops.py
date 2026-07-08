"""Comment thread operations: anchored text, replies, resolution (v0.1 V4).

Word models threading and resolution OUTSIDE `word/comments.xml`, in the
`w15` extension part `word/commentsExtended.xml`: one `w15:commentEx` per
comment (keyed by the `w14:paraId` of the comment's LAST paragraph) carrying
`w15:done` (resolved) and `w15:paraIdParent` (reply-of). This module reads
and writes that machinery on top of the upstream v1.2.0 comment support.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, List, Optional, Tuple

from docx import _clock
from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.protection import _refuse_if_protected

_W_DECL = nsdecls("w")

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.comments import Comment
    from docx.document import Document

_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
_COMMENT_EX = f"{{{_W15_NS}}}commentEx"
_PARA_ID_ATTR = qn("w14:paraId")
_PARA_ID = f"{{{_W15_NS}}}paraId"
_PARA_ID_PARENT = f"{{{_W15_NS}}}paraIdParent"
_DONE = f"{{{_W15_NS}}}done"
_P = qn("w:p")
_R = qn("w:r")
_W_ID = qn("w:id")
_RANGE_START = qn("w:commentRangeStart")
_RANGE_END = qn("w:commentRangeEnd")
_REFERENCE = qn("w:commentReference")
_T = qn("w:t")
_DEL_TEXT = qn("w:delText")

COMMENTS_EXTENDED_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml"
    ".commentsExtended+xml"
)
COMMENTS_EXTENDED_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)

_COMMENTS_EX_TEMPLATE = (
    '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>'
)


def _comments_part(document: "Document"):
    return document.part._comments_part  # noqa: SLF001 - same-package broker


def _comment_element(document: "Document", comment: "Comment") -> "_Element":
    return comment._comment_elm  # noqa: SLF001 - same-package access


def _last_paragraph(comment_elm: "_Element") -> "_Element":
    paragraphs = comment_elm.findall(_P)
    if not paragraphs:
        raise UnsupportedStructureError("comment has no paragraph to key on")
    return paragraphs[-1]


def _existing_para_ids(comments_root: "_Element") -> "List[int]":
    values = []
    for paragraph in comments_root.iter(_P):
        raw = paragraph.get(_PARA_ID_ATTR)
        if raw:
            try:
                values.append(int(raw, 16))
            except ValueError:
                continue
    return values


def _ensure_para_id(document: "Document", paragraph: "_Element") -> str:
    existing = paragraph.get(_PARA_ID_ATTR)
    if existing:
        return existing
    comments_root = _comments_part(document)._element  # noqa: SLF001
    next_value = max(_existing_para_ids(comments_root), default=0x10000000) + 1
    para_id = f"{next_value:08X}"
    paragraph.set(_PARA_ID_ATTR, para_id)
    return para_id


def _comments_extended_root(document: "Document", *, create: bool) -> "Optional[_Element]":
    part = document.part
    try:
        extended = part.part_related_by(COMMENTS_EXTENDED_RELATIONSHIP_TYPE)
    except KeyError:
        extended = None
    if extended is not None:
        element = getattr(extended, "_element", None)
        if element is None:  # pragma: no cover - registration in docx/__init__
            raise UnsupportedStructureError(
                "commentsExtended part loaded as an opaque blob; the package"
                " registration is missing, and writing to a parsed copy would"
                " be silently lost on save"
            )
        return element
    if not create:
        return None
    from docx.opc.packuri import PackURI
    from docx.opc.part import XmlPart

    root = parse_xml(_COMMENTS_EX_TEMPLATE)
    new_part = XmlPart(
        PackURI("/word/commentsExtended.xml"),
        COMMENTS_EXTENDED_CONTENT_TYPE,
        root,
        part.package,
    )
    part.relate_to(new_part, COMMENTS_EXTENDED_RELATIONSHIP_TYPE)
    return root


def _entry_for(root: "_Element", para_id: str, *, create: bool) -> "Optional[_Element]":
    for entry in root.findall(_COMMENT_EX):
        if entry.get(_PARA_ID) == para_id:
            return entry
    if not create:
        return None
    entry = parse_xml(
        '<w15:commentEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
        f' w15:paraId="{para_id}" w15:done="0"/>'
    )
    root.append(entry)
    return entry


def is_resolved(document: "Document", comment: "Comment") -> bool:
    """Whether `comment` is marked resolved (`w15:done`)."""
    root = _comments_extended_root(document, create=False)
    if root is None:
        return False
    para_id = _last_paragraph(_comment_element(document, comment)).get(_PARA_ID_ATTR)
    if not para_id:
        return False
    entry = _entry_for(root, para_id, create=False)
    return entry is not None and entry.get(_DONE) in ("1", "true")


def resolve(document: "Document", comment: "Comment", *, resolved: bool = True) -> None:
    """Mark `comment` resolved (or reopened) the way Word does."""
    _refuse_if_protected(document, "resolve a comment")
    para_id = _ensure_para_id(document, _last_paragraph(_comment_element(document, comment)))
    root = _comments_extended_root(document, create=True)
    entry = _entry_for(root, para_id, create=True)
    entry.set(_DONE, "1" if resolved else "0")


def parent_of(document: "Document", comment: "Comment") -> Optional[int]:
    """The comment-id this comment replies to, or None for a top-level one."""
    root = _comments_extended_root(document, create=False)
    if root is None:
        return None
    para_id = _last_paragraph(_comment_element(document, comment)).get(_PARA_ID_ATTR)
    if not para_id:
        return None
    entry = _entry_for(root, para_id, create=False)
    parent_para = entry.get(_PARA_ID_PARENT) if entry is not None else None
    if not parent_para:
        return None
    for candidate in document.comments:
        candidate_elm = _comment_element(document, candidate)
        if _last_paragraph(candidate_elm).get(_PARA_ID_ATTR) == parent_para:
            return candidate.comment_id
    return None


def _anchor_elements(document: "Document", comment_id: int):
    body = document.element.body
    start = end = reference_run = None
    for node in body.iter(_RANGE_START, _RANGE_END, _REFERENCE):
        if int(node.get(_W_ID)) != comment_id:
            continue
        if node.tag == _RANGE_START:
            start = node
        elif node.tag == _RANGE_END:
            end = node
        else:
            reference_run = node.getparent()
    return start, end, reference_run


def anchored_text(document: "Document", comment: "Comment") -> str:
    """The document text `comment` is anchored to (its range marks' span)."""
    start, end, _ = _anchor_elements(document, comment.comment_id)
    if start is None or end is None:
        raise TargetNotFoundError(
            f"comment {comment.comment_id} has no range anchor in the body"
        )
    # document-order walk over the whole story tree: comment ranges may
    # cross paragraph boundaries, where sibling iteration would truncate
    pieces: "List[str]" = []
    root = start.getroottree().getroot()
    inside = False
    for node in root.iter():
        if node is start:
            inside = True
            continue
        if node is end:
            break
        if inside and node.tag == _T:
            pieces.append(node.text or "")
    return "".join(pieces)


def reply(
    document: "Document",
    comment: "Comment",
    text: str,
    *,
    author: str,
    initials: Optional[str] = None,
    date: Optional[dt.datetime] = None,
) -> "Comment":
    """Add a threaded reply to `comment`, anchored to the same text range."""
    if not author:
        raise ValueError("author is required")
    _refuse_if_protected(document, "reply to a comment")
    parent_elm = _comment_element(document, comment)
    start, end, reference_run = _anchor_elements(document, comment.comment_id)
    if start is None or end is None or reference_run is None:
        raise TargetNotFoundError(
            f"comment {comment.comment_id} has no range anchor to thread onto"
        )
    parent_para_id = _ensure_para_id(document, _last_paragraph(parent_elm))

    new_comment = document.comments.add_comment(
        text=text, author=author, initials=initials or ""
    )
    new_elm = _comment_element(document, new_comment)
    new_elm.date = date if date is not None else _clock.now()
    new_para_id = _ensure_para_id(document, _last_paragraph(new_elm))

    new_id = new_comment.comment_id
    start.addnext(
        parse_xml(f'<w:commentRangeStart {_W_DECL} w:id="{new_id}"/>')
    )
    end.addprevious(parse_xml(f'<w:commentRangeEnd {_W_DECL} w:id="{new_id}"/>'))
    reference_run.addnext(
        parse_xml(
            f"<w:r {_W_DECL}><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr>"
            f'<w:commentReference w:id="{new_id}"/></w:r>'
        )
    )

    root = _comments_extended_root(document, create=True)
    _entry_for(root, parent_para_id, create=True)
    reply_entry = _entry_for(root, new_para_id, create=True)
    reply_entry.set(_PARA_ID_PARENT, parent_para_id)
    return new_comment


def comment_thread(document: "Document") -> Tuple[dict, ...]:
    """Every comment with its thread state: (id, author, text, resolved,
    parent_id, anchored_text where available)."""
    entries = []
    for comment in document.comments:
        try:
            anchor_text = anchored_text(document, comment)
        except TargetNotFoundError:
            anchor_text = None
        entries.append(
            {
                "comment_id": comment.comment_id,
                "author": comment.author,
                "text": comment.text,
                "resolved": is_resolved(document, comment),
                "parent_id": parent_of(document, comment),
                "anchored_text": anchor_text,
            }
        )
    return tuple(entries)

