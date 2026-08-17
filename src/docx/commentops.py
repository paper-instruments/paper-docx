"""Comment thread operations: anchored text, replies, resolution.

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
from docx._guard import check_install
from docx._ownership import require_comment_owner
from docx._textatoms import DEL_TEXT, INSTR_TEXT, is_direct_run_child, project_run_child
from docx._transaction import rollback_on_error
from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.protection import _refuse_if_protected

_W_DECL = nsdecls("w")

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.comments import Comment
    from docx.document import Document

__all__ = [
    "COMMENTS_EXTENDED_CONTENT_TYPE",
    "COMMENTS_EXTENDED_RELATIONSHIP_TYPE",
    "anchored_text",
    "comment_thread",
    "is_resolved",
    "parent_of",
    "reply",
    "resolve",
]

check_install()

_W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
_COMMENT_EX = f"{{{_W15_NS}}}commentEx"
_PARA_ID_ATTR = qn("w14:paraId")
_PARA_ID = f"{{{_W15_NS}}}paraId"
_PARA_ID_PARENT = f"{{{_W15_NS}}}paraIdParent"
_DONE = f"{{{_W15_NS}}}done"
_COMMENTS = qn("w:comments")
_COMMENT = qn("w:comment")
_P = qn("w:p")
_R = qn("w:r")
_W_ID = qn("w:id")
_RANGE_START = qn("w:commentRangeStart")
_RANGE_END = qn("w:commentRangeEnd")
_REFERENCE = qn("w:commentReference")
_SDT = qn("w:sdt")
_FLD_SIMPLE = qn("w:fldSimple")

COMMENTS_EXTENDED_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
COMMENTS_EXTENDED_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)

_COMMENTS_EX_TEMPLATE = (
    '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>'
)

_COMMENTS_PARTNAME = "/word/comments.xml"
_COMMENTS_EXTENDED_PARTNAME = "/word/commentsExtended.xml"


def _part_with_name(document: "Document", partname: str):
    package = document.part.package
    if package is None:
        return None
    folded = partname.casefold()
    return next(
        (
            part
            for part in package.iter_parts()
            if str(part.partname).casefold() == folded
        ),
        None,
    )


def _preflight_comment_add(document: "Document") -> None:
    """Validate the existing comments collection before upstream mutates it."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        part = document.part.part_related_by(RT.COMMENTS)
    except KeyError:
        if _part_with_name(document, _COMMENTS_PARTNAME) is not None:
            raise UnsupportedStructureError(
                "a package part already occupies /word/comments.xml without"
                " the document comments relationship; nothing was changed"
            )
        return
    except ValueError:
        raise UnsupportedStructureError(
            "multiple comments relationships make comment authoring ambiguous; nothing was changed"
        ) from None

    root = getattr(part, "_element", None)
    if root is None or root.tag != _COMMENTS:
        raise UnsupportedStructureError(
            "the comments relationship does not target a live w:comments part; nothing was changed"
        )
    seen: "set[int]" = set()
    for comment in root.findall(_COMMENT):
        raw = comment.get(_W_ID)
        try:
            if raw is None:
                raise ValueError
            comment_id = int(raw)
        except (TypeError, ValueError):
            raise UnsupportedStructureError(
                f"malformed comment w:id={raw!r}; nothing was changed"
            ) from None
        if comment_id in seen:
            raise UnsupportedStructureError(
                f"duplicate comment w:id={comment_id}; nothing was changed"
            )
        seen.add(comment_id)


def _preflight_comment_range(
    document: "Document", first_run: "_Element", last_run: "_Element", *, operation: str
) -> None:
    """Validate only the selected run interval before adding range markers."""
    _refuse_if_protected(document, operation)
    nodes = tuple(document.element.iter())
    positions = {id(node): index for index, node in enumerate(nodes)}
    first = positions.get(id(first_run))
    last = positions.get(id(last_run))
    if first is None or last is None:
        raise TargetNotFoundError("comment range runs are stale or detached; reacquire the span")
    if first > last:
        raise UnsupportedStructureError("comment range endpoints are reversed; nothing was changed")
    interval = nodes[first : last + 1]
    selected_runs = [node for node in interval if node.tag == _R]
    if any(next(run.iter(qn("w:fldChar")), None) is not None for run in selected_runs):
        raise UnsupportedStructureError(
            "cannot anchor a comment across a complex field boundary; nothing was changed"
        )
    selected_ids = {id(run) for run in selected_runs}
    depth = 0
    for node in nodes:
        if node.tag == qn("w:fldChar"):
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                depth += 1
            elif kind == "end" and depth:
                depth -= 1
        if id(node) in selected_ids and depth:
            raise UnsupportedStructureError(
                "cannot anchor a comment inside a complex field; nothing was changed"
            )
    for run in selected_runs:
        current = run.getparent()
        while current is not None:
            if current.tag == _FLD_SIMPLE:
                raise UnsupportedStructureError(
                    "cannot anchor a comment inside a field; nothing was changed"
                )
            current = current.getparent()
    from docx.controls import _refuse_control_write_restrictions

    controls = []
    for node in interval:
        current = node
        while current is not None:
            if current.tag == _SDT and not any(current is item for item in controls):
                controls.append(current)
            current = current.getparent()
    for control in controls:
        _refuse_control_write_restrictions(control)


def _validate_para_id(raw: "Optional[str]", *, attribute: str) -> str:
    if (
        raw is None
        or len(raw) != 8
        or any(character not in "0123456789abcdefABCDEF" for character in raw)
    ):
        raise UnsupportedStructureError(f"malformed {attribute}={raw!r}; nothing was changed")
    return raw.upper()


def _comments_part(document: "Document"):
    return document.part._comments_part  # noqa: SLF001 - same-package broker


def _comment_element(document: "Document", comment: "Comment") -> "_Element":
    require_comment_owner(document, comment)
    return comment._comment_elm  # noqa: SLF001 - same-package access


def _last_paragraph(comment_elm: "_Element") -> "_Element":
    paragraphs = list(comment_elm.iter(_P))
    if not paragraphs:
        raise UnsupportedStructureError("comment has no paragraph to key on")
    return paragraphs[-1]


def _existing_para_ids(comments_root: "_Element") -> "List[int]":
    values = []
    for paragraph in comments_root.iter(_P):
        raw = paragraph.get(_PARA_ID_ATTR)
        if raw is not None:
            values.append(int(_validate_para_id(raw, attribute="w14:paraId"), 16))
    if len(values) != len(set(values)):
        raise UnsupportedStructureError(
            "duplicate w14:paraId values make comment threading ambiguous; nothing was changed"
        )
    return values


def _ensure_para_id(document: "Document", paragraph: "_Element") -> str:
    existing = paragraph.get(_PARA_ID_ATTR)
    if existing is not None:
        return _validate_para_id(existing, attribute="w14:paraId")
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
    except ValueError:
        raise UnsupportedStructureError(
            "multiple commentsExtended relationships make comment threading"
            " ambiguous; nothing was changed"
        ) from None
    if extended is not None:
        element = getattr(extended, "_element", None)
        if element is None:  # pragma: no cover - registration in docx/__init__
            raise UnsupportedStructureError(
                "commentsExtended part loaded as an opaque blob; the package"
                " registration is missing, and writing to a parsed copy would"
                " be silently lost on save"
            )
        if element.tag != f"{{{_W15_NS}}}commentsEx":
            raise UnsupportedStructureError(
                "commentsExtended has an unexpected root element; nothing was changed"
            )
        seen: "set[str]" = set()
        for entry in element.findall(_COMMENT_EX):
            para_id = _validate_para_id(entry.get(_PARA_ID), attribute="w15:paraId")
            parent_id = entry.get(_PARA_ID_PARENT)
            if parent_id is not None:
                _validate_para_id(parent_id, attribute="w15:paraIdParent")
            normalized = para_id.upper()
            if normalized in seen:
                raise UnsupportedStructureError(
                    "duplicate w15:paraId values make comment threading"
                    " ambiguous; nothing was changed"
                )
            seen.add(normalized)
        return element
    if not create:
        return None
    if _part_with_name(document, _COMMENTS_EXTENDED_PARTNAME) is not None:
        raise UnsupportedStructureError(
            "a package part already occupies /word/commentsExtended.xml"
            " without the commentsExtended relationship; nothing was changed"
        )
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


def _preflight_comments_extended_write(document: "Document") -> None:
    """Validate extension state and any partname needed for its creation."""
    root = _comments_extended_root(document, create=False)
    if root is None and _part_with_name(document, _COMMENTS_EXTENDED_PARTNAME) is not None:
        raise UnsupportedStructureError(
            "a package part already occupies /word/commentsExtended.xml"
            " without the commentsExtended relationship; nothing was changed"
        )


def _migrate_comment_extension(
    document: "Document",
    comment_elm: "_Element",
    previous_last: "_Element",
) -> None:
    """Move thread and resolution state when a comment gains a new last paragraph."""
    current_last = _last_paragraph(comment_elm)
    if current_last is previous_last:
        return
    previous_raw = previous_last.get(_PARA_ID_ATTR)
    if previous_raw is None:
        return
    previous_id = _validate_para_id(previous_raw, attribute="w14:paraId")
    root = _comments_extended_root(document, create=False)
    if root is None:
        return
    own_entry = _entry_for(root, previous_id, create=False)
    child_entries = [
        entry
        for entry in root.findall(_COMMENT_EX)
        if entry.get(_PARA_ID_PARENT) is not None
        and _validate_para_id(
            entry.get(_PARA_ID_PARENT), attribute="w15:paraIdParent"
        )
        == previous_id
    ]
    if own_entry is None and not child_entries:
        return
    current_id = _ensure_para_id(document, current_last)
    if own_entry is not None:
        own_entry.set(_PARA_ID, current_id)
    for entry in child_entries:
        entry.set(_PARA_ID_PARENT, current_id)


def _entry_for(root: "_Element", para_id: str, *, create: bool) -> "Optional[_Element]":
    normalized = _validate_para_id(para_id, attribute="paragraph id")
    for entry in root.findall(_COMMENT_EX):
        existing = _validate_para_id(
            entry.get(_PARA_ID), attribute="w15:paraId"
        )
        if existing == normalized:
            return entry
    if not create:
        return None
    entry = parse_xml(
        '<w15:commentEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
        f' w15:paraId="{normalized}" w15:done="0"/>'
    )
    root.append(entry)
    return entry


def is_resolved(document: "Document", comment: "Comment") -> bool:
    """Whether `comment` is marked resolved (`w15:done`)."""
    comment_elm = _comment_element(document, comment)
    root = _comments_extended_root(document, create=False)
    if root is None:
        return False
    para_id = _last_paragraph(comment_elm).get(_PARA_ID_ATTR)
    if not para_id:
        return False
    entry = _entry_for(root, para_id, create=False)
    return entry is not None and entry.get(_DONE) in ("1", "true")


def resolve(document: "Document", comment: "Comment", *, resolved: bool = True) -> None:
    """Mark `comment` resolved (or reopened) the way Word does."""
    comment_elm = _comment_element(document, comment)
    _refuse_if_protected(document, "resolve a comment")
    # Validate a pre-existing extension part before adding a paraId to the
    # comment. An opaque extension must refuse without touching comments.xml.
    _preflight_comments_extended_write(document)
    with rollback_on_error(document):
        para_id = _ensure_para_id(document, _last_paragraph(comment_elm))
        root = _comments_extended_root(document, create=True)
        entry = _entry_for(root, para_id, create=True)
        entry.set(_DONE, "1" if resolved else "0")


def parent_of(document: "Document", comment: "Comment") -> Optional[int]:
    """The comment-id this comment replies to, or None for a top-level one."""
    comment_elm = _comment_element(document, comment)
    root = _comments_extended_root(document, create=False)
    if root is None:
        return None
    para_id = _last_paragraph(comment_elm).get(_PARA_ID_ATTR)
    if not para_id:
        return None
    entry = _entry_for(root, para_id, create=False)
    raw_parent_para = entry.get(_PARA_ID_PARENT) if entry is not None else None
    if not raw_parent_para:
        return None
    parent_para = _validate_para_id(
        raw_parent_para, attribute="w15:paraIdParent"
    )
    for candidate in document.comments:
        candidate_elm = _comment_element(document, candidate)
        raw_candidate = _last_paragraph(candidate_elm).get(_PARA_ID_ATTR)
        if raw_candidate is not None and _validate_para_id(
            raw_candidate, attribute="w14:paraId"
        ) == parent_para:
            return candidate.comment_id
    return None


def _anchor_elements(document: "Document", comment_id: int):
    body = document.element.body
    starts = []
    ends = []
    reference_runs = []
    for node in body.iter(_RANGE_START, _RANGE_END, _REFERENCE):
        raw = node.get(_W_ID)
        try:
            if raw is None:
                raise ValueError
            marker_id = int(raw)
        except (TypeError, ValueError):
            name = node.tag.rsplit("}", 1)[-1]
            raise UnsupportedStructureError(
                f"malformed {name} comment marker w:id={raw!r}; nothing was changed"
            ) from None
        if marker_id != comment_id:
            continue
        if node.tag == _RANGE_START:
            starts.append(node)
        elif node.tag == _RANGE_END:
            ends.append(node)
        else:
            reference_runs.append(node.getparent())
    if any(len(markers) > 1 for markers in (starts, ends, reference_runs)):
        raise UnsupportedStructureError(
            f"comment {comment_id} has duplicate anchor markers; nothing was changed"
        )
    return (
        starts[0] if starts else None,
        ends[0] if ends else None,
        reference_runs[0] if reference_runs else None,
    )


def anchored_text(document: "Document", comment: "Comment") -> str:
    """The document text `comment` is anchored to (its range marks' span)."""
    _comment_element(document, comment)
    start, end, _ = _anchor_elements(document, comment.comment_id)
    if start is None or end is None:
        raise TargetNotFoundError(f"comment {comment.comment_id} has no range anchor in the body")
    # document-order walk over the whole story tree: comment ranges may
    # cross paragraph boundaries, where sibling iteration would truncate
    pieces: "List[str]" = []
    root = start.getroottree().getroot()
    inside = False
    last_paragraph = None
    for node in root.iter():
        if node is start:
            inside = True
            continue
        if node is end:
            break
        if inside and is_direct_run_child(node) and node.tag not in (DEL_TEXT, INSTR_TEXT):
            projection = project_run_child(node)
            if projection.barrier or not projection.text:
                continue
            paragraph = node.getparent()
            while paragraph is not None and paragraph.tag != qn("w:p"):
                paragraph = paragraph.getparent()
            if last_paragraph is not None and paragraph is not last_paragraph:
                pieces.append("\n")
            last_paragraph = paragraph
            pieces.append(projection.text)
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
    from docx.search import _validate_xml_characters

    _validate_xml_characters(text, argument="text")
    _validate_xml_characters(author, argument="author")
    if initials is not None:
        _validate_xml_characters(initials, argument="initials")
    if date is not None and not isinstance(date, dt.datetime):
        raise TypeError("date must be a datetime or None")
    parent_elm = _comment_element(document, comment)
    _refuse_if_protected(document, "reply to a comment")
    start, end, reference_run = _anchor_elements(document, comment.comment_id)
    if start is None or end is None or reference_run is None:
        raise TargetNotFoundError(
            f"comment {comment.comment_id} has no range anchor to thread onto"
        )
    ordered = tuple(document.element.iter())
    positions = {id(node): index for index, node in enumerate(ordered)}
    start_position = positions.get(id(start))
    end_position = positions.get(id(end))
    if (
        start_position is None
        or end_position is None
        or start_position >= end_position
    ):
        raise UnsupportedStructureError(
            "comment range markers are stale, reversed, or detached; nothing was changed"
        )
    anchored_runs = [
        node
        for node in ordered[start_position + 1 : end_position]
        if node.tag == _R
    ]
    if not anchored_runs:
        raise UnsupportedStructureError(
            "comment range contains no live runs to anchor a reply; nothing was changed"
        )
    _preflight_comment_range(
        document,
        anchored_runs[0],
        anchored_runs[-1],
        operation="reply to a comment",
    )
    _preflight_comment_add(document)
    _preflight_comments_extended_write(document)
    with rollback_on_error(document):
        parent_para_id = _ensure_para_id(document, _last_paragraph(parent_elm))

        new_comment = document.comments.add_comment(
            text=text, author=author, initials=initials or ""
        )
        new_elm = _comment_element(document, new_comment)
        new_elm.date = date if date is not None else _clock.now()
        new_para_id = _ensure_para_id(document, _last_paragraph(new_elm))

        new_id = new_comment.comment_id
        start.addnext(parse_xml(f'<w:commentRangeStart {_W_DECL} w:id="{new_id}"/>'))
        end.addprevious(parse_xml(f'<w:commentRangeEnd {_W_DECL} w:id="{new_id}"/>'))
        reference_run.addnext(
            parse_xml(
                f'<w:r {_W_DECL}><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
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
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    _preflight_comment_add(document)
    try:
        comments_part = document.part.part_related_by(RT.COMMENTS)
    except KeyError:
        return ()
    except ValueError:
        raise UnsupportedStructureError(
            "multiple comments relationships make comment inspection ambiguous; nothing was changed"
        ) from None
    comments = comments_part.comments
    entries = []
    for comment in comments:
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
