"""Missing-table APIs: comments, lock, pictures, links, captions, notes."""

from __future__ import annotations

import io
import struct
from pathlib import Path

import pytest

import docx
from docx.commentops import COMMENTS_IDS_RELATIONSHIP_TYPE, delete_comment
from docx.drawing import Drawing
from docx.errors import (
    BoundaryViolationError,
    DocumentProtectedError,
    TargetNotFoundError,
    UnsupportedStructureError,
)
from docx.fields import add_caption
from docx.links import add_hyperlink
from docx.notes import add_endnote, add_footnote
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement, parse_xml
from docx.protection import acknowledge_protection, set_protection
from docx.search import find_one

from .harness.contract import save_and_reopen
from .harness.paths import fixture_path

MINIMAL = "generated/minimal-clean/minimal.docx"


def _doc(relpath: str = MINIMAL):
    return docx.Document(str(fixture_path(relpath)))


def _bmp(red: int, green: int, blue: int) -> bytes:
    return (
        struct.pack("<2sIHHI", b"BM", 58, 0, 0, 54)
        + struct.pack("<IiiHHIIiiII", 40, 1, 1, 1, 24, 0, 4, 2835, 2835, 0, 0)
        + bytes((blue, green, red, 0))
    )


def _blip_embed(drawing: Drawing) -> str:
    return drawing._drawing.xpath(".//pic:blipFill/a:blip/@r:embed")[0]


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


class DescribeProtectionSetter:
    def it_locks_the_file_for_the_recipient(self):
        document = _doc()
        status = set_protection(document, edit="readOnly")
        assert status.edit == "readOnly"
        assert status.enforced
        with pytest.raises(DocumentProtectedError):
            find_one(document, "perfectly ordinary").replace("nope")

    def it_clears_an_in_memory_ack_when_locking_again(self):
        document = _doc()
        set_protection(document, edit="readOnly")
        acknowledge_protection(document)
        find_one(document, "perfectly ordinary").replace("ok")
        set_protection(document, edit="comments")
        with pytest.raises(DocumentProtectedError):
            find_one(document, "ok").replace("nope")

    def it_inserts_protection_before_later_settings_children(self):
        document = _doc()
        settings = document.settings.element
        for child in list(settings):
            if child.tag == qn("w:defaultTabStop"):
                settings.remove(child)
        footnote_pr = OxmlElement("w:footnotePr")
        compat = settings.find(qn("w:compat"))
        if compat is not None:
            compat.addprevious(footnote_pr)
        else:
            settings.append(footnote_pr)
        set_protection(document, edit="readOnly")
        tags = [child.tag for child in settings]
        assert tags.index(qn("w:documentProtection")) < tags.index(qn("w:footnotePr"))


class DescribePictureReplace:
    def it_swaps_bytes_without_sharing_the_part(self):
        document = _doc()
        run = document.add_paragraph().add_run()
        run.add_picture(io.BytesIO(_bmp(255, 0, 0)))
        other = document.add_paragraph().add_run()
        other.add_picture(io.BytesIO(_bmp(255, 0, 0)))
        drawing = next(
            item for item in run.iter_inner_content() if isinstance(item, Drawing)
        )
        other_drawing = next(
            item for item in other.iter_inner_content() if isinstance(item, Drawing)
        )
        shared = _blip_embed(drawing)
        assert _blip_embed(other_drawing) == shared
        extent = drawing._drawing.xpath(".//wp:extent")[0]
        cx, cy = extent.get("cx"), extent.get("cy")
        drawing.replace_picture(io.BytesIO(_bmp(0, 0, 255)))
        assert _blip_embed(drawing) != shared
        assert _blip_embed(other_drawing) == shared
        assert extent.get("cx") == cx
        assert extent.get("cy") == cy

    def it_refuses_a_linked_picture(self):
        document = _doc()
        run = document.add_paragraph().add_run()
        run.add_picture(io.BytesIO(_bmp(255, 0, 0)))
        drawing = next(
            item for item in run.iter_inner_content() if isinstance(item, Drawing)
        )
        blip = drawing._drawing.xpath(".//pic:blipFill/a:blip")[0]
        blip.set(qn("r:link"), "rId99")
        with pytest.raises(UnsupportedStructureError, match="linked picture"):
            drawing.replace_picture(io.BytesIO(_bmp(0, 0, 255)))


class DescribeHyperlinks:
    def it_wraps_a_phrase_and_can_retarget(self, tmp_path: Path):
        document = _doc()
        span = find_one(document, "perfectly ordinary")
        link = add_hyperlink(document, span, "https://example.com/a")
        assert link.address == "https://example.com/a"
        link.address = "https://example.com/b"
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        addresses = [
            hyperlink.address
            for paragraph in reopened.paragraphs
            for hyperlink in paragraph.hyperlinks
        ]
        assert "https://example.com/b" in addresses

    def it_puts_the_relationship_on_the_header_part(self, tmp_path: Path):
        document = _doc()
        header = document.sections[0].header
        header.paragraphs[0].text = "HeaderLinkUnique"
        span = find_one(document, "HeaderLinkUnique", story="word/header1.xml")
        link = add_hyperlink(document, span, "https://example.com/header")
        header_part = document.sections[0].header.part
        r_id = link._hyperlink.get(qn("r:id"))
        assert r_id in header_part.rels
        assert header_part.rels[r_id].reltype == RT.HYPERLINK
        assert header_part.rels[r_id].target_ref == "https://example.com/header"
        assert link.address == "https://example.com/header"
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        addresses = [
            hyperlink.address
            for paragraph in reopened.sections[0].header.paragraphs
            for hyperlink in paragraph.hyperlinks
        ]
        assert "https://example.com/header" in addresses

    def it_refuses_to_nest_a_hyperlink(self):
        document = _doc()
        span = find_one(document, "perfectly ordinary")
        add_hyperlink(document, span, "https://example.com/a")
        again = find_one(document, "perfectly ordinary")
        with pytest.raises(UnsupportedStructureError, match="already inside a hyperlink"):
            add_hyperlink(document, again, "https://example.com/b")

    def it_refuses_intervening_markup_between_runs(self):
        document = _doc()
        paragraph = document.add_paragraph()
        first = paragraph.add_run("hello ")
        second = paragraph.add_run("world")
        first._r.addnext(
            parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="1" w:name="Term"/>')
        )
        second._r.addnext(parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="1"/>'))
        before = paragraph._p.xml
        with pytest.raises(UnsupportedStructureError, match="intervening markup"):
            add_hyperlink(document, find_one(document, "hello world"), "https://example.com/a")
        assert paragraph._p.xml == before

    def it_clears_a_stale_internal_anchor_when_retargeting(self):
        document = _doc()
        span = find_one(document, "perfectly ordinary")
        link = add_hyperlink(document, span, "https://example.com/a")
        link._hyperlink.anchor = "_TocOld"
        link.address = "https://example.com/b"
        assert link._hyperlink.anchor is None
        assert link.address == "https://example.com/b"

    def it_refuses_a_span_from_another_document(self):
        document = _doc()
        foreign = find_one(_doc(), "perfectly ordinary")
        with pytest.raises(BoundaryViolationError, match="different document"):
            add_hyperlink(document, foreign, "https://example.com/a")

    def it_refuses_a_stale_span(self):
        document = _doc()
        span = find_one(document, "perfectly ordinary")
        find_one(document, "perfectly ordinary").replace("changed")
        with pytest.raises(TargetNotFoundError, match="stale"):
            add_hyperlink(document, span, "https://example.com/a")

    def it_keeps_a_shared_relationship_when_retargeting_one(self):
        document = _doc()
        first = add_hyperlink(
            document, find_one(document, "perfectly ordinary"), "https://example.com/a"
        )
        second = add_hyperlink(
            document, find_one(document, "First body paragraph"), "https://example.com/a"
        )
        assert first._hyperlink.get(qn("r:id")) == second._hyperlink.get(qn("r:id"))
        first.address = "https://example.com/b"
        assert first.address == "https://example.com/b"
        assert second.address == "https://example.com/a"

    def it_refuses_a_field_result_span(self):
        document = _doc("generated/feature-isolated/fields.docx")
        span = find_one(document, "June 1, 2026")
        assert span.in_field
        with pytest.raises(UnsupportedStructureError, match="field result"):
            add_hyperlink(document, span, "https://example.com/a")

    def it_defines_the_hyperlink_character_style(self):
        document = _doc()
        assert "Hyperlink" not in document.styles
        add_hyperlink(document, find_one(document, "perfectly ordinary"), "https://example.com/a")
        assert "Hyperlink" in document.styles

    def it_refuses_a_data_bound_span(self):
        document = _doc()
        document.element.body.insert(
            len(document.element.body) - 1,
            parse_xml(
                f'<w:p {nsdecls("w")}>'
                '<w:sdt><w:sdtPr><w:tag w:val="bound"/>'
                '<w:dataBinding w:xpath="/x" w:storeItemID="{11111111-1111-1111-1111-111111111111}"/>'
                "</w:sdtPr>"
                "<w:sdtContent><w:r><w:t>BoundLink</w:t></w:r></w:sdtContent>"
                "</w:sdt></w:p>"
            ),
        )
        with pytest.raises(UnsupportedStructureError, match="data-bound"):
            add_hyperlink(document, find_one(document, "BoundLink"), "https://example.com/a")

    def it_refuses_a_locked_control_span(self):
        document = _doc()
        document.element.body.insert(
            len(document.element.body) - 1,
            parse_xml(
                f'<w:p {nsdecls("w")}>'
                '<w:sdt><w:sdtPr><w:tag w:val="locked"/>'
                '<w:lock w:val="contentLocked"/></w:sdtPr>'
                "<w:sdtContent><w:r><w:t>LockedLink</w:t></w:r></w:sdtContent>"
                "</w:sdt></w:p>"
            ),
        )
        with pytest.raises(UnsupportedStructureError, match="locked"):
            add_hyperlink(document, find_one(document, "LockedLink"), "https://example.com/a")

    def it_wraps_text_inside_an_unlocked_control(self):
        document = _doc()
        document.element.body.insert(
            len(document.element.body) - 1,
            parse_xml(
                f'<w:p {nsdecls("w")}>'
                '<w:sdt><w:sdtPr><w:tag w:val="plain"/></w:sdtPr>'
                "<w:sdtContent><w:r><w:t>PlainLink</w:t></w:r></w:sdtContent>"
                "</w:sdt></w:p>"
            ),
        )
        link = add_hyperlink(
            document, find_one(document, "PlainLink"), "https://example.com/a"
        )
        assert link.address == "https://example.com/a"

    def it_refuses_a_plain_text_control_span(self):
        document = _doc()
        document.element.body.insert(
            len(document.element.body) - 1,
            parse_xml(
                f'<w:p {nsdecls("w")}>'
                '<w:sdt><w:sdtPr><w:tag w:val="plain-text"/><w:text/></w:sdtPr>'
                "<w:sdtContent><w:r><w:t>PlainTextLink</w:t></w:r></w:sdtContent>"
                "</w:sdt></w:p>"
            ),
        )
        with pytest.raises(UnsupportedStructureError, match="plain-text"):
            add_hyperlink(
                document, find_one(document, "PlainTextLink"), "https://example.com/a"
            )


class DescribeCaptions:
    def it_inserts_a_seq_field(self):
        document = _doc()
        paragraph = document.add_paragraph()
        add_caption(paragraph, label="Figure", description="Diagram")
        xml = paragraph._p.xml
        assert "SEQ Figure" in xml
        assert paragraph._p.style == "Caption"


class DescribeNotes:
    def it_adds_a_footnote_and_endnote(self, tmp_path: Path):
        document = _doc()
        span = find_one(document, "perfectly ordinary")
        footnote_id = add_footnote(document, span, "A footnote.")
        endnote_id = add_endnote(document, span, "An endnote.")
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        footnotes = reopened.part.part_related_by(RT.FOOTNOTES)._element
        endnotes = reopened.part.part_related_by(RT.ENDNOTES)._element
        assert str(footnote_id) in {
            note.get(qn("w:id")) for note in footnotes.findall(qn("w:footnote"))
        }
        assert str(endnote_id) in {
            note.get(qn("w:id")) for note in endnotes.findall(qn("w:endnote"))
        }

    def it_keeps_ampersands_and_angles_as_text(self, tmp_path: Path):
        document = _doc()
        span = find_one(document, "perfectly ordinary")
        add_footnote(document, span, "A & B <C>")
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        footnotes = reopened.part.part_related_by(RT.FOOTNOTES)._element
        texts = [node.text or "" for node in footnotes.findall(".//" + qn("w:t"))]
        assert any("A & B <C>" in text for text in texts)

    def it_refuses_a_note_outside_the_body(self):
        document = _doc()
        document.sections[0].header.paragraphs[0].text = "HeaderNoteUnique"
        span = find_one(document, "HeaderNoteUnique", story="word/header1.xml")
        with pytest.raises(UnsupportedStructureError, match="main document body"):
            add_footnote(document, span, "nope")

    def it_refuses_a_span_from_another_document(self):
        document = _doc()
        foreign = find_one(_doc(), "perfectly ordinary")
        with pytest.raises(BoundaryViolationError, match="different document"):
            add_footnote(document, foreign, "nope")

    def it_refuses_a_stale_span(self):
        document = _doc()
        span = find_one(document, "perfectly ordinary")
        find_one(document, "perfectly ordinary").replace("changed")
        with pytest.raises(TargetNotFoundError, match="stale"):
            add_footnote(document, span, "nope")
