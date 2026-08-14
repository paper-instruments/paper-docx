"""Missing-table APIs: comments, lock, pictures, links, captions, notes."""

from __future__ import annotations

import io
import struct
from pathlib import Path

import pytest

import docx
from docx.commentops import delete_comment
from docx.composition import append_document
from docx.controls import get_control, set_control_value
from docx.drawing import Drawing
from docx.errors import DocumentProtectedError
from docx.fields import add_caption
from docx.links import add_hyperlink
from docx.notes import add_endnote, add_footnote
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.protection import set_protection
from docx.search import find_one

from .harness.contract import save_and_reopen
from .harness.paths import fixture_path

MINIMAL = "generated/minimal-clean/minimal.docx"
STORE_ID = "{11111111-1111-1111-1111-111111111111}"


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


class DescribeProtectionSetter:
    def it_locks_the_file_for_the_recipient(self):
        document = _doc()
        status = set_protection(document, edit="readOnly")
        assert status.edit == "readOnly"
        assert status.enforced
        with pytest.raises(DocumentProtectedError):
            find_one(document, "perfectly ordinary").replace("nope")


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


class DescribeDataBoundControls:
    def it_writes_the_custom_xml_store(self, tmp_path: Path):
        document = _doc()
        item = parse_xml(
            '<ns0:root xmlns:ns0="http://example.com/form">'
            "<ns0:name>old</ns0:name></ns0:root>"
        )
        item_part = XmlPart(
            PackURI("/customXml/item2.xml"),
            "application/xml",
            item,
            document.part.package,
        )
        props = parse_xml(
            '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/'
            'officeDocument/2006/customXml"'
            f' ds:itemID="{STORE_ID}"/>'
        )
        props_part = XmlPart(
            PackURI("/customXml/itemProps2.xml"),
            CT.OFC_CUSTOM_XML_PROPERTIES,
            props,
            document.part.package,
        )
        document.part.relate_to(item_part, RT.CUSTOM_XML)
        item_part.relate_to(props_part, RT.CUSTOM_XML_PROPS)
        body = document.element.body
        body.insert(
            len(body) - 1,
            parse_xml(
                '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:sdt><w:sdtPr><w:tag w:val=\"bound\"/>"
                "<w:dataBinding"
                " w:prefixMappings=\"xmlns:ns0='http://example.com/form'\""
                ' w:xpath="/ns0:root/ns0:name"'
                f' w:storeItemID="{STORE_ID}"/></w:sdtPr>'
                "<w:sdtContent><w:r><w:t>old</w:t></w:r></w:sdtContent>"
                "</w:sdt></w:p>"
            ),
        )
        set_control_value(document, "new", tag="bound")
        assert get_control(document, tag="bound").value == "new"
        reopened = save_and_reopen(document, tmp_path / "out.docx")
        store = None
        for part in reopened.part.package.iter_parts():
            if str(part.partname) == "/customXml/item2.xml":
                element = getattr(part, "_element", None)
                store = element if element is not None else parse_xml(part.blob)
        assert store is not None
        ns = {"ns0": "http://example.com/form"}
        assert store.xpath("/ns0:root/ns0:name", namespaces=ns)[0].text == "new"


class DescribeSourceLetterhead:
    def it_copies_source_headers_when_requested(self, tmp_path: Path):
        destination = _doc()
        source = _doc()
        source.sections[0].header.paragraphs[0].text = "Source letterhead"
        append_document(destination, source, headers="source")
        reopened = save_and_reopen(destination, tmp_path / "out.docx")
        assert "Source letterhead" in reopened.sections[-1].header.paragraphs[0].text
