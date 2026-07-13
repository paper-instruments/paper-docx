"""Initialize `docx` package.

Export the `Document` constructor function and establish the mapping of part-type to
the part-classe that implements that type.
"""

# ruff: noqa: E402

from __future__ import annotations

from typing import TYPE_CHECKING, Type

from docx._version import __paper_version__, assert_distribution_identity

assert_distribution_identity()

from docx.api import Document

if TYPE_CHECKING:
    from docx.opc.part import Part

__version__ = "1.2.0"


__all__ = ["Document", "__paper_version__"]


# -- register custom Part classes with opc package reader --

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import PartFactory
from docx.opc.parts.coreprops import CorePropertiesPart
from docx.parts.comments import CommentsPart
from docx.parts.document import DocumentPart
from docx.parts.endnotes import EndnotesPart
from docx.parts.footnotes import FootnotesPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.image import ImagePart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.styles import StylesPart


def part_class_selector(content_type: str, reltype: str) -> Type[Part] | None:
    if reltype == RT.IMAGE:
        return ImagePart
    return None


PartFactory.part_class_selector = part_class_selector
PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart
PartFactory.part_type_for[CT.WML_COMMENTS] = CommentsPart
PartFactory.part_type_for[CT.WML_DOCUMENT_MAIN] = DocumentPart
PartFactory.part_type_for[CT.WML_ENDNOTES] = EndnotesPart
PartFactory.part_type_for[CT.WML_FOOTER] = FooterPart
PartFactory.part_type_for[CT.WML_FOOTNOTES] = FootnotesPart
PartFactory.part_type_for[CT.WML_HEADER] = HeaderPart
PartFactory.part_type_for[CT.WML_NUMBERING] = NumberingPart
PartFactory.part_type_for[CT.WML_SETTINGS] = SettingsPart
PartFactory.part_type_for[CT.WML_STYLES] = StylesPart
# -- paper-docx additions: comments-extended thread/resolution part --
from docx.opc.part import XmlPart as _PaperXmlPart

PartFactory.part_type_for[
    "application/vnd.openxmlformats-officedocument.wordprocessingml"
    ".commentsExtended+xml"
] = _PaperXmlPart

del (
    CT,
    CorePropertiesPart,
    CommentsPart,
    DocumentPart,
    EndnotesPart,
    FooterPart,
    FootnotesPart,
    HeaderPart,
    NumberingPart,
    PartFactory,
    SettingsPart,
    StylesPart,
    _PaperXmlPart,
    part_class_selector,
)
