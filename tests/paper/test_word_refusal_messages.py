"""Refusal messages that name the defect Word objected to, and the fix.

Word refuses every package exercised here; only the wording is under test. Each case
asserts the distinguishing phrase rather than the whole string, so copy-editing a message
does not break the suite.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

import docx
import docx.opc.pkgreader as pkgreader_module
from docx.errors import MalformedPackageError
from docx.opc.pkgreader import _SerializedRelationships

_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _document_bytes() -> bytes:
    stream = io.BytesIO()
    docx.Document().save(stream)
    return stream.getvalue()


def _relationships_xml(second_target: str) -> bytes:
    return (
        f'<Relationships xmlns="{_RELATIONSHIPS_NS}">'
        '<Relationship Id="rId1" Type="urn:one" Target="one.xml"/>'
        f'<Relationship Id="rId1" Type="urn:one" Target="{second_target}"/>'
        "</Relationships>"
    ).encode("utf-8")


def _refusal(callable_, *args) -> str:
    with pytest.raises(MalformedPackageError) as exc_info:
        callable_(*args)
    return str(exc_info.value)


def _duplicate_first_default(data: bytes) -> tuple[bytes, str]:
    """Repeat the first `Default` element of `[Content_Types].xml` verbatim."""
    extension = ""
    destination = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as package, zipfile.ZipFile(
        destination, "w", zipfile.ZIP_DEFLATED
    ) as output:
        for name in package.namelist():
            blob = package.read(name)
            if name == "[Content_Types].xml":
                text = blob.decode("utf-8")
                element = re.search(r'<Default Extension="([^"]+)"[^>]*/>', text)
                assert element is not None
                extension = element.group(1)
                blob = text.replace(
                    element.group(0), element.group(0) * 2, 1
                ).encode("utf-8")
            output.writestr(name, blob)
    return destination.getvalue(), extension


def it_names_the_duplicated_extension_when_a_default_is_declared_twice():
    """Two entries declaring the same content type are duplicated, not ambiguous."""
    data, extension = _duplicate_first_default(_document_bytes())

    message = _refusal(docx.Document, io.BytesIO(data))

    assert "more than one Default" in message
    assert repr(extension) in message
    assert "Delete the duplicate Default" in message
    assert "ambiguous" not in message


def it_reports_the_missing_footer_rather_than_the_record_that_holds_it():
    """The actionable fact is where the archive stops, not the record it stops with."""
    message = _refusal(docx.Document, io.BytesIO(_document_bytes() + b"\x00" * 64))

    assert "does not end with an archive footer" in message
    assert "Re-saving the document" in message
    assert "end-of-central-directory" not in message


def it_does_not_blame_appended_bytes_for_a_file_that_is_merely_short():
    """The same refusal covers truncation and non-packages, so it may not assert a tail."""
    message = _refusal(docx.Document, io.BytesIO(_document_bytes()[:4096]))

    assert "truncated in transfer" in message
    assert "needs a fresh copy" in message


def it_calls_a_repeated_id_for_one_target_a_redundant_declaration():
    message = _refusal(
        _SerializedRelationships.load_from_xml, "/word", _relationships_xml("one.xml")
    )

    assert "twice for the same target" in message
    assert "Delete the repeated Relationship element" in message


def it_calls_a_repeated_id_for_two_targets_a_contradiction():
    message = _refusal(
        _SerializedRelationships.load_from_xml, "/word", _relationships_xml("two.xml")
    )

    assert "two different targets" in message
    assert "an Id of its own" in message


def it_distinguishes_the_redundant_and_contradictory_relationship_refusals():
    """Both stay refused, but the caller's next step differs."""
    redundant = _refusal(
        _SerializedRelationships.load_from_xml, "/word", _relationships_xml("one.xml")
    )
    contradictory = _refusal(
        _SerializedRelationships.load_from_xml, "/word", _relationships_xml("two.xml")
    )

    assert redundant != contradictory


def it_distinguishes_the_two_cases_through_the_object_model_path_too(monkeypatch):
    """The second call site validates the built objects; it must draw the same line.

    `_validate_relationship_records` reaches the raw XML first, so it is stubbed out here
    to let the object-model check run. Both sites stay, and both name their case.
    """
    monkeypatch.setattr(
        pkgreader_module, "_validate_relationship_records", lambda blob, base_uri: None
    )

    redundant = _refusal(
        _SerializedRelationships.load_from_xml, "/word", _relationships_xml("one.xml")
    )
    contradictory = _refusal(
        _SerializedRelationships.load_from_xml, "/word", _relationships_xml("two.xml")
    )

    assert "twice for the same target" in redundant
    assert "two different targets" in contradictory
    assert redundant != contradictory
