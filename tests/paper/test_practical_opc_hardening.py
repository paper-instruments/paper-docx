"""Focused regressions for practical OPC read/write hardening."""

from __future__ import annotations

import copy
import io
import stat
import zipfile
from pathlib import Path

import pytest

import docx
import docx.opc.package as package_module
from docx._transaction import rollback_on_error
from docx.errors import PackageLimitError
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.opc.pkgreader import _SerializedRelationships
from docx.search import find_one
from docx.story import story_parts


def _document_bytes() -> bytes:
    stream = io.BytesIO()
    docx.Document().save(stream)
    return stream.getvalue()


def _rewrite_package(data: bytes, transform) -> bytes:
    destination = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(
        destination, "w", zipfile.ZIP_DEFLATED
    ) as output:
        for name in source.namelist():
            replacement = transform(name, source.read(name))
            if replacement is not None:
                output.writestr(*replacement)
    return destination.getvalue()


def it_saves_paths_atomically_preserving_mode_and_following_destination_symlink(
    tmp_path: Path,
):
    target = tmp_path / "target.docx"
    target.write_bytes(b"old")
    target.chmod(0o640)
    link = tmp_path / "link.docx"
    link.symlink_to(target.name)

    document = docx.Document()
    document.add_paragraph("saved through link")
    document.save(link)

    assert link.is_symlink()
    assert stat.S_IMODE(target.stat().st_mode) == 0o640
    assert docx.Document(target).paragraphs[-1].text == "saved through link"
    assert not list(tmp_path.glob("*.partial"))


def it_refuses_if_the_destination_symlink_changes_during_save(
    tmp_path: Path, monkeypatch
):
    old_target = tmp_path / "old.docx"
    new_target = tmp_path / "new.docx"
    old_target.write_bytes(b"old target")
    new_target.write_bytes(b"new target")
    link = tmp_path / "link.docx"
    link.symlink_to(old_target.name)
    validate = package_module._validate_serialized_output

    def validate_and_repoint(source):
        validate(source)
        link.unlink()
        link.symlink_to(new_target.name)

    monkeypatch.setattr(
        package_module, "_validate_serialized_output", validate_and_repoint
    )

    with pytest.raises(OSError, match="symlink changed"):
        docx.Document().save(link)

    assert old_target.read_bytes() == b"old target"
    assert new_target.read_bytes() == b"new target"


def it_restores_a_seekable_stream_after_commit_error():
    class FailOnce(io.BytesIO):
        failed = False

        def write(self, data):
            if not self.failed:
                self.failed = True
                super().write(data[:7])
                raise OSError("commit failed")
            return super().write(data)

    original = b"prefix and original suffix"
    stream = FailOnce(original)
    stream.seek(7)

    with pytest.raises(OSError, match="commit failed"):
        docx.Document().save(stream)

    assert stream.getvalue() == original
    assert stream.tell() == 7


def it_validates_the_real_prefix_of_a_nonzero_position_stream(monkeypatch):
    prefix = b"real destination prefix"
    original = prefix + b"existing suffix"
    stream = io.BytesIO(original)
    stream.seek(len(prefix))
    validate = package_module._validate_serialized_output
    validated_prefixes = []

    def validate_and_record(source):
        source.seek(0)
        validated_prefixes.append(source.read(len(prefix)))
        source.seek(0)
        validate(source)

    monkeypatch.setattr(package_module, "_validate_serialized_output", validate_and_record)

    docx.Document().save(stream)

    assert validated_prefixes == [prefix]
    assert stream.getvalue().startswith(prefix)
    stream.seek(0)
    assert isinstance(docx.Document(stream), docx.document.Document)


def it_refuses_before_overwriting_an_unreadable_seekable_stream():
    class Unreadable(io.BytesIO):
        def read(self, size=-1):
            raise OSError("existing bytes are unavailable")

    original = b"existing destination"
    stream = Unreadable(original)

    with pytest.raises(OSError, match="could not be snapshotted"):
        docx.Document().save(stream)

    assert stream.getvalue() == original
    assert stream.tell() == 0


def it_restores_mutables_base_uri_and_detaches_corrupt_graph_residuals():
    class Participant:
        def __init__(self):
            self.values = ["original"]

    document = docx.Document()
    participant = Participant()
    values = participant.values
    relationships = document.part.rels
    base_uri = relationships._baseURI
    rogue = Part(
        PackURI("/word/rogue.bin"), "application/octet-stream", b"rogue", document.part.package
    )

    def fail_late():
        with rollback_on_error(document, participant):
            values.append("changed")
            relationships._baseURI = "/corrupt"
            relationships._target_parts_by_rId["rId-corrupt"] = rogue
            raise ValueError("late failure")

    with pytest.raises(ValueError, match="late failure"):
        fail_late()

    assert participant.values is values
    assert values == ["original"]
    assert relationships._baseURI == base_uri
    assert "rId-corrupt" not in relationships._target_parts_by_rId
    assert rogue.package is None


def it_resolves_unambiguous_physical_members_without_case_sensitivity():
    source = docx.Document()
    source.add_paragraph("case-insensitive story")
    stream = io.BytesIO()
    source.save(stream)
    data = _rewrite_package(
        stream.getvalue(),
        lambda name, blob: (
            ("WORD/DOCUMENT.XML", blob) if name == "word/document.xml" else (name, blob)
        ),
    )

    document = docx.Document(io.BytesIO(data))

    assert story_parts(document)[0] == "word/document.xml"
    assert find_one(document, "case-insensitive story").story == "word/document.xml"


def it_refuses_a_symlinked_member_in_an_expanded_package(tmp_path: Path):
    expanded = tmp_path / "expanded"
    with zipfile.ZipFile(io.BytesIO(_document_bytes())) as package:
        package.extractall(expanded)
    member = expanded / "word" / "document.xml"
    outside = tmp_path / "outside-document.xml"
    member.replace(outside)
    member.symlink_to(outside)

    with pytest.raises(PackageLimitError, match="symbolic link"):
        docx.Document(expanded)


def it_refuses_duplicate_relationship_ids():
    relationships = b"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="urn:one" Target="one.xml"/>
      <Relationship Id="rId1" Type="urn:two" Target="two.xml"/>
    </Relationships>"""

    with pytest.raises(PackageLimitError, match="duplicate Id"):
        _SerializedRelationships.load_from_xml("/word", relationships)


def it_refuses_multiple_main_document_relationships_before_replacing_a_path(
    tmp_path: Path,
):
    destination = tmp_path / "destination.docx"
    destination.write_bytes(b"original")
    document = docx.Document()
    document.part.package.load_rel(
        RT.OFFICE_DOCUMENT, document.part, "rIdDuplicateMain"
    )

    with pytest.raises(PackageLimitError, match="multiple officeDocument"):
        document.save(destination)

    assert destination.read_bytes() == b"original"


def it_accepts_xml_comments_in_relationship_parts():
    def transform(name: str, blob: bytes):
        if name == "_rels/.rels":
            blob = blob.replace(b"<Relationship ", b"<!-- retained --><Relationship ", 1)
        return name, blob

    document = docx.Document(
        io.BytesIO(_rewrite_package(_document_bytes(), transform))
    )

    assert document.paragraphs is not None


def it_refuses_unexpected_relationship_elements_instead_of_dropping_them():
    def transform(name: str, blob: bytes):
        if name == "word/_rels/document.xml.rels":
            from lxml import etree

            root = etree.fromstring(blob)
            relationship = next(
                child for child in root if (child.get("Type") or "").endswith("/styles")
            )
            relationship.tag = relationship.tag.replace("Relationship", "Unexpected")
            blob = etree.tostring(root, encoding="UTF-8", standalone=True)
        return name, blob

    with pytest.raises(PackageLimitError, match="unexpected element"):
        docx.Document(io.BytesIO(_rewrite_package(_document_bytes(), transform)))


def it_validates_case_insensitive_content_types_without_losing_duplicates():
    def transform(name: str, blob: bytes):
        if name != "[Content_Types].xml":
            return name, blob
        from lxml import etree

        root = etree.fromstring(blob)
        override = next(child for child in root if child.get("PartName"))
        duplicate = copy.deepcopy(override)
        duplicate.set("PartName", override.get("PartName").upper())
        root.append(duplicate)
        return (
            "[CONTENT_TYPES].XML",
            etree.tostring(root, encoding="UTF-8", standalone=True),
        )

    with pytest.raises(PackageLimitError, match="ambiguous Override"):
        docx.Document(io.BytesIO(_rewrite_package(_document_bytes(), transform)))


def it_refuses_a_known_relationship_role_with_the_wrong_content_type():
    def transform(name: str, blob: bytes):
        if name == "[Content_Types].xml":
            blob = blob.replace(
                b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                b"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            )
        return name, blob

    with pytest.raises(PackageLimitError, match="invalid content type"):
        docx.Document(io.BytesIO(_rewrite_package(_document_bytes(), transform)))


def it_refuses_a_relationship_to_a_missing_package_part():
    data = _rewrite_package(
        _document_bytes(),
        lambda name, blob: None if name == "word/styles.xml" else (name, blob),
    )

    with pytest.raises(PackageLimitError, match="targets missing package part"):
        docx.Document(io.BytesIO(data))


def it_refuses_a_relationship_target_without_a_declared_content_type():
    def transform(name: str, blob: bytes):
        if name == "[Content_Types].xml":
            from lxml import etree

            root = etree.fromstring(blob)
            for declaration in tuple(root):
                if declaration.get("PartName") == "/word/styles.xml" or declaration.get(
                    "Extension"
                ) == "xml":
                    root.remove(declaration)
            blob = etree.tostring(root, encoding="UTF-8", standalone=True)
        return name, blob

    with pytest.raises(PackageLimitError, match="no declared content type"):
        docx.Document(io.BytesIO(_rewrite_package(_document_bytes(), transform)))


def it_retains_default_creation_for_missing_optional_core_properties():
    def transform(name: str, blob: bytes):
        if name == "docProps/core.xml":
            return None
        if name == "_rels/.rels":
            from lxml import etree

            root = etree.fromstring(blob)
            for relationship in tuple(root):
                if relationship.get("Target") == "docProps/core.xml":
                    root.remove(relationship)
            blob = etree.tostring(root, encoding="UTF-8", standalone=True)
        return name, blob

    document = docx.Document(io.BytesIO(_rewrite_package(_document_bytes(), transform)))

    assert document.core_properties.title == "Word Document"
