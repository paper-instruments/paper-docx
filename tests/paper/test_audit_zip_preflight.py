"""Regressions for ZIP preflight and OPC-aware expansion limits."""

from __future__ import annotations

import struct
import zipfile
from pathlib import Path

import pytest

import docx
from docx import _paperpkg, _zipguard
from docx.errors import PackageLimitError
from docx.opc import phys_pkg

_END_RECORD = struct.Struct("<4s4H2LH")
_ZIP64_END_RECORD = struct.Struct("<4sQ2H2L4Q")
_ZIP64_LOCATOR = struct.Struct("<4sLQL")


def _write_archive(path: Path, members: tuple[tuple[str, bytes], ...]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
        for name, data in members:
            archive.writestr(name, data)


def _zip64_count_only_archive(member_count: int) -> bytes:
    zip64_end = _ZIP64_END_RECORD.pack(
        b"PK\x06\x06",
        44,
        45,
        45,
        0,
        0,
        member_count,
        member_count,
        0,
        0,
    )
    locator = _ZIP64_LOCATOR.pack(b"PK\x06\x07", 0, 0, 1)
    legacy_end = _END_RECORD.pack(
        b"PK\x05\x06",
        0,
        0,
        0xFFFF,
        0xFFFF,
        0xFFFFFFFF,
        0xFFFFFFFF,
        0,
    )
    return zip64_end + locator + legacy_end


class DescribeCentralDirectoryPreflight:
    def it_opens_package_kernel_paths_once_before_preflight(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        path = tmp_path / "selected.zip"
        replacement = tmp_path / "replacement.zip"
        _write_archive(path, (("selected.bin", b"selected"),))
        _write_archive(replacement, (("replacement.bin", b"replacement"),))
        original_preflight = _paperpkg.preflight_zip

        def preflight_then_swap(source: object) -> None:
            original_preflight(source)
            replacement.replace(path)

        monkeypatch.setattr(_paperpkg, "preflight_zip", preflight_then_swap)

        parts, _ = _paperpkg._read_zip(path)

        assert parts == {"selected.bin": b"selected"}

    def it_opens_document_paths_once_before_preflight(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        path = tmp_path / "selected.docx"
        replacement = tmp_path / "replacement.docx"
        selected = docx.Document()
        selected.add_paragraph("selected")
        selected.save(path)
        other = docx.Document()
        other.add_paragraph("replacement")
        other.save(replacement)
        original_preflight = phys_pkg.preflight_zip

        def preflight_then_swap(source: object) -> None:
            original_preflight(source)
            replacement.replace(path)

        monkeypatch.setattr(phys_pkg, "preflight_zip", preflight_then_swap)

        reopened = docx.Document(path)

        assert [paragraph.text for paragraph in reopened.paragraphs] == ["selected"]

    def it_refuses_zip64_member_count_before_constructing_zipfile(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        path = tmp_path / "too-many.docx"
        path.write_bytes(_zip64_count_only_archive(_zipguard.MAX_MEMBER_COUNT + 1))
        constructions = 0

        class UnexpectedZipFile:
            def __init__(self, *_args, **_kwargs):
                nonlocal constructions
                constructions += 1
                raise AssertionError("ZipFile must not be constructed before count refusal")

        monkeypatch.setattr(_paperpkg.zipfile, "ZipFile", UnexpectedZipFile)

        with pytest.raises(PackageLimitError, match="member count"):
            _paperpkg._read_zip(path)

        assert constructions == 0

    def it_preflights_an_ordinary_document_open_before_constructing_zipfile(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        path = tmp_path / "too-many-on-open.docx"
        path.write_bytes(_zip64_count_only_archive(_zipguard.MAX_MEMBER_COUNT + 1))
        constructions = 0

        class UnexpectedZipFile:
            def __init__(self, *_args, **_kwargs):
                nonlocal constructions
                constructions += 1
                raise AssertionError("ZipFile must not be constructed before count refusal")

        monkeypatch.setattr(phys_pkg, "ZipFile", UnexpectedZipFile)

        with pytest.raises(PackageLimitError, match="member count"):
            docx.Document(path)

        assert constructions == 0

    def it_scans_records_instead_of_trusting_a_forged_low_count(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        path = tmp_path / "forged-count.docx"
        _write_archive(path, (("one.bin", b"1"), ("two.bin", b"2")))
        data = bytearray(path.read_bytes())
        end_offset = data.rfind(b"PK\x05\x06")
        struct.pack_into("<HH", data, end_offset + 8, 1, 1)
        path.write_bytes(data)
        constructions = 0

        class UnexpectedZipFile:
            def __init__(self, *_args, **_kwargs):
                nonlocal constructions
                constructions += 1
                raise AssertionError("ZipFile must not be constructed before count refusal")

        monkeypatch.setattr(_paperpkg.zipfile, "ZipFile", UnexpectedZipFile)

        with pytest.raises(PackageLimitError, match="member count"):
            _paperpkg._read_zip(path)

        assert constructions == 0

    def it_bounds_central_directory_bytes_before_constructing_zipfile(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        path = tmp_path / "large-directory.docx"
        _write_archive(path, (("one.bin", b"1"),))
        monkeypatch.setattr(_zipguard, "MAX_CENTRAL_DIRECTORY_BYTES", 1)
        constructions = 0

        class UnexpectedZipFile:
            def __init__(self, *_args, **_kwargs):
                nonlocal constructions
                constructions += 1
                raise AssertionError("ZipFile must not be constructed before size refusal")

        monkeypatch.setattr(_paperpkg.zipfile, "ZipFile", UnexpectedZipFile)

        with pytest.raises(PackageLimitError, match="central directory size"):
            _paperpkg._read_zip(path)

        assert constructions == 0


class DescribeContentTypeAwareLimits:
    @pytest.mark.parametrize("declaration", ["override", "override-case", "default"])
    def it_treats_non_xml_suffixes_as_xml_when_content_types_says_so(
        self,
        declaration: str,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        content_type = "application/vnd.example.document+xml"
        if declaration.startswith("override"):
            part_name = (
                "/WORD/DOCUMENT.BIN" if declaration == "override-case" else "/word/document.bin"
            )
            entry = f'<Override PartName="{part_name}" ContentType="{content_type}"/>'
        else:
            entry = f'<Default Extension="bin" ContentType="{content_type}"/>'
        content_types = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f"{entry}</Types>"
        ).encode()
        path = tmp_path / f"xml-{declaration}.docx"
        _write_archive(
            path,
            (
                ("[Content_Types].xml", content_types),
                ("word/document.bin", bytes(range(256)) * 4),
            ),
        )
        monkeypatch.setattr(_zipguard, "MAX_XML_MEMBER_BYTES", 512)
        monkeypatch.setattr(_zipguard, "MAX_BINARY_MEMBER_BYTES", 2_048)

        with pytest.raises(PackageLimitError, match=r"word/document\.bin.*512-byte"):
            _paperpkg._read_zip(path)

    def it_applies_content_type_limits_on_an_ordinary_document_open(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ):
        content_types = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
            b'relationships+xml"/>'
            b'<Override PartName="/word/document.bin" ContentType="application/vnd.'
            b'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            b"</Types>"
        )
        relationships = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            b'relationships">'
            b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            b'officeDocument/2006/relationships/officeDocument" '
            b'Target="word/document.bin"/>'
            b"</Relationships>"
        )
        body = "".join(f"<w:p><w:r><w:t>{index:04d}</w:t></w:r></w:p>" for index in range(128))
        document = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f"<w:body>{body}</w:body></w:document>"
        ).encode()
        path = tmp_path / "renamed-main-part.docx"
        _write_archive(
            path,
            (
                ("[Content_Types].xml", content_types),
                ("_rels/.rels", relationships),
                ("word/document.bin", document),
            ),
        )
        monkeypatch.setattr(_zipguard, "MAX_XML_MEMBER_BYTES", 1_024)
        monkeypatch.setattr(_zipguard, "MAX_BINARY_MEMBER_BYTES", len(document) + 1)

        with pytest.raises(PackageLimitError, match=r"word/document\.bin.*1024-byte"):
            docx.Document(path)
