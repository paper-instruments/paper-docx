"""Tests for the v0.1 Phase 0c write guards (H5-H8) and triage (H10)."""

from __future__ import annotations

import datetime as dt
import shutil
import zipfile
from pathlib import Path

import pytest

import docx
from docx.errors import TargetNotFoundError, UnsupportedStructureError
from docx.package import diagnose
from docx.search import find_one

from .harness.contract import assert_refusal_atomic
from .harness.paths import fixture_path

FROZEN = dt.datetime(2026, 7, 8, 9, 0, 0, tzinfo=dt.timezone.utc)
MINIMAL = "generated/minimal-clean/minimal.docx"
BOOKMARKS = "generated/feature-isolated/bookmarks.docx"
NOISY = "generated/feature-isolated/noisy-markup.docx"


def _doc(relpath: str = MINIMAL):
    return docx.Document(str(fixture_path(relpath)))


class DescribeControlCharacterRefusal:
    """H6: raw \\n/\\t in written text would be verified-but-false structure."""

    @pytest.mark.parametrize("bad", ["Line one\nLine two", "col\tcol", "a\rb"])
    def it_refuses_control_characters_in_replacement_text(self, bad: str):
        span = find_one(_doc(), "perfectly ordinary")
        with pytest.raises(ValueError, match="control character"):
            span.replace(bad)

    def it_refuses_control_characters_in_inserted_paragraphs(self):
        from docx.blocks import insert_section_after

        with pytest.raises(ValueError, match="control character"):
            insert_section_after(
                _doc(), "First body paragraph",
                heading="OK", paragraphs=["fine", "not\nfine"],
            )

    def it_refuses_control_characters_in_tracked_replacements(self):
        from docx.blocks import tracked_replace_paragraphs

        with pytest.raises(ValueError, match="control character"):
            tracked_replace_paragraphs(
                _doc(), "Second body paragraph", ["multi\nline"],
                author="Carol QA", date=FROZEN,
            )


class DescribeFakeBulletRefusal:
    """H7: a list style whose numbering does not resolve must refuse."""

    def it_applies_a_list_style_whose_numbering_resolves(self):
        from docx.numbering import apply_list_style

        document = _doc()
        paragraph = document.add_paragraph("bulleted")
        apply_list_style(paragraph, "List Bullet")
        assert paragraph.style.name == "List Bullet"

    def it_refuses_a_list_style_with_dangling_numbering(self, tmp_path: Path):
        from docx.numbering import apply_list_style

        import re

        source = fixture_path(MINIMAL)
        stripped = tmp_path / "no-numbering.docx"
        with zipfile.ZipFile(source) as zin:
            with zipfile.ZipFile(stripped, "w") as zout:
                for name in zin.namelist():
                    if name == "word/numbering.xml":
                        continue
                    blob = zin.read(name)
                    if name == "word/_rels/document.xml.rels":
                        # drop the numbering relationship entirely (a dangling
                        # rel would break upstream's loader before our check)
                        blob = re.sub(rb'<Relationship [^>]*numbering\.xml"/>', b"", blob)
                    zout.writestr(name, blob)
        document = docx.Document(str(stripped))
        paragraph = document.add_paragraph("wants a bullet")
        with pytest.raises(TargetNotFoundError, match="fake"):
            apply_list_style(paragraph, "List Bullet")

    def it_still_applies_styles_without_numbering_bindings(self):
        from docx.numbering import apply_list_style

        document = _doc()
        paragraph = document.add_paragraph("just a heading")
        apply_list_style(paragraph, "Heading 3")
        assert paragraph.style.name == "Heading 3"


class DescribeUntrackedEditInsideInsertions:
    """H8a: untracked edits must not rewrite text attributed to an author."""

    def it_refuses_untracked_replace_inside_a_pending_insertion(self, tmp_path: Path):
        path = tmp_path / "doc.docx"
        shutil.copyfile(fixture_path(MINIMAL), path)
        document = docx.Document(str(path))
        find_one(document, "perfectly ordinary").replace(
            "written by Alice", tracked=True, author="Alice Editor", date=FROZEN
        )

        def untracked_rewrite(doc):
            find_one(doc, "written by Alice").replace("forged words")

        assert_refusal_atomic(
            document, untracked_rewrite, UnsupportedStructureError, on_disk=(path,)
        )


class DescribeBookmarkHollowing:
    """H8b: replaces must not silently empty cross-reference targets."""

    def it_refuses_a_replace_that_hollows_a_named_bookmark(self):
        document = _doc(BOOKMARKS)
        span = find_one(document, "See the Master Agreement for definitions.")
        with pytest.raises(UnsupportedStructureError, match="DefinedTerm"):
            span.replace("See the Purchase Agreement for definitions.")

    def it_allows_replacing_text_inside_the_bookmark(self):
        """Markers outside the span: the new text stays bookmarked."""
        document = _doc(BOOKMARKS)
        find_one(document, "the Master Agreement").replace("the Purchase Agreement")
        assert "the Purchase Agreement" in find_one(
            document, "See the Purchase Agreement for definitions."
        ).text

    def it_treats_point_bookmarks_as_transparent(self):
        """_GoBack noise must never block an edit."""
        document = _doc(NOISY)
        find_one(document, "Paragrah with a spelling issue.").replace(
            "Paragraph without a spelling issue."
        )


class DescribePackageDiagnosis:
    """H10: typed triage instead of raw KeyError/ValueError dead ends."""

    def it_diagnoses_a_healthy_docx(self):
        report = diagnose(fixture_path(MINIMAL))
        assert report.readable and report.kind == "docx" and not report.problems
        payload = report.to_dict()
        assert payload["schema"] == "paper_diagnosis" and payload["version"] == 1

    def it_diagnoses_missing_files(self, tmp_path: Path):
        report = diagnose(tmp_path / "nope.docx")
        assert not report.readable and report.kind == "missing"

    def it_diagnoses_encrypted_or_legacy_binaries(self, tmp_path: Path):
        cfb = tmp_path / "locked.docx"
        cfb.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64)
        report = diagnose(cfb)
        assert not report.readable
        assert report.kind == "encrypted-or-legacy-binary"
        assert "password" in report.problems[0]

    def it_diagnoses_non_zip_content(self, tmp_path: Path):
        text = tmp_path / "notes.docx"
        text.write_bytes(b"just some text pretending to be a docx")
        assert diagnose(text).kind == "not-a-zip"

    def it_diagnoses_macro_enabled_documents(self, tmp_path: Path):
        source = fixture_path(MINIMAL)
        docm = tmp_path / "macros.docm"
        with zipfile.ZipFile(source) as zin:
            with zipfile.ZipFile(docm, "w") as zout:
                for name in zin.namelist():
                    blob = zin.read(name)
                    if name == "[Content_Types].xml":
                        blob = blob.replace(
                            b"application/vnd.openxmlformats-officedocument"
                            b".wordprocessingml.document.main+xml",
                            b"application/vnd.ms-word.document.macroEnabled.main+xml",
                        )
                    zout.writestr(name, blob)
        report = diagnose(docm)
        assert report.kind == "docm" and not report.readable
        assert "macro" in report.problems[0]

    def it_diagnoses_the_corrupt_fixtures(self):
        report = diagnose(fixture_path("generated/corrupt/malformed-xml.docx"))
        # zip is fine; the damage is inside a part — readable at OPC level
        assert report.kind == "docx"

    def it_diagnoses_other_opc_packages(self, tmp_path: Path):
        fake_xlsx = tmp_path / "book.xlsx"
        with zipfile.ZipFile(fake_xlsx, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
            zf.writestr("_rels/.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>")
            zf.writestr("xl/workbook.xml", "<workbook/>")
        report = diagnose(fake_xlsx)
        assert report.kind == "xlsx" and not report.readable
