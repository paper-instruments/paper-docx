"""Audit regressions for scrub preflight atomicity."""

from __future__ import annotations

import io

import pytest

import docx
from docx.oxml.ns import qn
from docx.scrubbing import scrub
from docx.search import find_one


def _package_bytes(document) -> bytes:
    stream = io.BytesIO()
    document.save(stream)
    return stream.getvalue()


def it_resolves_hidden_text_before_removing_comments():
    document = docx.Document()
    paragraph = document.add_paragraph("comment target")
    find_one(document, "comment target").comment("Review note", author="Reviewer")
    size = paragraph.runs[0]._r.get_or_add_rPr().get_or_add_sz()
    size.set(qn("w:val"), "not-a-size")
    before = _package_bytes(document)

    with pytest.raises(ValueError, match="invalid literal"):
        scrub(document, comments=True, metadata=False, hidden_text=True)

    assert _package_bytes(document) == before
