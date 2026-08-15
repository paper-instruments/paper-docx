"""Independent-loader smoke: LibreOffice must open every clean fixture.

Marked `lo_smoke`; skipped automatically when no soffice binary is installed.
"""

from __future__ import annotations

import pytest

from .harness.lo import assert_libreoffice_opens
from .harness.paths import FIXTURES_DIR, fixture_path, iter_fixture_docx_paths

#: corrupt fixtures are negative-only; `large` is excluded solely to keep the
#: smoke suite fast (it IS reopened via docx.Document() at authoring time,
#: but is never LibreOffice-converted anywhere).
EXCLUDED_TAXONOMIES = {"corrupt", "large"}


def _smoke_relpaths() -> list:
    return [
        rel
        for rel in (
            p.relative_to(FIXTURES_DIR).as_posix() for p in iter_fixture_docx_paths()
        )
        if rel.split("/")[1] not in EXCLUDED_TAXONOMIES
    ]


@pytest.mark.lo_smoke
class DescribeLibreOfficeSmoke:
    @pytest.mark.parametrize("relpath", _smoke_relpaths())
    def it_opens_every_clean_fixture(self, relpath: str):
        assert_libreoffice_opens(fixture_path(relpath))

    def it_opens_a_compare_engine_redline(self, tmp_path):
        """Emitted tracked-changes markup must be valid to an independent
        loader."""
        import datetime as dt

        from docx.package import compare

        result = compare(
            str(fixture_path("generated/redline/compare-original.docx")),
            str(fixture_path("generated/redline/compare-revised.docx")),
            author="Compare Engine",
            date=dt.datetime(2026, 7, 8, 12, 0, 0, tzinfo=dt.timezone.utc),
        )
        out = tmp_path / "redline.docx"
        result.document.save(str(out))
        assert_libreoffice_opens(out)
