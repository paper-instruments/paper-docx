"""Independent-loader smoke: LibreOffice must open every clean fixture.

Marked `lo_smoke`; skipped automatically when no soffice binary is installed
(CONVENTIONS §4, assertion 4).
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
