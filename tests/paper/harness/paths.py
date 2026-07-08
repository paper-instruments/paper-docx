"""Filesystem layout of the frozen fixture corpus.

Fixtures live under ``tests/paper/fixtures/<provenance>/<taxonomy>/``.

Provenance buckets (CONVENTIONS §4): ``word`` (authored in desktop Microsoft
Word), ``google`` (exported from Google Docs), ``libreoffice`` (exported from
LibreOffice), ``other`` (other real-world producers), ``generated`` (produced
by this repo's own authoring code — never by code under test).

Taxonomy buckets: ``minimal-clean``, ``feature-isolated``, ``gauntlet``,
``corrupt``, ``large``.

Every ``.docx`` fixture has a same-stem ``.json`` sidecar next to it, and both
are hash-frozen in ``MANIFEST.sha256`` (see ``manifest.py``).
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

FIXTURES_DIR = Path(__file__).resolve().parent.parent / "fixtures"

PROVENANCE_BUCKETS = ("word", "google", "libreoffice", "other", "generated")
TAXONOMY_BUCKETS = ("minimal-clean", "feature-isolated", "gauntlet", "corrupt", "large")


def fixture_path(relpath: str) -> Path:
    """Absolute path of the fixture at `relpath` (relative to the fixtures root)."""
    path = FIXTURES_DIR / relpath
    if not path.is_file():
        raise FileNotFoundError(f"no such fixture: {relpath!r} (looked at {path})")
    return path


def sidecar_path(docx_relpath_or_path: str | Path) -> Path:
    """Path of the JSON sidecar belonging to the given `.docx` fixture."""
    path = Path(docx_relpath_or_path)
    if not path.is_absolute():
        path = FIXTURES_DIR / path
    return path.with_suffix(".json")


def iter_fixture_docx_paths() -> Iterator[Path]:
    """All frozen `.docx` fixture files, in deterministic (sorted) order."""
    yield from sorted(FIXTURES_DIR.rglob("*.docx"))


def iter_manifest_covered_paths() -> Iterator[Path]:
    """All files whose hashes are frozen in MANIFEST.sha256, sorted.

    That is every file under the fixtures root except the manifest itself and
    OS junk (dotfiles like Finder's `.DS_Store`), which must neither be frozen
    nor fail the coverage check.
    """
    for path in sorted(FIXTURES_DIR.rglob("*")):
        if not path.is_file() or path.name == "MANIFEST.sha256":
            continue
        rel = path.relative_to(FIXTURES_DIR)
        if any(part.startswith(".") for part in rel.parts):
            continue
        yield path


def provenance_bucket_of(path: Path) -> str:
    """The provenance bucket a fixture path belongs to (its first path segment)."""
    rel = path.resolve().relative_to(FIXTURES_DIR)
    return rel.parts[0]
