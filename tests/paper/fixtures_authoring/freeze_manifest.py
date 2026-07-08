#!/usr/bin/env python3
"""Rewrite tests/paper/fixtures/MANIFEST.sha256 from the on-disk corpus.

    uv run --no-sync python tests/paper/fixtures_authoring/freeze_manifest.py

This is the ONE sanctioned way to update fixture hashes (CONVENTIONS §4:
golden files update only via an explicit command). The manifest test fails on
any drift, so running this is always followed by human review of the diff in
the same PR that changes the fixtures.
"""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from tests.paper.harness.manifest import (  # noqa: E402
    MANIFEST_PATH,
    compute_manifest,
    read_manifest,
    render_manifest,
)


def main() -> int:
    current = compute_manifest()
    previous = read_manifest() if MANIFEST_PATH.is_file() else {}
    MANIFEST_PATH.write_text(render_manifest(current), encoding="utf-8")

    added = sorted(set(current) - set(previous))
    removed = sorted(set(previous) - set(current))
    changed = sorted(k for k in set(current) & set(previous) if current[k] != previous[k])
    print(f"froze {len(current)} files into {MANIFEST_PATH.relative_to(REPO_ROOT)}")
    for label, names in (("added", added), ("removed", removed), ("changed", changed)):
        for name in names:
            print(f"  {label}: {name}")
    if not (added or removed or changed):
        print("  (no drift)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
