# Contributing to paper-docx

Thanks for your interest in improving paper-docx. This document covers local
setup, how to run the checks CI runs, and the conventions that keep the fork
safe to change.

## Development setup

The project is managed with [uv](https://docs.astral.sh/uv/). Clone and sync:

```bash
git clone https://github.com/paper-instruments/paper-docx
cd paper-docx
uv sync
```

`uv sync` installs the package in editable mode with the `dev` dependency
group (pytest, behave, ruff, pyright, Sphinx).

## Running the checks

CI (`.github/workflows/test.yml`) runs the pytest and behave suites on Python
3.9–3.13, a LibreOffice contract job, a strict Sphinx docs build, and a
distribution build with an install-collision matrix. Locally:

```bash
uv run pytest                 # unit suite (also: make test)
uv run behave --stop          # acceptance suite (also: make accept)
uv run ruff check src tests   # lint
uv run pyright                # strict type checking
make docs                     # Sphinx build
```

The LibreOffice smoke tests (`pytest -m lo_smoke tests/paper/test_lo_smoke.py`)
require a local `soffice`; they are skipped automatically when it is absent.

## What a change must preserve

- **Refusal atomicity.** Every mutating Paper operation either does exactly
  what it claims or raises a typed `PaperRefusal` subclass and leaves the
  document unchanged in memory and on disk. New operations must validate
  before mutating, or wrap late phases in the `docx._transaction` helpers, and
  must ship a late-failure test.
- **The frozen import.** The importable package is `docx`, forever. The
  distribution name is `paper-docx`. Never change either.
- **Typed outcomes.** New failure modes use the existing `docx.errors`
  taxonomy; plain `ValueError`/`TypeError` remain reserved for programmer
  errors.
- **Upstream compatibility.** The inherited python-docx test suites must stay
  green. Changes to inherited behavior are exceptional, must be
  safety-motivated, and must be called out explicitly in the PR description.

## Tests and fixtures

Fork-specific tests live in `tests/paper/`. Fixtures are a frozen,
hash-pinned corpus under `tests/paper/fixtures/` with a `MANIFEST.sha256` and
per-fixture sidecars; do not modify a fixture in place. To add one, follow the
generation notes in `tests/paper/fixtures/README.md` (where present), pin its
hash in the manifest, and label its provenance (generated vs
LibreOffice-authored) honestly.

Behavioral tests should save and reopen the document rather than asserting on
in-memory state alone, and should assert a changed-part budget (see the
harness helpers in `tests/paper/harness/`).

## Submitting changes

- Open an issue first for anything that changes public API surface;
  signatures are pinned by `tests/paper/test_api_surface.py`.
- Keep PRs narrow: one behavior per PR, with the test that proves it.
- Write a clear, user-facing PR title: release notes are generated from
  merged PR titles when a release is tagged.
- Update `docs/user/paper-additions.rst` and the matching
  `docs/api/paper-*.rst` page for any public API change; the docs build runs
  with warnings as errors.

## Releases

Releases are tag-driven: pushing `v<version>` runs the full quality gate,
verifies the tag matches `docx.__paper_version__`, and publishes to PyPI via
trusted publishing. Only maintainers cut releases.
