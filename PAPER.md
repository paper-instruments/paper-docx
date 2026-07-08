# paper-docx Fork Ledger

Based on upstream tag `v1.2.0`, forked 2026-07-07, marker tag `paper-base`.

## v0 Additions (2026-07-07)

Purely additive per CONVENTIONS §1.1; exact signatures and refusal conditions
in `API-PROPOSAL.md` (enforced mechanically by
`tests/paper/test_api_surface.py`). New public surface:

- `docx.errors` — `PaperRefusal` hierarchy (safe refusals, distinct from bugs).
- `docx.package` — kernel re-exports (§7): `xml_equivalent`, `diff_package`,
  `patch_save` (implementation in `docx/_paperpkg.py`).
- `docx.story` — visibility-complete traversal: `story_parts`, `iter_blocks`,
  `outline`; new `FootnotesPart`/`EndnotesPart` registered.
- `docx.search` — `normalize_text`, `find_text`, `find_one`, `Span.replace`
  (surgical and tracked; revision vocabulary in `docx/oxml/revision.py`).
- `docx.blocks` — `insert_section_after`, `tracked_delete_paragraphs`,
  `tracked_replace_paragraphs` (same-parent rule; paragraph-mark stamping).
- `Document.revisions` (`docx.revision`) — enumeration + accept/reject, all
  or by author; tracked-edit algebra invariants tested.
- `docx.tableops` / `docx.numbering` — guarded table ops and
  apply-existing-numbering.

Upstream files touched (additive only): `docx/package.py` (re-export block),
`docx/__init__.py` (2 part registrations), `docx/document.py` (`revisions`
property), `docx/oxml/__init__.py` (3 tag registrations).

Test infrastructure: `tests/paper/` (contract harness, frozen 23-fixture
corpus with `MANIFEST.sha256`, sidecars, LO smoke). Human follow-ups tracked
in `FIXTURE-REQUESTS.md`.

## Baseline Test Runs

Environment: CPython 3.13.5 via `uv`; test dependencies installed from upstream
`requirements-test.txt`.

- `pytest -q`: failed during collection with 35 errors, all caused by
  `pyparsing.warnings.PyparsingDeprecationWarning: 'delimitedList' deprecated - use 'DelimitedList'`.
  This is pre-existing environment drift from current `pyparsing` plus upstream's
  warnings-as-errors pytest configuration.
- `pytest -q -W ignore::pyparsing.warnings.PyparsingDeprecationWarning`: ran the
  suite and reported `1600 passed, 9 errors`; the remaining errors are current
  pytest fixture deprecation/error behavior around class-scoped fixtures defined
  as instance methods in `tests/opc/test_phys_pkg.py`.
- `pytest -q -W ignore::pyparsing.warnings.PyparsingDeprecationWarning -W ignore::pytest.PytestRemovedIn10Warning`:
  `1609 passed`.
- `behave -q`: `67 features passed, 0 failed, 0 skipped`; `650 scenarios passed,
  0 failed, 0 skipped`; `1856 steps passed, 0 failed, 0 skipped`.
- `uv build`: built `dist/paper_docx-0.1.0.tar.gz` and
  `dist/paper_docx-0.1.0-py3-none-any.whl`. Setuptools emitted pre-existing
  license metadata deprecation warnings from upstream pyproject structure.
- Wheel smoke test:
  `uv run --isolated --no-project --with dist/*.whl python -c "import docx; print(docx.__paper_version__)"`
  printed `0.1.0`.
- Source distribution smoke test:
  `uv run --isolated --no-project --with dist/*.tar.gz python -c "import docx; print(docx.__paper_version__)"`
  printed `0.1.0`.

The CI workflow applies the two narrow pytest warning filters above when those
warning classes exist in the matrix environment, so current tooling can run the
upstream suite without modifying upstream source code.

## Publishing Safety

Publishing is intentionally disabled by default while this repository is
private. The release workflow targets the `pypi` environment and the publish
step is additionally guarded by `vars.PUBLISH_ENABLED == 'true'`. Configure
required reviewers on the `pypi` environment in GitHub before any release.

Do not push upstream `v*` tags to origin. Only the `paper-base` marker tag is
pushed during bootstrap.

## Sanctioned Deviations From Upstream Behavior

None.

## Upstream Merge Policy

Quarterly, run `git fetch upstream --tags`, identify whether a newer upstream
release tag exists, merge that release tag into `main`, and run both the pytest
and behave suites plus package smoke tests. Resolve conflicts using this file as
the map of intentional fork identity changes. Merge upstream releases; never
rebase `main`.
