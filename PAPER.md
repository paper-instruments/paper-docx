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

## Paper Ledger (additive work)

All entries below are purely additive (§1.1); upstream suites stayed green at
every commit (pytest 1609, behave 650 scenarios). Signatures are pinned in
`API-PROPOSAL.md` and mechanically enforced by `tests/paper/test_api_surface.py`.

### v0 (2026-07-07) — the safe-editing surface

Phases 0–9 of `agent_docs/PLAN-paper-docx.md`: `tests/paper/` harness +
frozen fixture corpus; `docx.errors`; `docx.package` kernel
(`xml_equivalent`/`diff_package`/`patch_save`); `docx.story`
visibility-complete traversal; `docx.search` (`find_text`/`Span.replace`,
surgical + tracked, new `w:ins`/`w:del` oxml vocabulary); `docx.blocks`;
`Document.revisions` with the tracked-edit algebra; guarded `docx.tableops` +
`docx.numbering`. Upstream files touched (additively only): re-export block
in `docx/package.py`, part registrations in `docx/__init__.py`,
`Document.revisions` property.

### v0.11 (2026-07-08) — revision completion, scrub, compare, composition

`agent_docs/PLAN-v0.11-paper-docx.md`. Refusal→capability conversions (the
sanctioned growth path, §1.1 scope rule) and new organs; every conversion
listed here.

- **Phase 1 — format-change and structural revision resolution.**
  Conversions: `w:rPrChange`/`w:pPrChange` (run, paragraph, paragraph-mark)
  accept/reject with stored-property restore; `w:trPr` row markers
  reclassified from insertion/deletion (which resolved only the marker —
  the ghost-row false state) to `row_insertion`/`row_deletion` with whole-row
  semantics and an empty-table refusal (batch-validated upfront for
  atomicity). Reject of a deletion now also restores `w:delInstrText` →
  `w:instrText`. The exotic remainder is enumerated and refused BY NAME:
  `table_property_change` (tblPrChange/tblPrExChange/tblGridChange/
  trPrChange/tcPrChange), `cell_revision` (cellIns/cellDel/cellMerge),
  `section_property_change`, `numbering_change`, `custom_xml_revision`.
  `Revisions.to_dict()` schema v3: new type names, and format-change
  revisions now carry the text they apply to (they were unaddressable with
  `text == ""`). Zero-markup rescan oracle `docx.revision._remaining_markup`.
- **Phase 2 — move resolution.** Conversion: `w:moveFrom`/`w:moveTo` with
  their range markers (paired by `w:name`) resolve as ONE unit — accepting
  or rejecting either site resolves both, never one side alone. Accept =
  destination becomes plain, source range (content + moved paragraph marks)
  disappears; reject = the symmetric inverse. Paragraph-mark move stamps
  ride the existing mark machinery (`w:moveFrom` del-like, `w:moveTo`
  ins-like). Orphaned, duplicated, nameless, or cross-story move markup →
  typed refusal naming the defect, validated upfront for batch atomicity.
  Comment anchors inside a dropped site are cleaned up with their comments.
  The gauntlet now resolves completely in both directions (rescan == zero
  markup), and `accept_all` on the multiround redline fixture reproduces
  the hand-computed accepted ground truth story-for-story.
- **Phase 3 — scrub/finalize + protection-aware mutations.** New surface
  (API-PROPOSAL §10): `Document.finalize(revisions=...)` (total resolution
  or typed refusal; certifies zero remaining markup), `Document.scrub(...)`
  (comments incl. extended/threading/people parts + anchors, core/app/custom
  metadata + thumbnail, trackRevisions setting, optional RSIDs, optional
  hidden text) returning an itemized `ScrubReport` (report-matches-diff
  contract, tested against `diff_package`). Metadata scrub refuses while
  revisions are pending. **Sanctioned deviation of paper APIs only** (the
  v0.1 honesty-recall precedent, per PLAN-v0.11): enforced
  `w:documentProtection` now makes every paper-docx mutating API refuse
  with `DocumentProtectedError` (new pinned refusal subclass — flagged for
  human sign-off) naming the mode and the single override affordance
  `docx.protection.acknowledge_protection(document)` (document-level,
  in-memory, never persisted). Protection is reported, never removed; no
  stripping verb exists; upstream APIs untouched. The scrubbed gauntlet is
  LO-smoke verified.

### v0.1 (2026-07-08) — honesty recall, everyday shapes, four verbs

`agent_docs/PLAN-v0.1.md`, driven by the 83-finding gap review. Standing job
evals (`tests/paper/test_jobs.py`) with the strict-xfail ledger discipline —
all markers flipped green by completion. Phase 0 honesty recall: tracked
moves and format-change revisions enumerated/counted/view-filtered (never
resolved silently; `remaining_unsupported` census), `Span.in_field` +
field-result refusal, placeholder-aware form fill, control-char refusals, no
fake bullets, untracked-in-insertion and bookmark-hollowing guards,
blind-region confession (outline schema v2), `docx.package.diagnose`.
Phase 1: cell-wise table guards, transparent Word noise in block ops,
`replace_all`, same-author redline layering, break-tolerant narrowing,
hyperlink-interior redlines. Phase 2: `docx.controls`, numbering authoring
(`ensure_*_definition`/`restart_numbering`), `insert_blocks_after` typed
blocks, `Span.comment` + `docx.commentops` (w15 replies/resolution),
`docx.package.text_diff`/`pending_changes`. Additional additive upstream
touch: commentsExtended part registration in `docx/__init__.py`.
