# paper-docx — v0 Implementation Plan

## What this repository is and why you're here

This is Paper Instruments' hard fork of **python-docx**, the standard Python library for
Word documents — roughly ninety million downloads a month and the quiet workhorse inside much of
the document-AI ecosystem. The fork point is upstream's latest release tag (see the `paper-base`
git tag and `PAPER.md`); packaging is renamed (`pip install paper-docx`) while the import name
`docx` is **frozen forever**, because this fork must remain a drop-in replacement for every
existing snippet, pipeline, and model prior that says `from docx import Document`.

Why fork at all: upstream's core is excellent — a lossless package layer that round-trips
content it doesn't understand, a disciplined declarative XML layer, a decade of absorbed
real-world edge cases — but its **editing surface** stalled years ago. There is no
tracked-changes support at all. There is no way to replace text that Word has fragmented across
runs, which is most text in real documents. There is no numbering creation. And standard
traversal is blind to text inside tracked insertions, content controls, and text boxes — the
library literally cannot see parts of the visible document. Production agent systems therefore
fall back to fragile raw-XML surgery for exactly the operations that matter most, and the
dominant failure mode of that surgery is **silent corruption**: files that open fine and are
quietly wrong.

What came before you: months of production harness work built and battle-tested helper scripts,
verifiers, and safe-editing workflows around the stock library. That work established which
operations matter, which invariants keep files safe (refusal atomicity, narrow saves,
save-reopen verification), and where the landmines are. All of it is distilled into
`reference/office-transfer/` in this repo. It is **executable specification** for what you will
build — mine its algorithms, normalization tables, refusal conditions, and invariants; improve
it where it is best-effort; do not paste it into `src/`.

Your mission: extend this fork into a **strict superset** — new organs implemented in upstream's
own architectural idiom, proven against a frozen fixture corpus, with zero changes to existing
behavior (v0 is purely additive; `CONVENTIONS.md` §1.1). End state: this package replaces stock
python-docx in our agent environments; everything that worked before still works; the dangerous
operations become safe first-class APIs; the invisible regions become visible.

**Read first, in order:** `CONVENTIONS.md` (the governing document — it wins over anything
here) → `reference/office-transfer/README.md` → `reference/office-transfer/skill/SKILL.md`
(read as user stories, not instructions) → `skill/references/*.md` → upstream
`docs/dev/analysis/` for each feature area as you reach it → `git log` of the comments (v1.2.0)
and hyperlink (v1.1.0) features as structural templates → upstream test layout (both pytest and
behave suites).

---

## Mining map (reference → what to extract → target)

| Reference file | Extract | Feeds |
|---|---|---|
| `office_helpers/ooxml_util.py` | canonical-compare semantics incl. meaningful-whitespace preservation; hashing approach. NOT the ordered-insertion tables (descriptors' `successors=` replaces them) | Phase 2 kernel |
| `office_helpers/package.py` | part-map reading, XML-aware diff, compare-based patch_save algorithm | Phase 2 kernel |
| `docx_helpers/outline.py` | story-part enumeration (body, headers, footers, footnotes, endnotes, comments, content controls, text boxes, w:ins/w:del), blind-region reporting, block schema | Phase 3 traversal |
| `docx_helpers/text.py` | normalization table (smart quotes, dashes, minus, nbsp, soft hyphen, tabs, case); normalized-match → XML text-atom mapping; `replace_text` run-preservation; `tracked_replace` minimal-span + w:ins/w:del/w:delText emission; every refusal condition | Phases 4–6 |
| `docx_helpers/block_ops.py` | same-parent safety rule; tracked paragraph-range replace; pPr preservation on tracked paragraph delete; anchor-relative insertion | Phase 7 |
| `docx_helpers/revisions.py` | revision enumeration across story parts; accept/reject semantics incl. author filter | Phase 8 |
| `docx_helpers/tables.py`, `numbering.py` | find-near-text; guards; apply-existing-numbering path | Phase 9 |
| `verify_docx.py` | mechanical checks → package regression tests (undefined styles/numbering IDs, broken rels, changed-part budgets, fake-bullet detection) | Phase 1 & ongoing |
| `SKILL.md`, `references/*.md` | user stories, sequencing expectations, the anti-pattern list | API design, docstrings |

---

## Phase 0 — Orientation (no code)

Map the three layers in the actual tree (locate by role — grep class names — don't trust
remembered paths): the opc package machinery; the oxml element-class + descriptor system and its
registration point; the api proxy layer. Identify how the comments feature (v1.2.0) is built end
to end — part, oxml classes, proxies, tests — it is your template for every organ here. Run both
upstream suites; confirm they match the baseline recorded in `PAPER.md`. Deliverable: a short
`ARCHITECTURE-NOTES.md` (10–20 bullets) proving you can name where each future organ will live.

## Phase 1 — Test infrastructure (first-class; nothing merges before it)

Implement CONVENTIONS §4:

- `tests/paper/fixtures/` + `MANIFEST.sha256` + sidecars (pinned schema in CONVENTIONS §4).
  Feature-isolated fixtures needed at minimum: tracked insertions AND deletions (two distinct
  authors); comments; a content control wrapping text; a text box with text; footnotes and
  endnotes; header/footer text across multiple sections; a merged-cell table and a nested
  table; a custom numbering definition in use; fragmented-run text containing smart quotes and
  an en-dash amid a dollar range (the classic `$75–100/hr` case); one gauntlet combining all of
  it; one corrupt-by-construction file. Bootstrap with LibreOffice-authored files, label
  provenance honestly, and write `FIXTURE-REQUESTS.md` for the real-Word versions a human must
  author.
- The contract harness (five assertions, CONVENTIONS §4) as shared conftest utilities.
- Frozen-clock utility for anything that stamps `w:date`.
- LO smoke helper with the `lo_smoke` skip marker.
- Port the mechanical `verify_docx` checks that assert package facts into plain pytest helpers.

## Phase 1.5 — PR-0: API Proposal (CONVENTIONS §8)

One markdown: exact signatures, return types, refusal conditions, and short examples for every
organ below, grounded in what Phase 0 actually found. Confirm the pinned CONVENTIONS shapes
(§2 exceptions and anchors, §4 sidecar schema, §7 kernel) against the real code, flagging any
mismatch for human decision. Humans approve before Phase 2 begins.

## Phase 2 — Package kernel (`docx.package`)

Implement CONVENTIONS §7 exactly: `xml_equivalent`, `diff_package` → typed `PackageDiff`,
`patch_save(original_path, document, out_path)` — compare-based, additive. Required invariants
and tests are pinned there; do not skip the meaningful-whitespace trap test (`w:t` trailing
space), the no-op byte-identity test, the zip-determinism policy, or the mid-write
failure-injection test.

## Phase 3 — Story-part traversal & inspection (opt-in APIs only)

New, explicitly named APIs — existing traversal semantics are untouched. Intent: enumerate every
story part (body, headers, footers, footnotes, endnotes, comments) and every visibility region
within them (plain, inside `w:ins`, inside `w:del`, inside content controls, inside text boxes),
yielding blocks with stable anchors (CONVENTIONS §2) and region flags. This is the perception
layer every later organ targets. An outline/inspection object with a golden-tested `.to_dict()`
JSON shape. Sidecar-driven tests: counts and text per fixture match hand-verified ground truth;
the gauntlet proves nothing visible is missed.

## Phase 4 — `find_text` and the Span object

Package-wide normalized search over Phase 3's traversal. Port the normalization table exactly
and test each rule with a fixture pair. Matching must assemble text across fragmented runs and
across the paragraph boundary (block ops will consume cross-paragraph spans; character edits
won't). Disambiguation: nth occurrence, near-anchor, story-part scoping. The Span is the pivotal
new object: it maps a visible-text interval back to concrete text atoms and carries its anchor.
Test the mapping *by use* — perform a replace through a returned span and assert the outcome —
not by inspecting offsets. Refusal: ambiguous match without disambiguators
(`AmbiguousTargetError`).

## Phase 5 — Run-preserving replace (`span.replace(new_text)`)

Splits runs at match boundaries; untouched runs keep byte-identical `rPr`; boundary fragments
inherit their original run's `rPr`. Refusals (from the reference, kept): span crosses deleted
text or field-instruction text; span crosses a content-control boundary
(`BoundaryViolationError`). Contract harness on the fragmented-run fixture, including a replace
spanning a bold→italic formatting transition. Invariant: replace(x→y) then (y→x) restores text
and formatting.

## Phase 6 — Tracked replace (`span.replace(new_text, tracked=True, author=..., date=...)`)

New oxml classes for the revision vocabulary (`w:ins`, `w:del`, and `w:delText` handling) via
the descriptor system, registered properly — this is the first PR that adds XML vocabulary, so
imitate the comments-feature commit structure. Semantics from the reference, kept and tested:
minimal changed span via common prefix/suffix trim (the redline marks `75-10 → 85-11`, not the
sentence); deleted text lives in `w:delText`, never live `w:t` inside `w:del`; unique revision
ids; injectable clock for `w:date`; `rPr` preserved on both sides; fragmented same-paragraph
targets supported; cross-paragraph targets refused (Phase 7's job).

## Phase 7 — Block operations

Anchor-relative composition and clause-level redlines: insert heading+paragraphs after an
anchor; tracked-insert a section; tracked-delete whole paragraphs (preserving paragraph
properties per the reference's approach); tracked-replace a paragraph range. The load-bearing
safety rule to keep: all selected paragraphs must share one parent — refuse edits spanning
unrelated story regions, table boundaries, or content-control boundaries. Changed-part budget
assertions matter here; this is where careless implementations churn the package.

## Phase 8 — Revisions (`doc.revisions`)

Enumeration across all story parts (type, author, date, text, anchor); accept / reject, all or
filtered by author. Then wire the tracked-edit algebra invariants from CONVENTIONS §4 — they
cross-check Phases 5, 6, and 8 against each other and are the highest-value tests in the repo.

## Phase 9 — Tables and numbering (narrow, guarded)

Tables: find-near-text; cell update routed through Phase 5's replace; row insert (copy format
from neighbor) and delete; **loud refusal** on `vMerge`/`gridSpan`/nested-table complexity
(`UnsupportedStructureError`) rather than any attempt at cleverness. Numbering: enumerate
definitions; apply an existing definition or list style to paragraphs; continue/restart where
the reference supports it; refuse when no compatible definition exists. Authoring new
`numbering.xml` definitions is explicitly out of v0.

---

## Order and dependencies

0 → 1 → 1.5 → 2 → 3 → 4 → 5 → 6 → 7 → 8 → 9. The kernel precedes all organs because the
contract harness's changed-part assertion depends on it. Traversal precedes find; find precedes
both replaces; tracked precedes block ops (shared revision vocabulary) and revisions.

## Prohibitions (repo-specific, beyond CONVENTIONS)

- No changes to existing traversal semantics; no changes to the plain-text setter; no changes to
  `save()`.
- No new runtime dependencies.
- No porting of CLI wrappers, print-based reporting, or the unpacked-package workflow into
  `src/`.
- No whole-document schema validation in tests; fragment-scoped only.
- No hand-built lxml in proxy code.

## Ask-for-help triggers

Real-Word fixture needs (via `FIXTURE-REQUESTS.md`); any organ that seems to need a behavior
change; any upstream test that starts failing; any place PR-0 signatures prove wrong in
implementation; any refusal condition you're tempted to soften to make a test pass.
