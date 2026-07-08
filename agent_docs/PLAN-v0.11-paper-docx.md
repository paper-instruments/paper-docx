# paper-docx — v0.11 Implementation Plan

## What v0.11 is and why you're doing it

v1 made this package **honest**: it sees everything visible, edits surgically, and — critically —
refuses rather than lies. v0.11 converts the most important of those refusals into capabilities,
and then builds the three organs that turn a single-document editor into a document *pipeline*:
resolution of every revision type Word actually produces, a compare engine that *generates*
redlines, and cross-document composition. Together with field authoring and a formatting
resolver, this is the release after which the five core jobs — deck/report from spec, refresh,
redline round-trip, template fill, report assembly — either complete or refuse loudly, with no
silent middle.

The organizing logic, so you can make judgment calls consistently:

1. **Revision completion is first because everything downstream depends on it.** Scrub cannot
   certify a clean file if any revision type is unresolvable. Compare must consume documents that
   may carry pending revisions. Unattended redline finalization is gated entirely on this.
2. **Compare is the trust mechanism generalized.** Once the package can diff two documents and
   emit native tracked changes, *any* agent editing session — however the edits were made —
   can be rendered as one reviewable redline. It also decouples edit speed from edit
   reviewability.
3. **Composition is how professional documents are actually made.** Nobody starts from scratch:
   proposals are assembled from clause libraries, reports from section documents. Copying
   formatted content between documents without corruption is style/numbering/relationship
   reconciliation — exactly the package-level, corruption-prone mechanics this fork exists to own.
4. **Fields are formulas; static text is a pasted value.** A cross-reference that renumbers
   itself is the difference between a draft and a deliverable — and between a correct contract
   and a silent legal error after clauses shift.

**Scope rule (write this on your wall):** converting a documented typed refusal into a correct
operation is the *sanctioned growth path* of this package and is NOT a behavior change under
CONVENTIONS §1.1. Every such conversion gets a one-line PAPER.md entry. Changing the behavior of
any currently-*successful* operation remains forbidden.

**Precondition:** this plan assumes the v1 wave (honesty recall, everyday shapes, the four verbs)
is merged. Verify against PAPER.md and the code before starting. If anything you depend on below
is missing, stop and report it as a Phase 0 blocker — do not build around it.

**Read first:** CONVENTIONS.md (still the law) → PAPER.md (what actually shipped in v1) → the
v1 gap-review dossier in agent_docs/ (your findings are cited below by tier/number) →
`docs/dev/analysis/` for revision handling and fields → the reference material's
`diff_docx.py` and any unmined redline notes.

---

## Mining map

| Source | Extract | Feeds |
|---|---|---|
| v1 revision enumeration (moves/rPrChange detect-and-refuse) | the detection machinery — you are upgrading refusal → resolution | Phases 1–2 |
| v1 Span + tracked_replace | atom mapping, minimal-span trim, revision-id allocation, injectable clock | Phase 4 (compare emission) |
| v1 block_ops | same-parent safety, tracked paragraph insert/delete mechanics | Phases 4–5 |
| reference `diff_docx.py` | semantic text-diff approach, block alignment instincts | Phase 4 |
| v1 content-control + numbering organs | placeholder handling, definition application — composition must remap both | Phase 5 |
| gap-review Tier 1 items 1–2, 8 | the exact probe fixtures for moves, rPrChange, row revisions, paragraph-mark cases | Phase 0 fixtures |
| `docs/dev/analysis/` + ECMA-376 Part 1 §17.13 (revisions), §17.16 (fields) | authoritative semantics — establish mechanics from spec + probing, not memory | all phases |

---

## Phase 0 — Fixtures and verification of the ground you stand on

The critical path of this entire release is a **real-Word-authored multi-round redline corpus**.
LibreOffice does not faithfully produce moves, format-change markup, or proofing noise. Write
`FIXTURE-REQUESTS.md` entries immediately — this is the one external dependency that can stall
Phases 1–2, so it goes out on day one:

- A document edited across two rounds by two authors with track changes on, containing: at least
  one drag-move and one cut-paste move; bold/italic applied to existing text (rPrChange); a
  paragraph-style change (pPrChange); a table with one inserted and one deleted row; a deleted
  paragraph mark (two paragraphs merged under tracking); comments anchored inside tracked
  regions.
- The same document with revisions accepted in Word (ground truth for resolution).
- A document pair (original, revised-clean) produced by editing with track changes OFF — the
  compare fixture, with Word's own Compare output saved as a third file for reference.
- A document with editing restrictions enforced via Word's Restrict Editing pane (forms-only
  protection, and a read-only variant) — the protection-awareness fixtures for Phase 3.

Bootstrap what you can in LibreOffice with honest provenance labels; sidecars per CONVENTIONS §4.
Also verify v1's detection actually fires on these files before building on it.

## Phase 1 — Format-change and structural revision resolution

Upgrade `doc.revisions` and accept/reject from enumerate-and-refuse to full resolution for:

- **Run format changes** (`w:rPrChange`): accept = drop the stored previous properties, keep
  current; reject = restore stored properties. Same pattern for **paragraph format changes**
  (`w:pPrChange`).
- **Table row revisions** (insert/delete markers in row properties): accept row-insert = keep
  row, remove marker; accept row-delete = remove the row. This closes the "ghost rows" defect
  from your own Tier 1 findings.
- **Paragraph-mark revisions** (the pilcrow itself inserted/deleted, carried in the paragraph
  mark's run properties): accepting a deleted paragraph mark **merges the paragraph with its
  successor** — subtle, test it explicitly both directions.
- Table-property and cell-level revision markup (property changes, cell merge revisions): if not
  resolved in v0.11, they must be *enumerated and refused by name* — never absent from
  `doc.revisions`, never silently passed by accept_all.

Invariant: after `accept_all()` or `reject_all()` succeeds, a rescan finds zero revision markup
of any kind, and the file matches the Word-accepted ground-truth fixture in visible text.

## Phase 2 — Move resolution

`w:moveFrom`/`w:moveTo` runs with their paired range markers, matched by move name. Model a move
as ONE revision object with two sites. Accept = destination text becomes plain, source range
removed; reject = source restored to plain, destination removed. Orphaned or cross-story pairs →
typed refusal naming the orphan (never partial resolution — refusal atomicity applies to the
pair as a unit). Anchors overlapping either site keep the v1 flag until resolution. After this
phase, the v1 "refusing to report clean" message should be reachable only on genuinely malformed
markup.

## Phase 3 — Scrub / finalize (depends on 1–2)

The compliance verb: produce a file safe to send externally.

- `doc.finalize(revisions="accept"|"reject")` — total resolution or typed refusal naming what
  blocked it (defensive; post-Phase-2 this should be rare).
- `doc.scrub(...)` with explicit, individually-toggleable targets: comments (all comment parts
  and their extended/threading parts), core/app/custom metadata and personal info, the
  track-changes-enabled setting, optional RSIDs, optional hidden text. Document protection
  settings: **report always, never remove** — there is no protection-stripping verb in v0.11
  (removing enforcement has legal implications; the override mechanism below governs editing,
  not stripping).
- Returns a `ScrubReport` (typed, `.to_dict()`, goldenable) itemizing everything removed.
- Tests: scrubbed gauntlet reopens clean, LO smoke passes, a rescan finds zero revision markup,
  zero comment parts, empty personal-info fields; changed-part budget matches the report exactly.

**Also in this phase (independent of 1–2; may land earlier as its own small PR):
protection-aware mutations.** Documents can carry an enforcement setting
(`w:documentProtection`: read-only, forms-only, comments-only, tracked-changes-enforced) that
Word honors and this library currently ignores — editing a forms-locked legal template without
noticing is a fail-loudly violation of the same species as the v1 honesty recall, and that
precedent is the license for this change. Contract: **this package's own mutating APIs** (spans,
tracked ops, block ops, table ops, scrub/finalize) check the setting and raise a typed refusal
naming the protection mode, with one explicit, documented override affordance (exact shape —
per-call kwarg vs. one document-level acknowledgment — is a PR-0-style proposal decision; the
refusal message must name the override path). Scope it precisely: **upstream APIs are untouched**
— upstream code editing a protected document behaves exactly as before, preserving the strict
superset; our organs were born with this contract. The refusal docs state plainly that
protection is advisory, not security. New exception subclass (e.g., `DocumentProtectedError`)
extends the pinned taxonomy — flag it for human sign-off in the PR and register it in
CONVENTIONS when approved. Ledger the whole item in PAPER.md.

## Phase 4 — Compare: generate a native redline from two documents

`compare(original, revised, *, author, date=None, granularity="word")` → a new document whose
tracked changes transform original into revised.

Pipeline guidance (establish exact mechanics from spec + your v1 machinery):
(1) If either input carries pending revisions, refuse unless the caller passes an explicit
materialize policy (which routes through Phase 1–2 resolution on working copies).
(2) Per story: align blocks by normalized-text fingerprint (LCS/patience over paragraph
fingerprints); unmatched blocks become whole tracked paragraph insertions/deletions using the v1
block-ops mechanics.
(3) Matched blocks: word-level token diff, mapped to atoms via the v1 span machinery, emitted as
minimal `w:ins`/`w:del` with the caller's author/date and the injectable clock.
(4) Tables: align rows by fingerprint → row-revision markup from Phase 1; matched rows recurse
cell-wise.
(5) Formatting differences on identical text: **detect and report** in v0.11 (a findings list on
the result); emitting rPrChange markup is behind a default-off flag.
(6) Changed images/objects: report-only.

The invariants that make this organ trustworthy — implement them as tests before the engine:
- `accept_all(compare(A, B))` ≡ B in visible text, across every story.
- `reject_all(compare(A, B))` ≡ A.
- `compare(A, A)` yields zero revisions.
- Deterministic: identical inputs → byte-identical output (goldenable).
Validate against Word's own Compare output fixture for reasonableness (not byte parity — Word's
alignment choices differ; visible-text equivalence is the bar).

Declared limits, stated in the API docs: compare emits insertions/deletions only (it does not
synthesize move markup); no cross-story move detection; a perf budget with a typed refusal above
a documented size.

## Phase 5 — Cross-document composition (PR-gated design)

This is the largest organ. Before implementation, an API-proposal PR (same protocol as PR-0)
covering:

- `doc.insert_blocks_from(source_doc, source_range, anchor, *, styles=..., numbering=...,
  media=...)` and `append_document(source, *, section=..., headers=...)`.
- **Style reconciliation** — the hard core. Two modes: `match_by_name` (map source styles to
  destination styles of the same name; destination definition wins — content adopts the house
  look) and `import_renamed` (clone colliding-but-different definitions under new ids/names and
  remap references — content keeps its source look). basedOn chains import transitively;
  a style table diff appears in the operation's report.
- **Numbering remap**: colliding numbering ids get fresh ids with deep-copied definitions;
  paragraph references remapped; document the continue-vs-restart semantics you choose.
- **Relationships/media**: images copied as new parts with fresh rel ids; hyperlinks recreated;
  embedded OLE objects → typed refusal in v0.11 (declare it).
- **Bookmarks**: rename on collision, remap REF fields inside the copied range, ids reallocated.
- **Source revisions/comments** inside the copied range: default is refuse-unless-materialized
  (route through Phase 1–2), with carry-through as an explicit option only if cheap.
- **Append semantics**: section break policy; v0.11 keeps destination headers/footers
  (keep-source is a declared future mode).
- Changed-part budgets are *declared per operation* in the returned report (composition legally
  touches styles.xml, numbering.xml, rels, media — the harness assertion is report-matches-diff,
  not small-diff).

Acceptance is job-shaped, not just unit-shaped: assemble a proposal from a base doc + two clause
files + a CV section with tables, images, lists, and headings; result reopens clean, LO smoke
passes, lists number correctly, no style explosion (bounded style count), and the whole run is
frozen as a standing eval task.

## Phase 6 — Bookmarks and field authoring (author-and-delegate)

- Bookmark API: enumerate, create on a span, delete; globally unique ids.
- Field authoring: page number and page count (footers), date, TOC (with heading-level switches),
  cross-reference to a bookmark (text / number / page variants). Simple fields via the simple
  form; TOC via the begin/separate/end complex form. Every inserted field carries placeholder
  result text and sets the document's update-fields-on-open flag — **this package authors
  formulas; it never computes their values.** Pagination is a renderer's job; Word or headless
  LibreOffice recomputes on open, and the harness may force an update pass.
- Self-consistency test: the v1 `in_field` guard must recognize fields this phase authors —
  a span landing in our own TOC refuses, same as one landing in Word's.

## Phase 7 — Effective-format resolver (read-only, provenance-bearing)

The perception organ v0/v1 never built: "what formatting does this text *actually* carry."
Resolution chain: document defaults → paragraph-style inheritance chain → character style →
direct formatting, with correct **toggle-property semantics** (bold/italic/caps and friends
XOR through style layers per spec — the famous gotcha; test it with a nested-toggle fixture).
Every value carries a provenance chain per CONVENTIONS §2. Declare honestly what v0.11 does not
resolve (table-style conditional formatting, numbering-mark properties) — "unresolved" is a
legal answer; a wrong guess is not. Deliver `format_of(target)` plus a match-surrounding helper
consumed by Phase 5/6 insertion paths, so inserted content can adopt its neighbors' look.

---

## Order and dependencies

0 → 1 → 2 → 3 → 4 → 5 → 6 → 7, with the PR gate before Phase 5. The chain is real: scrub needs
1–2; compare needs 1–2 (materialization) and v1 spans (emission); composition wants compare's
fixtures and materialization; fields want bookmarks; the resolver is standalone but its
consumers land in 5–6. Phase 0's fixture requests are the schedule's external dependency —
file them before writing any code.

## Prohibitions

- Never compute pagination or field values; never render.
- Compare emits ins/del only in v0.11 — no move synthesis.
- No OLE authoring; composition refuses embedded objects, typed.
- No behavior changes to currently-successful operations; refusal→capability conversions only,
  each ledgered in PAPER.md. (One sanctioned exception this release: the Phase 3
  protection-awareness item converts a *silently unsafe* success into a typed refusal on this
  package's own APIs only — the v1 honesty-recall precedent — with upstream APIs untouched.)
- No protection-stripping verb; no decryption of password-protected files (typed refusal on the
  encrypted container stands).
- **No public document-QA / `check()` API.** Judging arbitrary documents — verifier families,
  layout QA, repair loops — is harness territory and stays out of this package permanently. The
  package's outward obligation is narrower and already largely built: load-time and
  operation-time failures on bad input (corrupt zip, encrypted container, dangling
  relationships) speak as typed, specific refusals — never raw tracebacks. Verify that coverage
  as part of this release; add typed wrapping where any raw error remains.
- No cross-story move resolution; typed refusal.
- All prior CONVENTIONS prohibitions stand (no reformatting upstream files, no new runtime
  deps, fragment-scoped schema validation only, no hand-built lxml in proxy code).

## Definition of done (release-level, beyond per-organ CONVENTIONS §9)

- The redline round-trip job, the compare job, and the proposal-assembly job run as standing
  eval tasks in tests/paper and finish green or with typed refusals — never silently wrong.
- `finalize` + `scrub` on the multi-round real-Word fixture produces a file Word opens with
  zero pending revisions (human checklist item).
- The compare algebra invariants hold on the full corpus.
- PAPER.md updated with every refusal→capability conversion and every declared limit.

## Ask-for-help triggers

Real-Word fixture fulfillment (day-one blocker for Phases 1–2); any revision markup in the wild
corpus not covered by Phase 1–2's taxonomy (enumerate it, refuse it, report it — don't improvise
resolution); compare outputs that diverge from Word's Compare in *visible-text* terms; any
composition case where both style modes produce wrong-looking output; anything tempting you to
compute a field value.
