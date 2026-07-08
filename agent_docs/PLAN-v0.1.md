# paper-docx — v0.1 Plan: the honesty recall and the everyday-shapes wave

**Status:** accepted plan. Successor to `PLAN-paper-docx.md` (v0, complete).
`CONVENTIONS.md` remains the law; everything here is still purely additive to
upstream, still gated by both upstream suites, still refusal-atomic.

**Provenance.** v0 shipped, then a five-lens adversarial gap review (83 raw
findings, ~40 distinct) found something worse than incompleteness: **v0 lies
in specific, reproducible ways** — it reports documents clean that Word shows
as revised, doubles moved text in every view, "fills" forms Word still treats
as empty, and lets edits land in field results that Word will regenerate.
This plan is structured around that finding.

**The governing principle (do not reorder these phases):** a capability added
on top of false state inherits the false state. Honesty fixes come first —
not because they are the flashiest, but because every later verb (and every
agent) trusts the perception layer and the resolution reports. Phase 0's bar
is *never report false state*, which is much cheaper than *support
everything*: most items are detect-and-refuse, not full support.

**The standing rule for every item:** probe → frozen fixture → failing test →
fix. No fix merges without a fixture reproducing the lie and a test that
fails before the fix. Upstream pytest + behave stay green throughout.

---

## Phase E — Job evals (lands before any Phase 0 fix merges)

The gap review's panel found what 326 organ tests didn't, because the test
suite tests *organs* and the panel tested *jobs*. Convert the three simulated
jobs into standing scenario tests under `tests/paper/test_jobs.py`:

- **Multi-round redline:** find clause → tracked replace → second reviewer
  layers edits → enumerate revisions → resolve → verify text equals the
  intended outcome in both accept and reject worlds.
- **Form fill:** fill placeholder content controls, `{{token}}` replacement
  across fragmented runs, text-box and table-cell targets.
- **Report assembly:** insert a section (heading + paragraphs; richer content
  as Phase 2 lands) after an anchor, verify structure and changed-part
  budget.

Definition of done for every job step, now and forever: **green or
explicitly-refusing — never silently wrong.** Steps that today produce false
state are written as `xfail`/refusal assertions and flip green as phases
land. These evals are the regression net for everything below.

---

## Phase 0 — Honesty recall

Bar: never report false state. Explicit **non-goal**: full tracked-move and
format-change *resolution* (accept/reject of moves is Phase 2+; anyone
expanding Phase 0 scope toward it is scope-creeping — stop them).

- **H1 — Tracked moves enter the perception model (detect, count, filter —
  not resolve).** Parse `w:moveFrom`/`w:moveTo` (+ their range markers) far
  enough to: (a) treat `moveFrom` as deletion-like and `moveTo` as
  insertion-like in story/search view filters — `view="current"` stops
  doubling moved text, `view="original"` stops omitting it; (b) enumerate
  them in `doc.revisions` as `revision_type="move_from"/"move_to"` with
  author/date/text; (c) count them in `blind_region_counts`. This also cures
  the anchor poisoning the review found: with correct view text, block
  hashes and span staleness detection (load-bearing for every write API) are
  computed over what Word actually shows.
  *Note: this goes one step beyond the accepted feedback (which asked only
  for enumerate/count/refuse). Leaving text doubled in the views would
  itself be reporting false state — the view filter IS the honesty fix, and
  it is ~a dozen lines in the existing visitor. Resolution stays out.*
- **H2 — Format-change revisions confessed.** Enumerate the `*Change` family
  (`w:rPrChange`, `w:pPrChange`, `w:tblPrChange`, `w:sectPrChange`,
  `w:numberingChange`, and table `cellIns`/`cellDel`/`cellMerge`) generically
  as `revision_type="format_change"` (element tag recorded); count them in
  `blind_region_counts`.
- **H3 — Resolution never claims more than it did.** `accept_all()` /
  `reject_all()` validate the **selected set first** (atomicity): if the
  selection contains revision types they cannot resolve (moves, format
  changes, table-structure revisions), raise `UnsupportedStructureError`
  before touching anything. A filtered call whose selected set is fully
  resolvable proceeds; it claimed nothing about the rest. The returned count
  and `to_dict()` gain a `remaining_unsupported` census so "doc is clean"
  can never be inferred from a successful call that left revisions behind.
- **H4 — Fields become visible and protected.** Detect both field shapes
  (`w:fldSimple` wrappers and `w:fldChar begin/separate/end` + `w:instrText`
  runs). `Span` gains `in_field`; `Span.replace` refuses field-result edits
  (`UnsupportedStructureError`: the edit would vanish on the next field
  update); story blocks flag field presence; `blind_region_counts` counts
  fields.
- **H5 — Form fill actually fills.** `Span.replace` into a content control
  showing placeholder text clears `w:showingPlcHdr` and the
  `PlaceholderText` run style. Full fix, not detect-and-refuse — it is
  cheap, and it is the down payment on Phase 2's content-control API.
- **H6 — Control characters refused in written text.** `\n`/`\t`/`\r` in
  `Span.replace` new text and in block-op paragraph strings raise
  `ValueError` (programmer error, per CONVENTIONS §2) telling the caller to
  pass separate paragraphs / use future rich insertion. Cite the asymmetry
  in the PR: the search side already refuses spans crossing `w:br`/`w:tab`
  precisely because replacing across one would silently drop it — writing
  one in must not be allowed either.
- **H7 — No fake bullets.** `apply_list_style` validates that the style's
  effective `numId` resolves to a live definition in the numbering part;
  refuses loudly (`TargetNotFoundError`) otherwise. (Creation of
  definitions is Phase 2 — this item only removes the silent failure.)
- **H8 — Untracked edits stop rewriting history and hollowing anchors.**
  (a) Untracked `Span.replace` intersecting a pending `w:ins` refuses
  (`UnsupportedStructureError`) — the tracked path already refuses the
  straddle for exactly this reason; the untracked path currently fabricates
  attribution silently. (b) Replace refuses when a **non-empty named
  bookmark's entire range** lies inside the span (hollowing the target of
  every `REF`/`PAGEREF`/TOC entry). Point/empty bookmarks (`_GoBack`) are
  explicitly transparent — this carve-out is required or Phase 1's
  noise-tolerance work would contradict this item.
- **H9 — "Visibility-complete" stops over-claiming.** `blind_region_counts`
  confesses what traversal cannot read: `math`, `embedded_objects`
  (SmartArt/charts/OLE), `alt_chunks`, `hidden_text` (`w:vanish`), plus the
  new `moves`, `format_changes`, `fields` counts from H1/H2/H4. Reading
  those regions is tail work; *confessing* them is Phase 0.
- **H10 — Typed triage for unreadable input.** New additive API
  `docx.package.diagnose(path) -> PackageDiagnosis` (readable? encrypted?
  macro-enabled `.docm`? missing required parts? which check failed), with
  `.to_dict()`. *Adjusted from the accepted feedback:* changing the
  exception types `docx.Document()` raises would alter existing-API behavior
  (§1.1 strict superset), so triage ships as a new surface and the raw
  upstream errors are documented as "call `diagnose()`".

Fixtures: moves and `rPrChange` can be bootstrapped as hand-built XML
(`generated` provenance, honestly labeled), but LibreOffice converts moves to
plain ins/del on round-trip — so a **real-Word tracked-move + format-change
fixture joins FIXTURE-REQUESTS.md as the top request**, alongside a real
TOC/fields fixture and a placeholder-control form.

## Phase 1 — Everyday shapes (guard relaxation)

The Tier 2 refusals that fire on shapes real documents carry by default.
Mostly relaxing over-broad guards — cheaper than it looks, but every
relaxation gets a corruption-attempt test proving the narrower guard still
holds.

- **S1 — Cell-wise table guards.** One merged cell no longer poisons the
  whole table: `update_cell` refuses only if the *target* cell participates
  in a merge or holds a nested table; row insert/delete refuse only if the
  *affected rows* intersect a vertical merge. Merged header row + plain data
  rows — the default shape of real tables — becomes editable.
- **S2 — Transparent markup in block ops.** Tracked delete/replace tolerate
  `w:proofErr`, point bookmarks (`_GoBack`), and comment range markers as
  transparent noise (preserved, not treated as unsupported structure). Word
  scatters these through virtually every saved document; today they make
  every commented clause and every TOC-referenced heading un-deletable.
  Real-Word fixture required (LibreOffice does not emit `proofErr`).
- **S3 — `replace_all`.** New API with **pinned invalidation semantics,
  settled in this phase's PR: recommendation is single-scan,
  apply-in-reverse-document-order** (later matches first, so earlier atom
  offsets stay valid; deterministic; no span-version machinery). Result
  reports per-match outcomes. The ten-placeholder form-fill case from the
  job evals is the acceptance test.
- **S4 — Second-round redlines.** Same-author layering: a tracked replace
  straddling that author's own pending `w:ins` extends/merges the insertion
  instead of refusing. Cross-author straddles keep refusing (fabricated
  history), with the error message teaching the accept-then-re-mark
  workflow. Design note in PR before implementation.
- **S5 — Tab/break-tolerant editing.** Replace spanning `w:tab`/`w:br`
  becomes supported by preserving the break elements in place and editing
  the text segments around them (equivalently: auto-splitting the
  replacement at the break positions when lengths permit, else refusing as
  today with a message naming the segments).
- **S6 — Hyperlink-interior tracked replace.** Emit `w:ins`/`w:del` inside
  the `w:hyperlink` container (Word supports this); refusal narrows to
  spans *crossing* the hyperlink boundary.

## Phase 2 — New verbs (only four make the cut) + one promotion

- **V1 — Content controls become a real surface.** Enumerate
  (alias/tag/type/current value), `get`/`set_control_value` by tag or
  anchor; checkbox, dropdown/combo (validated against `w:listItem`
  choices), date, and plain/rich text controls; placeholder state handled
  (H5's machinery). Data-bound (`w:dataBinding`) controls refuse with an
  explanation. This is the templating primitive.
- **V2 — Minimal guarded numbering authoring.** Ship exactly two canonical
  definitions (one clean decimal, one bullet) clonable into
  `word/numbering.xml`, plus `w:lvlOverride`/`w:startOverride` restart
  support. Everything exotic (custom level text, legal numbering, images as
  bullets) refuses. Closes the literal "cannot make a real bullet" gap.
- **V3 — Rich block insertion.** `insert_section_after` (and a new
  `insert_blocks_after`) accept a **small typed block list** — paragraph
  (with bold/italic runs), bullet/numbered list (via V2), simple table —
  not arbitrary richness. This was the wall that ended the report-assembly
  job in the review. Images and page/section breaks stay tail items.
- **V4 — Span-anchored comments.** `span.comment(text, author, initials)`
  bridging to the upstream v1.2.0 comment machinery (which needs runs —
  spans know theirs); read side gains anchored-text reporting, and
  replies/resolution ride the `w15` extensions (`commentsExtended`) as far
  as read + reply + resolve toggling.
- **V5 — Human-readable diff, promoted from the reference.** An honesty
  recall deserves a verification lens humans can read: promote
  `diff_docx.py` semantics into the package (`docx.package.text_diff(a, b)`
  or similar — name settled in PR): per-story text diff over the same view
  machinery, plus revision-aware summary ("what would change if accepted").
  `diff_package` keeps answering *which parts*; this answers *what*.

## Tail — triggered entries (one line each; no phase until triggered)

| Item | Trigger |
|---|---|
| Full move/format-change **resolution** (accept/reject) | first real workflow blocked by H3 refusals |
| Scoped find (`within=` block range / heading section) | first agent job needing clause-scoped disambiguation `near=` can't express |
| Cross-document assembly (append, copy blocks with style/numbering/rel remapping) | first report-assembly job sourcing from more than one file |
| Hyperlink/footnote/endnote/field authoring; image swap + alt text; table furniture (widths, header rows); column ops | per-verb, on first concrete demand |
| Protection & track-changes settings awareness (`w:documentProtection`, enforced tracking) | first deployment editing docs from a DMS that sets them |
| Case-sensitive / typography-preserving matching modes | first defined-terms redline miss traced to casefolding |
| Locale style ids (Überschrift2 …) | first non-English template in production |
| Unicode hygiene in matching (NFC, zero-width folds) | first miss traced to invisible characters |
| Reading math/SmartArt/chart/altChunk text | first corpus where H9 counts are non-zero and content matters |
| Performance (indexed find, 500-page docs) | first profile showing find dominating a job |
| Pagination/layout oracle | first signature-page or page-reference verification need |

## Definition of done, v0.1

Every Phase 0 item: the lie is reproducible in a frozen fixture, a test
fails before the fix, and afterward the same input yields either true state
or a typed refusal. Every Phase 1 relaxation: the previously-refused
everyday shape works AND a corruption-attempt test proves the narrowed guard
still refuses the genuinely unsafe case. Every Phase 2 verb: contract
harness (save→reopen, changed-part budget, refusal atomicity, LO smoke) plus
a job-eval step exercising it. All three job evals green-or-refusing.
Upstream pytest and behave identical to baseline. `PAPER.md` entries per
phase; API-PROPOSAL.md amended in the same commit as any signature change.

## Notes on the accepted feedback (where this plan adjusts it)

1. **H1 goes further than "enumerate and refuse":** view filtering for moves
   is included in Phase 0 because doubled text in `outline()` is itself
   false state — the feedback's own bar — and the fix is small. Resolution
   remains excluded exactly as directed.
2. **H10 ships as an additive `diagnose()`** rather than retyping
   `docx.Document()` exceptions: changing exceptions on the existing entry
   point is a behavior change §1.1 forbids.
3. **H8's bookmark refusal carries an explicit `_GoBack`/point-bookmark
   carve-out**, without which it would contradict Phase 1's S2
   (noise-transparency) and re-refuse everything S2 unblocks.
4. **H3 keys refusal to the selected set**, not the whole document, so
   author-filtered resolution of clean subsets keeps working; the census
   field carries the rest of the honesty load.
