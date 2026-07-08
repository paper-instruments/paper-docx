# API Proposal (PR-0) — paper-docx v0 organs

Per CONVENTIONS §8: exact signatures, return types, refusal conditions and
examples for every v0 organ, grounded in the actual upstream code (Phase 0,
`ARCHITECTURE-NOTES.md`) and the reference helpers
(`reference/office-transfer/skill/scripts/`). The signature-conformance test
`tests/paper/test_api_surface.py` enforces this document mechanically: as each
phase lands, its surface checks flip from `skip` to enforced. Deviations
discovered during implementation are amended HERE, in the same commit, never
silently.

Everything below is **purely additive** (§1.1). Upstream files receive only:
one import block at the bottom of `src/docx/package.py` (F1), part-class
registrations in `src/docx/__init__.py`, and one `revisions` property on
`Document`. No existing behavior changes.

---

## 0. Resolutions of Phase 0 flags

- **F1 (`docx.package` collision).** The §7 kernel keeps its pinned public
  path: `docx.package.xml_equivalent / diff_package / patch_save`.
  Implementation lives in a new module `src/docx/_paperpkg.py`; upstream
  `src/docx/package.py` gains only a clearly-marked additive import block
  re-exporting the three names (plus the result types). Nothing existing in
  that module is touched or shadowed.
- **F2 (zip determinism).** `patch_save` owns its zip writing: entry order =
  the candidate serialization's order, all entries stamped with the fixed
  date-time `(1980, 1, 1, 0, 0, 0)`, `ZIP_DEFLATED`. Upstream `save()` is
  untouched.
- **F3 (parse-time whitespace).** The kernel parses with its own lxml parser
  (**no** `remove_blank_text`), so `xml_equivalent` sees text nodes exactly as
  stored. The pinned trailing-space-in-`w:t` inequality test runs on raw part
  bytes.
- **F4 (stale `uv.lock`).** Out of v0 code scope; dev workflow uses
  `uv run --no-sync` until the lockfile is deliberately regenerated.

Phase 1 findings folded in: `.rels` parts compare as relationship multisets
and `[Content_Types].xml` as effective (extension→type restricted to in-use
extensions, partname→type) maps — real producers differ freely in order,
whitespace, and inert defaults. Duplicate zip member names: first entry wins,
counted once (documented; full duplicate detection is a future check).

Confirmed pinned shapes against real code: §2 exceptions (`docx/errors.py` is
a free name; upstream uses `exceptions.py`), §2 anchors (below), §4 sidecar
schema (already shipping in Phase 1), §7 kernel (F1 resolution above).

**v0.1 amendments (honesty recall, per PLAN-v0.1.md):** `Outline.to_dict`
version bumps to 2 — `blind_region_counts` gains `moves`, `format_changes`,
`fields` and the confession keys (`math`, `embedded_objects`, `alt_chunks`,
`hidden_text`); blocks gain `has_field`. `Span` gains public `in_field`;
`Span.replace` refuses field-result targets. Story views treat `w:moveFrom`
as deletion-like and `w:moveTo` as insertion-like (moved text appears once
per view). `Revision.revision_type` gains `move_from` / `move_to` /
`format_change` (enumerated, counted, NOT resolvable — individual and batch
resolution refuse); `Revisions.to_dict` bumps to v2 with a
`remaining_unsupported` census, and `accept_all`/`reject_all` validate the
selected set atomically before touching anything.

**v0.1 Phase 1-2 amendments (everyday shapes + verbs, per PLAN-v0.1.md):**
table guards are cell-wise/row-wise (S1); tracked block ops treat
proofErr/point-bookmarks/comment-anchors as transparent (S2);
`docx.search.replace_all` (S3, reverse-document-order semantics, schema
`paper_replace_all` v1); same-author redline layering (S4); break-tolerant
narrowing with whitespace↔tab alignment in kept regions (S5);
hyperlink-interior tracked replace with boundary refusal (S6). New verbs:
`docx.controls` (list/get/set by tag or alias; checkbox/dropdown/date typed;
data-bound refuses), `docx.numbering.ensure_bullet_definition` /
`ensure_decimal_definition` / `restart_numbering` (two canonical definitions,
level-0 restarts; everything exotic still refuses),
`docx.blocks.insert_blocks_after` with the typed `RichParagraph` /
`ListBlock` / `TableBlock` vocabulary (tracked table insertion refuses),
`Span.comment` + `docx.commentops` (anchored text, w15 replies/resolution;
main-story anchors in v0.1), and `docx.package.text_diff` /
`pending_changes` (schema `paper_text_diff` v1) promoted from the reference
`diff_docx.py`. All signatures enforced in `tests/paper/test_api_surface.py`.

**Post-implementation amendments (adversarial review):** duplicate zip member
names collapse with LAST entry winning (`zipfile` semantics — corrected from
this document's first draft which said "first"); `xml_equivalent` refuses
DTD-bearing parts loudly (`UnsupportedXmlError`) and compares prolog/epilog
comments/PIs and PI targets; QName-valued attributes compare textually
(documented limit — prefix rebinding with identical QName text is not
detected; OOXML producers keep standard prefixes and the diff error direction
stays conservative elsewhere).

---

## 1. `docx.errors` — the refusal hierarchy (§2, pinned)

```python
class PaperRefusal(Exception): ...          # base for every safe refusal
class AmbiguousTargetError(PaperRefusal): ...
class TargetNotFoundError(PaperRefusal): ...
class UnsupportedStructureError(PaperRefusal): ...
class BoundaryViolationError(PaperRefusal): ...
class RelationshipPolicyError(PaperRefusal): ...
```

Programmer errors stay `TypeError`/`ValueError` (e.g. `tracked=True` without
`author`). Every refusal is atomic (§1.3): raised **before** any mutation.

## 2. Anchors and typed results (§2, pinned)

```python
@dataclass(frozen=True)
class Anchor:
    story: str          # part name, e.g. "word/document.xml", "word/header2.xml"
    index: int          # 0-based block index within that story
    content_hash: str   # first 8 hex chars of SHA-256 of the block's normalized text
    def to_dict(self) -> dict: ...
```

Raw indices are never accepted alone; the hash detects staleness. All
`.to_dict()` payloads: snake_case keys, deterministic order, top-level
`"schema"` string + integer `"version"`, 0-based indices, lengths in EMU ints.

Dates: every date-stamping API takes `date: datetime | None = None`; `None`
resolves through the injectable clock `docx._clock.now()` (module-level,
monkeypatchable; tests freeze it).

## 3. Phase 2 — `docx.package` kernel (§7)

```python
def xml_equivalent(a: bytes, b: bytes) -> bool
```
Structural equality: element tags/attributes in Clark notation (namespace-
prefix-insensitive), attribute dicts equal, `text`/`tail` equal **verbatim**
(whitespace is content, §3), children compared recursively in document order.
Malformed XML raises `lxml.etree.XMLSyntaxError` — callers decide.

```python
def diff_package(path_a, path_b) -> PackageDiff
```
Part-by-part: `added`/`removed` names; common parts byte-compared, then XML
parts semantically compared (`xml_equivalent`; `.rels` and
`[Content_Types].xml` as maps per above; malformed XML → conservatively
changed), binary parts by size+SHA-256.

```python
@dataclass(frozen=True)
class PartDiff:   part: str; kind: str            # "xml" | "binary"
                  before_sha256: str; after_sha256: str
                  semantic_change: bool
@dataclass(frozen=True)
class PackageDiff:
    added: tuple[str, ...]; removed: tuple[str, ...]
    changed: tuple[PartDiff, ...]                  # byte-changed parts
    byte_identical_count: int
    def semantic_changed_parts(self) -> tuple[str, ...]
    @property def is_semantically_empty(self) -> bool
    def to_dict(self) -> dict                      # schema="paper_package_diff", version=1
```

```python
def patch_save(original_path, document, out_path) -> PatchSaveResult
```
Compare-based (no opc internals, no dirty flags): serialize `document` the
normal way, then for every XML part semantically identical to `original_path`'s
part, restore the original bytes. If after restoration the part set and every
blob are byte-equal to the original package, `out_path` is written as a
**verbatim copy** of the original file (no-op → byte-identical, §7 test).
Deterministic zip (F2). All writes: temp file in `out_path`'s directory →
`os.replace` (a mid-write crash leaves any existing `out_path` intact —
failure-injection tested). `original_path == out_path` is allowed;
`document` saving over its own open source file is fine (bytes are read
up front). Returns:

```python
@dataclass(frozen=True)
class PatchSaveResult:
    restored_parts: tuple[str, ...]   # semantically identical, original bytes kept
    changed_parts: tuple[str, ...]    # genuinely different content
    added_parts: tuple[str, ...); removed_parts: tuple[str, ...]
    verbatim_copy: bool               # True on the no-op path
    def to_dict(self) -> dict         # schema="paper_patch_save", version=1
```

```python
doc = docx.Document("contract.docx")
doc.paragraphs[0].add_run(" (amended)")
result = docx.package.patch_save("contract.docx", doc, "contract-out.docx")
assert result.changed_parts == ("word/document.xml",)
```

Pinned tests: no-op byte identity; single-part edit → exactly that part;
trailing-space `w:t` pair compares NOT equivalent; failure injection;
zip determinism (same input → same bytes, twice).

## 4. Phase 3 — `docx.story` traversal & inspection (opt-in)

New part classes `FootnotesPart` / `EndnotesPart`
(`src/docx/parts/footnotes.py`, `endnotes.py`, both `StoryPart` subclasses),
registered for `CT.WML_FOOTNOTES` / `CT.WML_ENDNOTES` in
`src/docx/__init__.py` (constants already exist). v0 reads these parts; it
does not create them, so no `default()` factory ships (amended from the
first draft, which mentioned templates).

```python
def story_parts(document) -> tuple[str, ...]
    # every story part present, sorted: word/document.xml, word/header*.xml,
    # word/footer*.xml, word/footnotes.xml, word/endnotes.xml, word/comments.xml

def iter_blocks(document, *, view: str = "current") -> Iterator[Block]
def outline(document, *, view: str = "current") -> Outline
```

`view` (pinned vocabulary, from the reference's extraction modes):
`"current"` = tracked insertions included, deleted text excluded (the
document as it stands if all changes were accepted); `"original"` = deletions
included, insertions excluded; `"all"` = everything.

```python
@dataclass(frozen=True)
class Block:
    story: str; kind: str                  # "paragraph" | "table"
    index: int                             # 0-based within story
    anchor: Anchor
    text: str                              # visible text under `view`
    style_id: str | None
    in_insert: bool; in_delete: bool
    in_content_control: bool; in_text_box: bool
    table: TableShape | None               # kind=="table" only: rows, columns,
                                           # has_merges, has_nested_table
class Outline:
    blocks: tuple[Block, ...]; story_parts: tuple[str, ...]
    blind_region_counts: dict[str, int]    # tracked_insertions, tracked_deletions,
                                           # content_controls, text_boxes
    def to_dict(self) -> dict              # schema="paper_outline", version=1 (golden-tested)
```

Traversal rules (pinned):
- Paragraphs inside table cells are NOT emitted as separate blocks (their
  text belongs to the table block); paragraphs inside `w:sdt` content and
  `w:txbxContent` ARE emitted, flagged.
- `mc:AlternateContent`: only the first `mc:Choice` is traversed, `mc:Fallback`
  is skipped — one visible text box yields its text exactly once (the
  LibreOffice textbox fixture is the regression test).
- Empty-text paragraph blocks are emitted (index stability matters more than
  compactness); `Outline.to_dict()` is byte-deterministic run-to-run.

```python
from docx.story import outline
o = outline(docx.Document("contract.docx"))
[b.text for b in o.blocks if b.in_text_box]   # text nobody else can see
```

## 5. Phase 4 — `docx.search`: `find_text` and Span

```python
def normalize_text(value: str) -> str
```
Exact reference table: curly quotes→ASCII, en/em-dash/minus→hyphen,
NBSP/figure/narrow/thin spaces→space, soft hyphen→deleted, tab→space,
CR→NL, whitespace runs collapsed to one space, casefolded. Each rule gets a
fixture-pair test.

```python
def find_text(document, needle: str, *,
              nth: int | None = None, near: str | None = None,
              story: str | None = None, view: str = "current") -> list[Span]
def find_one(document, needle: str, *, nth=None, near=None, story=None,
             view="current") -> Span
    # 0 matches -> TargetNotFoundError; >1 after disambiguators -> AmbiguousTargetError
```
Matching runs over Phase 3 traversal text per story part: assembles across
fragmented runs and across paragraph boundaries (paragraph breaks normalize
to a single space in the needle). `near` ranks by distance from the match to
the nearest occurrence of `near`'s normalized text in the same story
(improvement over the reference's in-match heuristic); `nth` is 1-based
among the (possibly `near`-ranked) matches.

```python
class Span:
    text: str; story: str; anchor: Anchor
    in_insert: bool; in_delete: bool; in_content_control: bool; in_text_box: bool
    crosses_paragraphs: bool
    def replace(self, new_text: str, *, tracked: bool = False,
                author: str | None = None,
                date: datetime | None = None) -> ReplaceResult
```
A Span holds live references into the document tree plus its captured text;
`replace` revalidates against the tree first and raises `TargetNotFoundError`
if the underlying text has changed (stale span). Span mapping is tested **by
use** (perform the replace, assert the outcome), not by inspecting offsets.

```python
span = find_one(doc, "$75–100/hr")          # smart quotes/dashes normalized
span = find_one(doc, "payment terms", near="Net 30")
```

## 6. Phases 5–6 — `Span.replace`

**Untracked (`tracked=False`).** In-place atom editing (reference semantics,
kept): the start `w:t` keeps its prefix + the replacement, interior `w:t`
atoms are emptied, the end atom keeps its tail. Every run keeps its `rPr`
byte-identical; replacement text renders with the start run's formatting;
`xml:space="preserve"` is managed on touched `w:t` nodes only.

**Tracked (`tracked=True`).** `author` is required (`ValueError` otherwise —
programmer error). Common prefix/suffix trim first: the redline marks
`75-10 → 85-11`, not the sentence. Emits `w:del` (containing `w:delText`,
never live `w:t`) and `w:ins` after the start run, both stamped
author/date/unique `w:id` (allocated above the story's current max), `rPr`
cloned from the start run to both sides. Same-paragraph targets only.

Refusal conditions (both modes, all atomic):
- Span includes deleted (`w:delText`) or field-instruction (`w:instrText`)
  text → `UnsupportedStructureError`.
- Span crosses a tab or line break (`w:tab`/`w:br`/`w:cr`) →
  `UnsupportedStructureError` (matching sees them as spaces; replacing
  across one would silently drop it).
- Span starts and ends in different content-control (`w:sdt`) scopes →
  `BoundaryViolationError`.
- Span crosses a paragraph boundary → `BoundaryViolationError` (block ops
  are Phase 7's job).
- Stale, consumed, or detached span → `TargetNotFoundError` (detached = the
  span's containing structure was removed from the tree after capture).
- Tracked replacement where trim leaves no change → `TargetNotFoundError`
  ("nothing to change" is a targeting failure, not silence).
- Tracked only: span inside a hyperlink, or straddling a pending `w:ins`
  boundary (layered revision history cannot be represented faithfully) →
  `UnsupportedStructureError`; a span wholly inside one `w:ins` nests fine.
- Tracked only: kept prefix/suffix characters are clamped to their own runs
  so unchanged characters never migrate formatting; the redline may re-state
  run-crossing kept characters inside the marked span instead.

```python
find_one(doc, "$75–100/hr").replace("$85–110/hr")                       # surgical
find_one(doc, "forty-two units").replace(
    "forty-seven units", tracked=True, author="Alice Editor")           # redline
```
Invariant (Phase 5): `replace(x→y)` then `replace(y→x)` restores visible text
always, and restores formatting fully when the span's own text carries
uniform formatting. **Amendment (found in implementation):** a span covering
a formatting transition collapses its interior formatting into the start
run's when replaced — the information is destroyed by any replacement (the
reference helpers share this property) — so the inverse restores text plus
all *out-of-span* formatting, and the in-span text renders uniformly.
Invariant (Phase 6, closed in Phase 8): accept(tracked replace) ≡ plain
replace; reject ≡ original.

`ReplaceResult`: `(story, deleted_text, inserted_text, tracked, revision_ids)`
with `.to_dict()` (`schema="paper_replace", version=1`).

**New oxml vocabulary (Phase 6, comments-feature commit shape):**
`src/docx/oxml/revision.py` defines `CT_RunTrackChange` for `w:ins`/`w:del`
(RequiredAttribute `w:id`/`w:author`, OptionalAttribute `w:date`;
`r = ZeroOrMore("w:r")`), registered in `src/docx/oxml/__init__.py`;
`w:delText` registers to the existing `CT_Text` shape. Emission goes through
these classes — never hand-assembled lxml in proxy code.

## 7. Phase 7 — `docx.blocks`

```python
AnchorLike = str | Block | Span | Anchor   # str -> find_one (must be unambiguous)

def insert_section_after(document, anchor: AnchorLike, *,
                         heading: str, paragraphs: Sequence[str],
                         heading_style: str = "Heading2",
                         body_style: str | None = None,
                         tracked: bool = False, author: str | None = None,
                         date: datetime | None = None) -> BlockEditResult

def tracked_delete_paragraphs(document, start_anchor: AnchorLike, *,
                              end_anchor: AnchorLike | None = None,
                              count: int = 1, author: str,
                              date: datetime | None = None) -> BlockEditResult

def tracked_replace_paragraphs(document, start_anchor: AnchorLike,
                               replacement_paragraphs: Sequence[str], *,
                               end_anchor: AnchorLike | None = None,
                               count: int = 1, body_style: str | None = None,
                               author: str, date: datetime | None = None) -> BlockEditResult
```
- Anchor resolution: `Anchor`/`Block`/`Span` are verified against the tree
  (`content_hash` must still match → else `TargetNotFoundError`); a `str` is
  found via `find_one` (`AmbiguousTargetError`/`TargetNotFoundError` apply).
- **Same-parent rule (load-bearing, kept):** every selected paragraph must
  share one parent element; selections spanning story regions, table
  boundaries, or content-control boundaries → `BoundaryViolationError`.
- Tracked paragraph delete preserves `pPr` (reference approach) and
  additionally marks the **paragraph mark** deleted (`w:pPr/w:rPr/w:del`) so
  accepting removes the paragraph instead of leaving an empty one — a
  documented improvement over the best-effort reference.
- `style` ids are validated against the styles part → `TargetNotFoundError`
  for undefined styles (no silent fake styling).
- Changed-part budget: body-only edits touch exactly `word/document.xml`.

`BlockEditResult`: `(story, inserted_blocks, deleted_blocks, deleted_text,
revision_ids)` + `.to_dict()` (`schema="paper_block_edit", version=1`).

## 8. Phase 8 — `Document.revisions`

New module `src/docx/revision.py` (proxy layer); additive property
`Document.revisions` on the existing proxy.

```python
class Revision:
    revision_type: str            # "insertion" | "deletion"
    author: str; date: datetime | None
    text: str; story: str; anchor: Anchor
    def accept(self) -> None
    def reject(self) -> None

class Revisions(Sequence[Revision]):
    def accept_all(self, *, author: str | None = None) -> int   # count resolved
    def reject_all(self, *, author: str | None = None) -> int
    def to_dict(self) -> dict     # schema="paper_revisions", version=1
```
Semantics: accept insertion = unwrap children in place; reject insertion =
remove; accept deletion = remove content (and the paragraph itself when its
paragraph mark is revision-deleted and nothing visible remains); reject
deletion = `w:delText`→`w:t`, unwrap runs in place. Enumeration covers every
story part. In-memory operations (caller saves; `patch_save` recommended).
Filtered resolution that matches nothing returns 0 — not a refusal.

**Tracked-edit algebra invariants (§4, the highest-value tests):** on the
tracked fixtures — `accept(tracked_replace(x→y))` yields visible text equal
to plain `replace(x→y)`; `reject(...)` yields the original; same for block
operations.

## 9. Phase 9 — `docx.tableops` and `docx.numbering`

```python
def find_table(document, *, near_text: str) -> Table          # TargetNotFoundError
def update_cell(table: Table, row: int, column: int, new_text: str, *,
                tracked: bool = False, author: str | None = None,
                date: datetime | None = None) -> ReplaceResult
def insert_row_after(table: Table, row: int, values: Sequence[str], *,
                     copy_format_from: int | None = None) -> None
def delete_row(table: Table, row: int) -> None
```
`row`/`column` are 0-based. Any structural op or cell update on a table with
`vMerge`, `gridSpan`, or nested tables → `UnsupportedStructureError`
(loud refusal, no cleverness — pinned). `update_cell` routes through Span
replace over the cell's visible text (first-run formatting preserved);
empty cells get a plain styled run.

```python
def list_numbering(document) -> NumberingReport               # definitions + numbered paragraphs
def apply_numbering(paragraph, *, num_id: int, level: int = 0) -> None
    # num_id must exist in word/numbering.xml -> TargetNotFoundError
def apply_list_style(paragraph, style_name: str) -> None
    # style must exist -> TargetNotFoundError
```
Authoring new `numbering.xml` definitions, and restart/continue mechanics
that require new definitions, are explicitly out of v0 →
`UnsupportedStructureError` with a message that says so.

`NumberingReport`: `(definitions=(num_id, abstract_num_id, level_formats),
numbered_paragraphs=(story, index, num_id, level, text))` + `.to_dict()`
(`schema="paper_numbering", version=1`).

---

## 10. v0.11 — revision completion, scrub/finalize, protection (PLAN-v0.11)

**Semantic amendments (refusal→capability conversions, ledgered in
PAPER.md):** `Revisions` resolution now covers `format_change`
(`w:rPrChange`/`w:pPrChange` incl. paragraph marks), `row_insertion`/
`row_deletion` (`w:trPr` markers — new type names; they previously
misclassified as insertion/deletion), and moves as PAIRED UNITS (either
site resolves both). `Revisions.to_dict()` schema is v3. The exotic
remainder refuses BY NAME: `table_property_change`, `cell_revision`,
`section_property_change`, `numbering_change`, `custom_xml_revision`.

```python
Document.finalize(*, revisions: str = "accept") -> int
    # total resolution or typed refusal naming what blocked it; after a
    # successful return a rescan finds ZERO revision markup of any kind
Document.scrub(*, comments: bool = True, metadata: bool = True,
               track_changes_setting: bool = True, rsids: bool = False,
               hidden_text: bool = False) -> ScrubReport
    # ScrubReport (docx.scrubbing): itemized removals, .to_dict(),
    # schema="paper_scrub_report" v1; report-matches-diff is the contract.
    # metadata=True refuses while revisions are pending (attribution would
    # survive); document protection is REPORTED, never removed.
```

```python
# docx.protection — protection-aware mutations (Phase 3)
def protection_status(document) -> ProtectionStatus   # edit/enforced/acknowledged
def acknowledge_protection(document) -> ProtectionStatus
    # THE one override affordance (document-level, in-memory, never saved)
```

`w:documentProtection` with `w:enforcement` truthy makes every paper-docx
mutating API refuse with `DocumentProtectedError` (new member of the pinned
refusal hierarchy — flagged for human sign-off) naming the mode and the
override path. Uniform across modes in v0.11 (readOnly/forms/comments/
trackedChanges all refuse — conservative, documented). Upstream APIs are
untouched (strict superset). There is NO protection-stripping verb.

```python
# docx.package (kernel re-export; impl in docx/_compare.py) — Phase 4
def compare(original, revised, *, author: str, date: datetime | None = None,
            granularity: str = "word",           # "word" | "block"
            materialize: str | None = None,      # None | "accept" | "reject"
            ) -> CompareResult                   # .document/.findings/.to_dict()
```
The algebra is pinned by tests: `accept_all(compare(A,B))` == B and
`reject_all` == A in visible text across every story; `compare(A,A)` emits
zero revisions; identical inputs (fixed `date`) → byte-identical output.
Inputs with pending revisions refuse unless `materialize` resolves working
copies (files untouched). Declared limits: ins/del only (no move/rPrChange
synthesis); formatting-only, image/object, hyperlink, content-control,
section-break and comment differences are REPORT-ONLY findings; block
add/remove outside the main body, changed merged-cell rows, sdt/fldSimple
paragraphs in deletions, story-part set mismatch, and stories over the
documented block budget refuse (typed).

## 11. v0.11 Phase 5 — cross-document composition (PR-gated design)

This section IS the design gate the plan requires before implementation —
review it as the proposal. New module `docx.composition`:

```python
def insert_blocks_from(
    document, source,                 # destination Document, source Document
    start_anchor, *,                  # AnchorLike into the SOURCE body
    anchor,                           # AnchorLike into the DESTINATION body
    end_anchor=None, count=1,         # range selection, blocks.py semantics
    styles: str = "match_by_name",    # | "import_renamed"
) -> CompositionReport

def append_document(
    document, source, *,
    section: str = "new_page",        # | "continuous"
    styles: str = "match_by_name",
) -> CompositionReport
```

**Style reconciliation (the hard core).** Styles referenced by the copied
range (pStyle/rStyle/tblStyle + transitive basedOn/link/next chains):
`match_by_name` — a destination style with the same NAME wins (ids remapped;
content adopts the house look); missing names import their definitions.
`import_renamed` — same name + equivalent definition reuses; colliding-but-
different definitions clone under fresh ids and a deterministic
" (imported)" name suffix (content keeps its source look). The report
carries the full style map.

**Numbering.** Every `numId` referenced in the range gets a FRESH
destination id with deep-copied num + abstractNum definitions — copied
lists always RESTART (the pinned continue-vs-restart choice; continuing a
destination list is a future mode).

**Relationships/media.** Images copied as new parts with fresh rIds
(`r:embed` remapped); external hyperlinks recreated; embedded OLE objects →
typed refusal (declared). **Bookmarks** rename on collision (deterministic
suffix), ids reallocated, and REF/PAGEREF instructions INSIDE the range
remap to the new names; references leaving the range are report-only
findings.

**Source revisions/comments in the range → typed refusal** directing the
caller to `finalize()`/`scrub()` the source first (carry-through is not
cheap and is deferred). **Append semantics:** v0.11 keeps destination
headers/footers; `section="new_page"` prefixes the appended content with a
page break, `"continuous"` appends flush — no new `w:sectPr` is authored
(keep-source-headers is a declared future mode).

**Report-matches-diff.** `CompositionReport` declares every part the
operation may touch (document.xml, styles.xml, numbering.xml, rels,
[Content_Types].xml, media/*) plus the maps above and findings;
`.to_dict()` goldenable. The harness assertion is report-matches-diff,
not small-diff. Acceptance is job-shaped: the proposal-assembly eval in
tests/paper (base + clause docs + CV section with tables/images/lists/
headings → reopens clean, LO smoke, lists number, bounded style count).

## 12. v0.11 Phases 6-7 — bookmarks, field authoring, format resolver

```python
# docx.bookmarks — Phase 6
def list_bookmarks(document) -> list[BookmarkInfo]     # name/id/story/text
def create_bookmark(document, span, name) -> BookmarkInfo
    # wraps EXACTLY the span (edge runs split); Word-legal unique names;
    # globally unique ids
def delete_bookmark(document, name) -> None
    # markers only, text stays; refuses while a REF/PAGEREF references it

# docx.fields — author-and-delegate: placeholder results + updateFields
# flag, NEVER computed values (pagination is a renderer's job)
def add_page_number_field(paragraph) -> None           # PAGE
def add_page_count_field(paragraph) -> None            # NUMPAGES
def add_date_field(paragraph, *, date_format=None) -> None
def add_reference_field(paragraph, *, bookmark, kind="text") -> None
    # kind: "text" (REF) | "page" (PAGEREF) | "number" (REF \r)
def insert_toc_after(document, anchor, *, levels=(1, 3)) -> None
    # complex begin/separate/end form, dirty-flagged
```
Self-consistency is pinned: the v0.1 `in_field` guard refuses spans landing
in fields THIS module authors, same as Word's.

```python
# docx.formatting — Phase 7, read-only + provenance-bearing
def format_of(target) -> EffectiveFormat        # Run | Paragraph | Span
def surrounding_format(document, anchor) -> EffectiveFormat
```
Resolution: docDefaults → paragraph-style chain (basedOn, root→leaf) →
character-style chain → direct, with §17.7.3 toggle XOR (direct is
absolute; style layers XOR their TRUE occurrences — nested bold cancels,
pinned by test). Every value is a `ResolvedValue(value, source, chain)`;
span disagreement reports "mixed"; the declared-unresolvable list
(table-style conditional formatting, numbering-mark properties, EA/CS
variants) rides every result. `surrounding_format` is the Phase 5/6
match-the-neighbors helper.

## Module map (all new files unless marked additive)

| Path | Contents |
|---|---|
| `src/docx/errors.py` | refusal hierarchy |
| `src/docx/_clock.py` | injectable clock |
| `src/docx/_paperpkg.py` | kernel implementation |
| `src/docx/package.py` | **additive** re-export block (F1) |
| `src/docx/story.py` | Block/Anchor/Outline, traversal |
| `src/docx/parts/footnotes.py`, `parts/endnotes.py` | new story parts |
| `src/docx/__init__.py` | **additive** part registrations |
| `src/docx/search.py` | normalize_text, find_text/find_one, Span |
| `src/docx/oxml/revision.py` | CT_RunTrackChange (+ registrations, additive in `oxml/__init__.py`) |
| `src/docx/blocks.py` | block operations |
| `src/docx/revision.py` | Revision/Revisions proxies; `Document.revisions` **additive** |
| `src/docx/tableops.py` | guarded table ops |
| `src/docx/numbering.py` | numbering enumeration/application |

Reference files mined per organ are named in the plan's mining map; the
ordered-insertion tables of `ooxml_util.py` are NOT ported (descriptors'
`successors=` is the in-library mechanism, §3).
