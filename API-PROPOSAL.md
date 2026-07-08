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
- Span starts and ends in different content-control (`w:sdt`) scopes →
  `BoundaryViolationError`.
- Span crosses a paragraph boundary → `BoundaryViolationError` (block ops
  are Phase 7's job).
- Stale span → `TargetNotFoundError`.
- Tracked replacement where trim leaves no change → `TargetNotFoundError`
  ("nothing to change" is a targeting failure, not silence).

```python
find_one(doc, "$75–100/hr").replace("$85–110/hr")                       # surgical
find_one(doc, "forty-two units").replace(
    "forty-seven units", tracked=True, author="Alice Editor")           # redline
```
Invariant (Phase 5): `replace(x→y)` then `replace(y→x)` restores visible text
and formatting. Invariant (Phase 6, closed in Phase 8): accept(tracked
replace) ≡ plain replace; reject ≡ original.

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
