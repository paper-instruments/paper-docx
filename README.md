# paper-docx

[![PyPI](https://img.shields.io/pypi/v/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![License](https://img.shields.io/pypi/l/paper-docx.svg)](LICENSE)
[![CI](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml)
[![Downloads](https://img.shields.io/pypi/dm/paper-docx.svg)](https://pypi.org/project/paper-docx/)

paper-docx is an agent-first Python library for safely inspecting, editing,
reviewing, and composing existing Word documents. It is a drop-in hard fork of
[python-docx](https://github.com/python-openxml/python-docx): the distribution
is renamed, the `docx` import is frozen, and every fork addition either does
exactly what it claims or raises a typed refusal and leaves the document
byte-for-byte unchanged.

```python
import docx                       # the import name is unchanged
doc = docx.Document("contract.docx")
```

## Table of contents

- [Why it exists](#why-it-exists)
- [Quick start](#quick-start)
- [Feature tour](#feature-tour)
- [What we changed, and why](#what-we-changed-and-why)
- [paper-docx vs python-docx at a glance](#paper-docx-vs-python-docx-at-a-glance)
- [Drop-in compatibility](#drop-in-compatibility)
- [Documentation](#documentation)
- [How it's tested](#how-its-tested)
- [Roadmap and known limitations](#roadmap-and-known-limitations)
- [Contributing](#contributing)
- [Community](#community)
- [Security](#security)
- [Citation](#citation)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## Why it exists

python-docx is excellent at creating documents. Its lossless package layer and
disciplined XML mapping are why this fork builds on it. But its editing
surface stalled years ago, and changing a real document — a contract, a
policy, a filing — is a different problem from writing a new one. Production
agent systems fall back to raw XML surgery for exactly the operations that
matter most, and the dominant failure mode of that surgery is silent
corruption: a file that opens fine and is quietly wrong. An agent cannot
eyeball the result, so the library has to see everything, refuse rather than
guess, and report every outcome as typed data.

The biggest upstream gaps, each covered in
[What we changed, and why](#what-we-changed-and-why):

- No tracked-changes support at all: revisions are invisible to traversal and
  cannot be authored, accepted, or rejected
  ([reviewing](#reviewing-and-finalizing-documents)).
- No way to find or replace text Word has fragmented across runs — which is
  most text in real documents
  ([finding and editing](#finding-and-editing-real-world-text)).
- Traversal is blind to headers, footers, footnotes, comments, content
  controls, and text boxes
  ([finding and editing](#finding-and-editing-real-world-text)).
- `save()` writes in place, so a mid-write failure truncates the original
  file; untrusted packages go straight to `zipfile`
  ([package and save integrity](#package-and-save-integrity)).
- Structured surfaces — content controls, comments, bookmarks, fields,
  numbering, merged tables — can only be edited as raw text, silently breaking
  their semantics ([structured surfaces](#structured-surfaces)).

## Quick start

Install from PyPI:

```bash
python -m pip uninstall -y python-docx paper-docx
python -m pip install paper-docx
```

The clean uninstall matters when migrating from python-docx: both
distributions own the same `docx` import package, and pip cannot safely
overlay two distributions that own the same files. Verify the install:

```bash
paper-docx-doctor
```

Then create a native Word redline from two document versions:

```python
import tempfile, docx
from docx.package import compare

tmp = tempfile.mkdtemp()
a = docx.Document()
a.add_paragraph("Payment is due within thirty calendar days of the invoice date.")
a.save(f"{tmp}/v1.docx")
b = docx.Document()
b.add_paragraph("Payment is due within thirty business days of the invoice date.")
b.save(f"{tmp}/v2.docx")

result = compare(f"{tmp}/v1.docx", f"{tmp}/v2.docx", author="Reviewer")
[(r.revision_type, r.text) for r in result.document.revisions]
# [('deletion', 'calendar'), ('insertion', 'business')]

result.document.revisions.accept_all()
result.document.paragraphs[0].text
# 'Payment is due within thirty business days of the invoice date.'
```

That redline is real `w:ins`/`w:del` markup Word renders as tracked changes.
Before returning, `compare` proves on private copies that accepting the result
reproduces the revised document and rejecting it reproduces the original. If a
difference cannot be represented safely, it raises a typed refusal instead of
returning a persuasive but lossy redline.

## Feature tour

Every example below was executed and verified against the current release;
output comments are observed output.

### Perception: `docx.story`

One deterministic traversal across body, headers, footers, footnotes,
endnotes, and comments, with anchored blocks and explicit revision views
(`current`, `original`, `all`). `outline()` also counts what it cannot safely
read — fields, math, embedded objects, hidden text — instead of pretending
full visibility.

```python
from docx.story import outline, iter_blocks

o = outline(doc)
[b.text for b in o.blocks if b.style_id and b.style_id.startswith("Heading")]
# ['Master Services Agreement', '1. Scope of Work']
o.blind_region_counts["text_boxes"]
# 0
```

### Search and replace: `docx.search`

`find_one` and `find_text` match normalized text (smart quotes, dashes,
non-breaking spaces, case) across run fragmentation and story boundaries. The
returned `Span` replaces text while preserving every untouched run, or emits
the replacement as a genuine tracked change.

```python
from docx.search import find_one

p = doc.add_paragraph()
p.add_run("Hourly rate: ")
p.add_run("$75-100/hr").bold = True
p.add_run(", invoiced monthly.")

find_one(doc, "$75-100/hr").replace("$85-110/hr")
[(r.text, r.bold) for r in p.runs]
# [('Hourly rate: ', None), ('$85-110/hr', True), (', invoiced monthly.', None)]
```

Tracked, with authorship:

```python
find_one(doc2, "within 30 days").replace("within 45 days", tracked=True, author="Reviewer")
[(r.revision_type, r.author, r.text) for r in doc2.revisions]
# [('deletion', 'Reviewer', '30'), ('insertion', 'Reviewer', '45')]
```

`replace_all` applies a whole batch from one stable scan and reports every
outcome — including per-match refusals — instead of silently skipping:

```python
from docx.search import replace_all

outcome = replace_all(doc, "Contractor", "Consultant")
outcome.replaced_count, len(outcome.refused)
# (2, 0)
```

### Block editing: `docx.blocks`

Insert, delete, or replace whole paragraphs relative to a text anchor, plainly
or as tracked changes. `insert_blocks_after` takes a typed grammar — rich
paragraphs, bullet and numbered lists backed by real numbering definitions,
rectangular tables — so agents never fake bullets with Unicode glyphs.

```python
from docx.blocks import insert_blocks_after, RichParagraph, TextRun, ListBlock

insert_blocks_after(doc, "Deliverables are set out below.", blocks=[
    RichParagraph([TextRun("Phase one", bold=True), TextRun(" includes:")]),
    ListBlock(["Data pipeline audit", "Migration plan", "Runbook handoff"], kind="bullet"),
])
```

### Tables and numbering: `docx.tableops`, `docx.numbering`

Cells are addressed by the visual layout grid, not raw XML indexes, and
operations refuse on structures they cannot edit safely — merged cells,
nested tables, revision-bearing rows.

```python
from docx.tableops import find_table, update_cell, insert_row_after

table = find_table(doc, near_text="Milestone")
update_cell(table, row=1, column=1, new_text="$6,500")
insert_row_after(table, row=1, values=["Implementation", "$12,000"])
```

`docx.numbering` resolves effective numbering through style chains, creates
canonical bullet/decimal definitions idempotently, and restarts lists without
mutating shared definitions.

### Content controls: `docx.controls`

Controls are typed values, not decorative text. Setting a value respects the
control type — plain and rich text, checkbox glyphs, dropdown choices, date
metadata — clears placeholder state, and refuses locked or data-bound
controls.

```python
from docx.controls import list_controls, set_control_value

[(c.tag, c.control_type, c.value) for c in list_controls(doc)]
# [('inline-field-1', 'rich_text', 'controlled text'), ...]
set_control_value(doc, "Acme Analytics LLC", tag="inline-field-1")
```

### Bookmarks and fields: `docx.bookmarks`, `docx.fields`

Bookmarks are created over an exact span with Word-legal unique names, and
deletion refuses while any field, hyperlink, TOC, or formula still references
the name. Fields (PAGE, NUMPAGES, DATE, REF, TOC) are authored as balanced
field code with placeholder results; paper-docx never fakes a computed value —
Word refreshes them.

```python
from docx.bookmarks import create_bookmark
from docx.fields import add_reference_field, insert_toc_after

bm = create_bookmark(doc, find_one(doc, "60 days notice"), "notice_period")
add_reference_field(doc.add_paragraph("See "), bookmark="notice_period")
insert_toc_after(doc, "Contents", levels=(1, 2))
```

### Effective formatting: `docx.formatting`

What a run actually looks like is resolved through document defaults, style
chains, and direct formatting — with provenance for every value and an
explicit `unresolved` list rather than fabricated certainty.

```python
from docx.formatting import format_of

fmt = format_of(find_one(doc, "of the essence"))
fmt.properties["bold"]
# ResolvedValue(value=True, source='direct', chain=('direct',))
fmt.properties["size_pt"]
# ResolvedValue(value=11.0, source='doc_defaults', chain=('doc_defaults',))
```

### Revisions, finalize, scrub: `Document.revisions`

Enumerate every tracked change — insertions, deletions, paired moves, format
changes, row revisions — then accept or reject individually, filtered by
author, or all at once. `finalize()` proves no revision markup remains;
`scrub()` removes reviewing residue and reports exactly what was removed.

```python
doc.finalize(revisions="accept")
report = doc.scrub()
report.metadata_fields_cleared
# ['core:author', 'core:comments', 'core:created', 'core:modified', 'core:revision']
```

### Protection: `docx.protection`

Restrict-Editing settings are respected, not silently bypassed. Mutating
operations refuse on an enforced-protection document until the caller
explicitly acknowledges it for that open package; the protection itself is
always preserved.

```python
from docx.protection import protection_status, acknowledge_protection

protection_status(doc).to_dict()
# {'edit': None, 'formatting': False, 'enforced': False, 'acknowledged': False}
```

### Package operations: `docx.package`

`patch_save` keeps semantically unchanged parts byte-identical — a no-op edit
round-trips to the same bytes. `diff_package` classifies exactly what changed.
`diagnose` explains why an unreadable file cannot be opened instead of raising
a raw ZIP error.

```python
from docx.package import patch_save, diagnose

saved = patch_save("original.docx", doc, "patched.docx")
saved.verbatim_copy, saved.changed_parts
# (True, ())

diagnose("not-a-docx.txt").to_dict()
# {'schema': 'paper_diagnosis', 'version': 1, 'path': 'not-a-docx.txt',
#  'readable': False, 'kind': 'not-a-zip',
#  'problems': ['not a ZIP archive, so not an OPC package']}
```

### Cross-document composition: `docx.composition`

Copy formatted content between documents while reconciling styles, numbering,
media, hyperlinks, bookmarks, and control IDs, and get a report of every part
touched.

```python
from docx.composition import append_document

report = append_document(main, exhibit, section="new_page")
report.inserted_blocks, report.declared_parts
# (2, ['word/document.xml', 'word/styles.xml', 'word/numbering.xml',
#      'word/_rels/document.xml.rels', '[Content_Types].xml'])
```

### Typed refusals: `docx.errors`

Every unsafe operation raises a `PaperRefusal` subclass — `PackageLimitError`,
`AmbiguousTargetError`, `TargetNotFoundError`, `UnsupportedStructureError`,
`BoundaryViolationError`, `RelationshipPolicyError`, or
`DocumentProtectedError` — distinct from programmer errors, and always leaves
the document unchanged.

```python
from docx.errors import AmbiguousTargetError

try:
    find_one(doc, "The fee is payable")
except AmbiguousTargetError as e:
    print(e)
# 2 matches for 'The fee is payable' (at word/document.xml#0, word/document.xml#1);
# disambiguate with nth=, near=, or story=
```

## What we changed, and why

paper-docx is one large change on top of python-docx 1.2.0: 221 files and
roughly 36,000 added lines, almost all of it new modules and tests. Every
claim below is backed by tests in this repository, cited inline.

### Distribution safety

**The problem.** A fork that keeps upstream's names is indistinguishable at
runtime, and pip will happily overlay two distributions that own the same
`docx` file tree — leaving an import that is partly one library, partly the
other.

**What we did.** The distribution is renamed `paper-docx`; the import stays
`docx`. `docx.__version__` remains `"1.2.0"` for compatibility and
`docx.__paper_version__` identifies the fork. `import docx` raises
`ImportError` outright if both distributions' metadata are installed, and the
`paper-docx-doctor` console script verifies the environment — metadata, wheel
`RECORD` hashes of every `docx` file (checked before importing them), import
path, and sentinel — before an agent runs a document job.
*Proof:* [`tests/paper/test_distribution_identity.py`](tests/paper/test_distribution_identity.py),
[`tests/paper/test_distribution_doctor.py`](tests/paper/test_distribution_doctor.py).

### Package and save integrity

**The problem.** Upstream hands untrusted ZIPs straight to `zipfile`, trusts
relationship and content-type declarations, writes saves in place (a mid-write
failure truncates the original file), and re-serializes every XML part on
every save so byte diffs are meaningless.

**What we did.**

- Every ZIP member is preflighted before any decompression or XML parse:
  duplicate/case/NFC name collisions, encryption, header inconsistencies, and
  resource limits (4,096 members, 100:1 expansion ratio, 512 MiB expanded
  total, among others). Opening an untrusted attachment becomes a bounded,
  typed operation.
  *Proof:* [`tests/paper/test_audit_zip_preflight.py`](tests/paper/test_audit_zip_preflight.py).
- Relationships and content types are validated at load: missing targets,
  duplicate IDs, multiple document roots, and role/type mismatches refuse at
  open instead of failing as a `KeyError` mid-edit.
  *Proof:* [`tests/paper/test_practical_opc_hardening.py`](tests/paper/test_practical_opc_hardening.py).
- `save()` is staged, validated, and atomic: serialize to a sibling temp file,
  reopen and validate it, fsync, then atomically replace — with deterministic
  ZIP output. A failed save leaves the previous file intact.
  *Proof:* [`tests/paper/test_practical_opc_hardening.py`](tests/paper/test_practical_opc_hardening.py),
  [`tests/paper/test_kernel.py`](tests/paper/test_kernel.py).
- `patch_save` restores the original bytes of semantically unchanged parts and
  returns the original verbatim for a no-op; `diff_package` and
  `xml_equivalent` separate real change from serialization noise; `diagnose`
  classifies unreadable inputs (`.doc`, encrypted, wrong Office family).
  *Proof:* [`tests/paper/test_kernel.py`](tests/paper/test_kernel.py),
  [`tests/paper/test_audit_package.py`](tests/paper/test_audit_package.py).
- Multi-part mutations snapshot XML, relationships, parts, and blobs, and roll
  back on late failure, so a raised refusal is safe to catch and re-plan
  around.
  *Proof:* [`tests/paper/test_audit_late_failure_atomicity.py`](tests/paper/test_audit_late_failure_atomicity.py).

### Finding and editing real-world text

**The problem.** Word fragments text across runs unpredictably, so
`paragraph.text` matching misses most real phrases, and assigning
`paragraph.text` destroys run formatting. There is no supported way to replace
a phrase Word has split across formatting boundaries, and traversal cannot see
headers, footers, notes, comments, content controls, text boxes, or either
side of a tracked change.

**What we did.** `docx.story` traverses every story part with anchored,
hashed blocks and explicit revision views, and counts its blind spots
honestly. `docx.search` matches normalized text across fragmentation;
`Span.replace` edits only the matched atoms, preserving untouched runs, or
emits a minimal genuine tracked change (`w:del`/`w:ins`) with authorship and
dates. Anchors carry content hashes, so a stale or foreign target refuses
instead of editing the wrong content. `replace_all` applies a batch from one
stable scan with per-item outcomes and whole-batch rollback on staleness.
Tabs, breaks, and field instructions are modeled explicitly; unknown visible
content becomes a barrier a match cannot silently cross.
*Proof:* [`tests/paper/test_story.py`](tests/paper/test_story.py),
[`tests/paper/test_search_replace.py`](tests/paper/test_search_replace.py),
[`tests/paper/test_audit_editing.py`](tests/paper/test_audit_editing.py).

### Reviewing and finalizing documents

**The problem.** Upstream has no tracked-changes support: revisions cannot be
enumerated, authored, accepted, or rejected. Naively stripping revision
wrappers corrupts property state, paragraph joins, paired moves, and table
rows. Documents shipped after review still leak comments, metadata, and hidden
text, and Restrict-Editing protection is silently bypassed by XML edits.

**What we did.** `Document.revisions` enumerates insertions, deletions, paired
moves, format changes, and row revisions; accept/reject works per revision,
filtered by author, or in full, resolving compound units transactionally and
refusing exotic markup by name rather than flattening it. `finalize()` proves
zero revision markup remains. `scrub()` removes comments, metadata, RSIDs, and
hidden text on request and reports exactly what it removed. Paper mutators
refuse enforced protection until the caller explicitly acknowledges it —
package-local, never persisted, and the protection itself is preserved.
*Proof:* [`tests/paper/test_resolution.py`](tests/paper/test_resolution.py),
[`tests/paper/test_audit_revision_compounds.py`](tests/paper/test_audit_revision_compounds.py),
[`tests/paper/test_scrub.py`](tests/paper/test_scrub.py).

### Structured surfaces

**The problem.** Content controls, comments, bookmarks, fields, numbering, and
merged tables all carry semantics that raw text editing violates: locks and
data bindings, thread and resolution state, field dependencies, shared
numbering definitions, visual-vs-XML grid divergence.

**What we did.** Typed operations per surface: control values set by type with
placeholder clearing and lock/binding refusals; comment threads with reply and
resolve; bookmarks with dependency-aware deletion (REF, PAGEREF, HYPERLINK,
TOC, and formula references detected even when split across field runs);
balanced field authoring without faked results; layout-grid table addressing
with refusals on merged and nested structures; real numbering graphs with
idempotent canonical definitions; effective-formatting resolution with
provenance.
*Proof:* [`tests/paper/test_tableops_numbering.py`](tests/paper/test_tableops_numbering.py),
[`tests/paper/test_fields_bookmarks.py`](tests/paper/test_fields_bookmarks.py),
[`tests/paper/test_audit_commentops.py`](tests/paper/test_audit_commentops.py),
[`tests/paper/test_formatting.py`](tests/paper/test_formatting.py).

### Cross-document workflows

**The problem.** There is no supported way to turn two document versions into
a tracked-changes redline, and copying content between documents leaves
relationship IDs pointing into the source package and collides styles,
numbering, bookmarks, and control IDs.

**What we did.** `compare()` generates a native redline and, before returning,
mechanically proves that accepting it reproduces the revised package and
rejecting it reproduces the original — across every package part. If a
difference cannot survive that proof, it refuses. `composition` copies body
content while reconciling style graphs, allocating fresh numbering, recreating
media and hyperlink relationships, renaming colliding bookmarks (rewriting
their references), and reporting every part it may have touched. `text_diff`
and `pending_changes` give per-story human-readable diffs.
*Proof:* [`tests/paper/test_compare.py`](tests/paper/test_compare.py),
[`tests/paper/test_audit_compare.py`](tests/paper/test_audit_compare.py),
[`tests/paper/test_composition.py`](tests/paper/test_composition.py).

### What did not change, and what is intentionally out of scope

Honesty about limits is part of the safety contract:

- `compare()` is a strict text/table-row redliner, not a general Word compare.
  Formatting-only, field, content-control, image, hyperlink, and merged-table
  differences raise a typed refusal rather than producing a lossy redline.
- Story traversal reports `blind_spots` — math, embedded objects, `altChunk`,
  hidden text are counted, not read.
- Revision resolution refuses exotic forms by name: table/cell/section
  property, numbering, and custom-XML revisions.
- Effective formatting names its unresolved categories (theme resolution,
  table conditional formatting, numbering-mark formatting).
- Protection enforcement covers Paper mutators; most inherited upstream
  mutators are intentionally untouched — it is a policy surface, not an
  access-control system.
- Composition is body-only; `section='new_page'` inserts a page break, not a
  new Word section.
- A write-only, non-seekable stream is validated before the first byte is
  written, but a failure during the final copy cannot be rolled back.
- The fork is *almost* purely additive. Deliberate exceptions, all
  safety-motivated and tested: inherited `save()` is now staged and atomic,
  package loading validates ZIP and relationship structure, and the inherited
  comment APIs gained protection and rollback checks. Code written against
  python-docx 1.2.0 keeps working; hostile or broken files that used to load
  ambiguously now refuse with a typed error.

## paper-docx vs python-docx at a glance

| Capability | python-docx | paper-docx |
|---|---|---|
| Create documents, paragraphs, tables, styles | ✅ | ✅ (inherited) |
| Find text across run fragmentation | — | `docx.search` |
| Replace text preserving untouched formatting | — | `Span.replace` |
| Author tracked changes (`w:ins`/`w:del`) | — | `Span.replace(tracked=True)`, `docx.blocks` |
| Enumerate / accept / reject revisions (incl. moves, rows) | — | `Document.revisions` |
| Traverse headers, footers, notes, comments, text boxes | partial, manual | `docx.story`, with blind-spot census |
| Fill content controls by type | — | `docx.controls` |
| Comment threads: reply, resolve | add only | `docx.commentops` |
| Table edits by visual grid, merge-aware | raw XML indexes | `docx.tableops` |
| Create real list numbering | — | `docx.numbering` |
| Dependency-aware bookmarks | — | `docx.bookmarks` |
| Author fields (PAGE, REF, TOC) | — | `docx.fields` |
| Effective formatting with provenance | direct properties only | `docx.formatting` |
| Finalize and scrub review residue | — | `finalize()`, `scrub()` |
| Respect Restrict-Editing protection | — | `docx.protection` |
| Untrusted-package preflight (zip bombs, bad OPC) | — | on every open |
| Atomic, deterministic save | — | on every save |
| Byte-preserving minimal save | — | `patch_save` |
| Tracked-changes redline from two versions | — | `compare()`, algebra-proven |
| Cross-document composition with reconciliation | — | `docx.composition` |
| Typed, catchable refusals; document never half-mutated | — | `docx.errors`, transactions |
| Install verification | — | `paper-docx-doctor` |

## Drop-in compatibility

Only the distribution and repository are renamed. The importable package stays
`docx` — the same distribution/import split as Pillow (`pip install pillow`,
`import PIL`). Existing code, snippets, and model priors keep working.

- GitHub repository / PyPI distribution: **`paper-docx`**
- Python import: **`docx`**
- Upstream compatibility version: `docx.__version__ == "1.2.0"`
- Fork sentinel: `docx.__paper_version__ == "0.1.2"`

Two caveats, both consequences of the shared import name:

- pip does not treat `paper-docx` as satisfying another package's declared
  dependency on `python-docx`. That dependency will reinstall upstream and
  overwrite shared `docx` files. Replace or remove the dependency, or isolate
  that package in its own environment. `import docx` detects the resulting
  dual install and raises `ImportError` rather than running a mixed overlay.
- In a controlled deployment, a pip constraint containing `python-docx<0`
  makes direct or transitive attempts to install upstream fail loudly. Apply
  it to every install in that environment.

## Documentation

- [`docs/user/paper-additions.rst`](docs/user/paper-additions.rst) — a guided
  tour of everything the fork adds.
- [`docs/api/paper-*.rst`](docs/api) — API reference for each fork module
  (`story`, `search`, `blocks`, `tableops`, `numbering`, `controls`,
  `bookmarks`, `fields`, `formatting`, `revisions`, `scrubbing`,
  `protection`, `package`, `composition`, `errors`).
- Everything inherited works as documented in the
  [python-docx documentation](https://python-docx.readthedocs.io/).

## How it's tested

- Upstream's full pytest and behave suites run on every commit, on Python
  3.9 through 3.13, to hold the drop-in promise.
- Fork behavior is covered by `tests/paper/`: the paper-specific suite
  passes 1,020 tests, the combined suite passes 2,597, and the behave
  acceptance suite passes 650 scenarios (verified on this release).
- Fixtures are a frozen, hash-pinned corpus (`MANIFEST.sha256`) of generated
  and LibreOffice-authored documents, including adversarial and
  corrupt-by-construction files.
- A contract harness asserts refusal atomicity (failed operations leave
  bytes unchanged), save/reopen validity, and changed-part budgets; a
  separate CI job loads fixtures through headless LibreOffice as a
  cross-producer smoke test.
- Signatures of the public fork API are pinned by
  [`tests/paper/test_api_surface.py`](tests/paper/test_api_surface.py).

## Roadmap and known limitations

Current limits are listed in
[What did not change](#what-did-not-change-and-what-is-intentionally-out-of-scope).
The near-term direction:

- Word- and Google Docs-authored fixture buckets (the corpus is currently
  generated + LibreOffice-authored).
- Reporting formatting-only differences in `compare()` as findings rather
  than refusals.
- Broader revision-resolution coverage for table, cell, and section-property
  revision forms as real-world fixtures are collected.
- Uniform JSON schema/version fields across the remaining small inspection
  records.

## Contributing

See [`CONTRIBUTING.md`](CONTRIBUTING.md) for setup (`uv sync`), the check
suite (pytest, behave, ruff, pyright), fixture rules, and what every change
must preserve — above all, refusal atomicity. Bug reports and PRs are welcome
on [GitHub issues](https://github.com/paper-instruments/paper-docx/issues).

## Community

Questions and ideas belong in
[GitHub Discussions](https://github.com/paper-instruments/paper-docx/discussions);
bug reports in
[GitHub issues](https://github.com/paper-instruments/paper-docx/issues).
If you are building agent systems on paper-docx, we would like to hear what
refuses that shouldn't, and what doesn't refuse that should.

## Security

paper-docx routinely opens untrusted files, and we treat parsing
vulnerabilities as high severity. Please report vulnerabilities privately —
see [`SECURITY.md`](SECURITY.md). Do not open public issues for security
reports.

## Citation

If you use paper-docx in research, cite both the fork and the upstream
project it builds on (see [`CITATION.cff`](CITATION.cff)):

```bibtex
@software{paper_docx,
  title   = {paper-docx: an agent-first fork of python-docx for safely editing Word documents},
  author  = {{Paper Instruments, Inc.} and Canny, Steve},
  year    = {2026},
  url     = {https://github.com/paper-instruments/paper-docx},
  version = {0.1.2},
  license = {MIT}
}
```

## License

MIT, inherited from python-docx. Original work © 2013 Steve Canny and the
python-docx contributors; fork additions © 2026 Paper Instruments, Inc. The
upstream license and attribution are preserved. See [`LICENSE`](LICENSE).

## Acknowledgments

paper-docx exists because python-docx is good. Steve Canny and the
python-docx contributors built a lossless package layer, a disciplined XML
mapping, and a decade of absorbed real-world edge cases — the foundation every
module in this fork stands on. This project forked to change how *editing*
works, not because anything upstream was careless; we hope the safety work
here is useful upstream too.
