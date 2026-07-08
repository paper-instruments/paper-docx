# paper-docx

**A drop-in fork of [`python-docx`](https://github.com/python-openxml/python-docx)
that turns a document *generator* into a document *collaborator*: it sees
everything Word displays, edits real-world text without breaking formatting,
proposes and resolves genuine tracked changes, compares and composes whole
documents, and finalizes files for delivery — all behind typed, atomic
refusals instead of silent corruption.**

```python
import docx                     # the import name never changes
doc = docx.Document("contract.docx")
```

Everything `python-docx` did, this package still does, unchanged — upstream's
own pytest and behave suites run green on every commit. What follows is what
it does *now* that stock `python-docx` cannot.

## The shortest demonstration

Two versions of a document in, a native Word redline out — with a guarantee.
This is self-contained; paste and run it:

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

# generate real w:ins/w:del markup that transforms v1 into v2
result = compare(f"{tmp}/v1.docx", f"{tmp}/v2.docx", author="Reviewer")
[(r.revision_type, r.text) for r in result.document.revisions]
# [('deletion', 'calendar'), ('insertion', 'business')]

# the algebra is guaranteed and tested: accepting the redline yields exactly v2
result.document.revisions.accept_all()
result.document.paragraphs[0].text
# 'Payment is due within thirty business days of the invoice date.'
```

`compare` emits markup Word renders as tracked changes; `accept_all(compare(a, b))`
reproduces `b` and `reject_all` reproduces `a`, across every part of the
document. That round-trip guarantee is the whole idea, applied everywhere.

## Why this fork exists

`python-docx` has an excellent core: a lossless package layer that
round-trips content it doesn't understand, and a disciplined XML layer with a
decade of absorbed edge cases. That's why we forked it rather than rebuilt.

But its *editing* surface stalled years ago, and the gaps are exactly where
professional, agent-driven document work lives:

- **It can't see whole regions of the document.** Text inside tracked
  insertions, content controls, text boxes, footnotes, and endnotes is
  invisible to `.paragraphs`/`.runs` — the library literally cannot read
  parts of what Word displays.
- **It can't edit real-world text.** Word fragments most text across
  multiple runs (`$75`&#8203;`–`&#8203;`100/hr` is typically three runs), so there is no
  way to find or replace it without destroying formatting.
- **It has no tracked changes** — no way to propose a redline, resolve one,
  or generate one by comparing two documents.
- **It can't finalize or compose.** No way to accept/scrub a file for
  delivery, and no way to assemble one document from others without
  corrupting styles and numbering.
- **Its saves churn the whole package**, so reviewing "what did this edit
  actually change?" is impossible at the file level.

Production systems have historically papered over these gaps with raw-XML
surgery, whose dominant failure mode is **silent corruption**: files that
open fine and are quietly wrong. This fork replaces that surgery with
first-class APIs whose failure mode is a loud, typed refusal.

## What it does

Grouped into three tiers. One tight bullet per organ — full signatures,
return types, and refusal conditions live in
[`API-PROPOSAL.md`](API-PROPOSAL.md) and [`PAPER.md`](PAPER.md).

### Perceive and edit one document

- **`docx.story`** — visibility-complete traversal of every story part (body,
  headers, footers, footnotes, endnotes, comments) and every region standard
  traversal is blind to, under a *view* (`"current"` / `"original"` / `"all"`).
  `outline(doc)` confesses what it couldn't read.
- **`docx.search`** — normalized, run-fragmentation-tolerant find. `find_one`
  returns a `Span` you can `replace` surgically (untouched runs keep formatting
  byte-for-byte), `replace(..., tracked=True)` as a minimal redline, or
  `comment` on. `replace_all` does it in one pass.
- **`docx.blocks`** — insert, delete, or replace whole paragraphs relative to
  a content anchor, plain or as a tracked redline that stamps paragraph marks
  so Word accepts/rejects them exactly.
- **`docx.tableops` / `docx.numbering`** — guarded cell/row edits and list
  numbering (apply existing definitions or author real bullet/decimal ones),
  refusing loudly on merged cells, nested tables, or undefined numbering.
- **`docx.controls`** — fill content controls type-correctly, clearing
  placeholder state so Word treats them as genuinely filled.
- **`docx.bookmarks` / `docx.fields`** — create bookmarks on a span; author
  page numbers, dates, cross-references, and a TOC — as *formulas* with
  placeholder results, never computed values.
- **`docx.formatting`** — read-only: what formatting does this text *actually*
  carry, resolved through defaults → styles → direct with correct toggle
  semantics, each value naming the layer it came from.

### Review, resolve, and finalize

- **`doc.revisions`** — enumerate every tracked change across every story
  part and resolve it: insertions, deletions, run/paragraph format changes,
  table-row revisions, and moves (as paired units). Exotic markup is
  enumerated and refused *by name*, never silently passed.
- **`doc.finalize()` / `doc.scrub()`** — resolve all revisions (or refuse,
  naming what blocked it), then strip reviewing residue (comments, metadata,
  RSIDs), returning a `ScrubReport` that itemizes exactly what left the file.
- **`docx.protection`** — every mutating API refuses with
  `DocumentProtectedError` on a Restrict-Editing setting rather than silently
  editing a locked template; one explicit override, and the setting is never
  stripped.

### Work across documents

- **`docx.package.compare`** — generate a native tracked-change redline from
  two documents, with the tested accept/reject algebra shown above.
- **`docx.package.patch_save` / `diff_package` / `text_diff`** — a
  compare-based narrow save (unchanged parts keep their original bytes, so a
  file-level diff shows your edit and nothing else) plus the diffs that prove
  it. `diagnose` triages an unopenable file into a typed verdict.
- **`docx.composition`** — copy formatted content between documents without
  corruption, reconciling styles, numbering, media, hyperlinks, and
  bookmarks; the returned `CompositionReport` declares every part it touched.
- **`docx.errors`** — every refusal is a typed `PaperRefusal` subclass, so a
  caller tells a safe refusal apart from a bug. A refused edit is a success
  mode; a quietly wrong file is the failure this package eliminates.

## The honest ceiling

What it deliberately does *not* do — each a trust signal, not an apology:

- **It never computes field values or paginates.** It authors formulas and
  sets the update-on-open flag; a renderer (Word, or headless LibreOffice)
  computes them.
- **`compare` emits insertions and deletions only** — no move synthesis, no
  cross-story move detection. Formatting-only, image, and object differences
  are *reported*, not redlined.
- **No OLE authoring**, and composition refuses embedded objects (typed).
- **No protection stripping, and no decryption** of password-protected files
  — the encrypted container gets a typed refusal.
- **No document-QA / `check()` API.** Judging arbitrary documents is harness
  territory and stays out of the package; load- and edit-time failures on bad
  input speak as typed, specific refusals, never raw tracebacks.
- **The format resolver declares what it can't resolve** (table-style
  conditional formatting, theme fonts, numbering-mark properties) rather than
  guess.

## How it stays trustworthy

- **Strict superset.** No existing behavior changes; new capability is new,
  explicitly named API. Upstream's full pytest (1609) and behave (650
  scenarios) suites gate every commit. The one sanctioned exception —
  protection-aware refusals on the fork's *own* APIs — is documented in
  `PAPER.md`; upstream APIs are untouched.
- **Refusal atomicity.** Validate fully, then mutate. Every documented
  refusal is tested to leave the document — in memory and on disk — byte-for-byte
  as it was.
- **Whitespace is content.** No comparison or rewrite path normalizes
  meaningful text whitespace — a trailing space in `w:t` is a real character.
- **Provenance-tested.** A frozen, hash-manifested fixture corpus
  (`tests/paper/fixtures/`) spanning generated and LibreOffice-authored files,
  with hand-verified ground truth and a LibreOffice headless load oracle;
  desktop-Word fixtures are tracked in `FIXTURE-REQUESTS.md`.
- **Adversarially reviewed.** Each release wave is attacked by a multi-agent
  sweep whose only bar is silent corruption / false state / broken invariant;
  every confirmed finding is fixed with a regression test before ship.

## Naming

Four names to keep distinct — the mismatch is intentional (same pattern as
Pillow/PIL):

- GitHub repository / PyPI distribution: **`paper-docx`**
- Python import package: **`docx`** — frozen forever; millions of existing
  snippets and model priors say `from docx import Document`, and drop-in
  compatibility is the entire thesis of this fork
- Fork sentinel: `docx.__paper_version__`

Never rename `src/docx`, and never write `import paper_docx`.

## Installation

This repository is private and publication to PyPI is intentionally gated.
For now, install from Git:

```bash
pip install "paper-docx @ git+https://github.com/The-LLM-Data-Company/paper-docx.git@main"
```

Verify the fork sentinel:

```bash
python -c "import docx; print(docx.__paper_version__)"
```

## Repository map

```
src/docx/            the package (import name `docx`), upstream + fork organs:
  errors.py            PaperRefusal hierarchy (typed, atomic refusals)
  story.py             visibility-complete traversal, views, anchors
  search.py            normalized find + Span (surgical / tracked replace, comment)
  blocks.py            clause-level insert/delete/replace, tracked
  revision.py          Document.revisions — enumerate + resolve
  tableops.py          guarded table cell/row ops
  numbering.py         list reporting, application, and authoring
  controls.py          content-control enumeration and filling
  commentops.py        comment threads (reply, resolve, anchored text)
  bookmarks.py         bookmark enumerate / create / delete
  fields.py            field authoring (page, date, cross-ref, TOC)
  formatting.py        effective-format resolver (provenance-bearing)
  protection.py        Restrict-Editing awareness + override
  scrubbing.py         finalize / scrub (Document.finalize / .scrub)
  composition.py       cross-document composition
  package.py           kernel: xml_equivalent, diff_package, patch_save,
                        diagnose, text_diff, compare (impl in _paperpkg/_compare)
docs/                Sphinx docs — see docs/user/paper-additions.rst
tests/paper/         contract harness + frozen fixture corpus
```

## Going deeper

- [`docs/`](docs/) — Sphinx reference; start at `docs/user/paper-additions.rst`
  and the `api/paper-*.rst` pages (everything inherited works as documented at
  the [python-docx docs](https://python-docx.readthedocs.io/)).
- [`API-PROPOSAL.md`](API-PROPOSAL.md) — the full approved surface across all
  waves: signatures, return types, refusal conditions (mechanically enforced by
  `tests/paper/test_api_surface.py`).
- [`ARCHITECTURE-NOTES.md`](ARCHITECTURE-NOTES.md) — how the codebase is
  layered and where each organ lives.
- [`PAPER.md`](PAPER.md) — fork lineage, baseline test results, every
  refusal→capability conversion, and upstream merge policy.

`python-docx` is by Steve Canny and contributors (MIT); this fork exists
because that foundation was worth building on.
