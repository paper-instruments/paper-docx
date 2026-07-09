# paper-docx

**paper-docx reads and edits Microsoft Word (`.docx`) documents from Python.**
It is a drop-in fork of [`python-docx`](https://github.com/python-openxml/python-docx)
built for programmatic and agent-driven document work: seeing everything the
document contains, editing real-world text without breaking formatting,
proposing and resolving tracked changes, comparing and composing whole
documents, and saving files that are safe to send.

```python
import docx                       # the import name is unchanged
doc = docx.Document("contract.docx")
```

Everything `python-docx` does, paper-docx still does, unchanged. Its upstream
test suites run green on every commit. The rest of this document describes what
paper-docx adds on top.

## Where it came from

`python-docx` has an excellent foundation: a lossless package layer that
round-trips content it doesn't understand, and a disciplined XML layer with
years of absorbed edge cases. It is, however, aimed at *creating* documents.
paper-docx keeps that foundation and adds the surface you need to *work with
documents that already exist* — the operations that dominate real editing and
review, and that a program or agent driving Word cannot otherwise perform
safely.

The failure mode it is designed to avoid is **silent corruption**: hand-rolled
XML edits that produce a file which opens without error and is quietly wrong.
Every operation paper-docx adds either does the right thing or refuses with a
typed, specific error — it never guesses and never half-finishes.

## A short example

Two versions of a document go in; a native Word redline comes out.

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

`compare` emits markup Word renders as tracked changes. Accepting the redline
reproduces the revised document and rejecting it reproduces the original,
across every part of the file — that round-trip guarantee holds throughout the
library.

## What it adds

### Reading and editing one document

- **`docx.story`** — traversal that covers every part of a document (body,
  headers, footers, footnotes, endnotes, comments) and the regions ordinary
  traversal misses (tracked insertions, content controls, text boxes), under a
  choice of view: the document as it stands, as it was before pending
  revisions, or everything at once.
- **`docx.search`** — find text the way a person writes it (normalized for
  smart quotes, dashes, and spacing) even when Word has split it across many
  runs. A match returns a `Span` you can replace surgically, replace as a
  tracked change, or attach a comment to.
- **`docx.blocks`** — insert, delete, or replace whole paragraphs relative to a
  text anchor, plainly or as a tracked change.
- **`docx.tableops` / `docx.numbering`** — cell and row edits and list
  numbering that refuse on structures they cannot handle safely (merged cells,
  nested tables, undefined numbering) rather than corrupt them.
- **`docx.controls`** — fill content controls with the correct value type and
  clear their placeholder state, so Word treats them as genuinely filled.
- **`docx.bookmarks` / `docx.fields`** — create bookmarks over a span, and
  author page numbers, dates, cross-references, and tables of contents as
  fields with placeholder results.
- **`docx.formatting`** — resolve what formatting a piece of text actually
  carries, following document defaults, styles, and direct formatting, with
  each value reporting the layer it came from.

### Reviewing and finalizing

- **`doc.revisions`** — enumerate every tracked change across every part of the
  document and resolve it: insertions, deletions, run and paragraph format
  changes, table-row revisions, and moves. Markup it cannot resolve is listed
  by name rather than silently passed over.
- **`doc.finalize()` / `doc.scrub()`** — accept or reject all revisions, then
  remove reviewing residue (comments, metadata, revision-save ids), returning a
  report of exactly what was removed.
- **`docx.protection`** — respect a document's Restrict-Editing setting:
  mutating operations refuse on a protected document unless the caller
  explicitly overrides. The setting itself is never stripped.

### Working across documents

- **`docx.package.compare`** — generate a native tracked-change redline from
  two documents, with the accept/reject round-trip shown above.
- **`docx.package.patch_save` / `diff_package` / `text_diff`** — a save that
  leaves parts you didn't change byte-for-byte identical, so a file-level diff
  shows only your edit, plus the diffs that demonstrate it. `diagnose` reports
  why an unreadable file cannot be opened.
- **`docx.composition`** — copy formatted content from one document into
  another, reconciling styles, numbering, media, hyperlinks, and bookmarks, and
  reporting every part it touched.
- **`docx.errors`** — every refusal is a typed exception, so a caller can tell
  a safe refusal apart from a bug.

## What it does not do

These limits are deliberate:

- **It never computes field values or paginates.** It writes fields and sets
  the update-on-open flag; Word or a headless renderer computes the results.
- **`compare` emits insertions and deletions only.** Moves are not synthesized;
  formatting-only, image, and object differences are reported rather than
  redlined.
- **No OLE authoring**, and composition refuses embedded objects.
- **No protection stripping and no decryption.** A password-protected file gets
  a typed refusal.
- **No document-quality judgment.** paper-docx edits documents; it does not
  grade them. Bad input surfaces as a specific, typed error rather than a raw
  traceback.
- **The formatting resolver states what it cannot resolve** (such as
  table-style conditional formatting or theme fonts) instead of guessing.

## How it stays compatible

- **Strict superset.** No existing behavior changes; new capability is new,
  explicitly named API. Upstream's own pytest and behave suites run on every
  commit.
- **Atomic operations.** Each operation validates fully before it mutates. A
  refused operation leaves the document, in memory and on disk, exactly as it
  was.
- **Whitespace is content.** No comparison or rewrite path normalizes
  meaningful text — a trailing space in a run is a real character.
- **Tested against real files.** A frozen, hash-verified fixture corpus spans
  generated and LibreOffice-authored documents and is checked against a
  headless LibreOffice load.

## Naming

Four names, kept distinct on purpose (the same pattern as Pillow and PIL):

- GitHub repository / PyPI distribution: **`paper-docx`**
- Python import package: **`docx`** — unchanged, so existing code and snippets
  keep working; drop-in compatibility is the point of the fork
- Fork sentinel: `docx.__paper_version__`

## Installation

```bash
pip install "paper-docx @ git+https://github.com/The-LLM-Data-Company/paper-docx.git@main"
```

Confirm the install:

```bash
python -c "import docx; print(docx.__paper_version__)"
```

## Repository map

```
src/docx/          the package (import name `docx`) — upstream modules plus:
  errors.py          typed, atomic refusals
  story.py           whole-document traversal, views, anchors
  search.py          normalized find and Span (replace, tracked replace, comment)
  blocks.py          paragraph-level insert / delete / replace
  revision.py        doc.revisions — enumerate and resolve tracked changes
  tableops.py        guarded table cell and row operations
  numbering.py       list reporting, application, and authoring
  controls.py        content-control enumeration and filling
  commentops.py      comment threads (reply, resolve, anchored text)
  bookmarks.py       bookmark enumerate / create / delete
  fields.py          field authoring (page, date, cross-reference, TOC)
  formatting.py      effective-format resolver
  protection.py      Restrict-Editing awareness
  scrubbing.py       doc.finalize / doc.scrub
  composition.py     cross-document composition
  package.py         xml_equivalent, diff_package, patch_save, diagnose,
                     text_diff, compare
docs/              Sphinx documentation (start at docs/user/paper-additions.rst)
tests/             upstream suites plus tests/paper (contract harness + fixtures)
```

## Documentation

The Sphinx docs extend the upstream python-docx documentation to cover the
fork's additions: start with `docs/user/paper-additions.rst` and the
`docs/api/paper-*.rst` reference pages. Everything inherited from python-docx
works as documented at the
[python-docx documentation](https://python-docx.readthedocs.io/).

## License and credit

paper-docx is distributed under the MIT License. `python-docx` is by Steve
Canny and contributors; this fork builds on that work and preserves its
license and attribution.
