<div align="center">
  <a href="https://github.com/paper-instruments/paper-docx">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset=".github/assets/logo-dark.svg">
      <img alt="paper-docx logo" src=".github/assets/logo-light.svg" height="128">
    </picture>
  </a>
  <h1>paper-docx</h1>

[![PyPI](https://img.shields.io/pypi/v/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![License](https://img.shields.io/pypi/l/paper-docx.svg)](LICENSE)
[![CI](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml)
[![Downloads](https://img.shields.io/pypi/dm/paper-docx.svg)](https://pypi.org/project/paper-docx/)

</div>

<div align="center">
  <a href="https://github.com/paper-instruments/paper-docx">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset=".github/assets/logo-dark.svg">
      <img alt="paper-docx logo" src=".github/assets/logo-light.svg" height="128">
    </picture>
  </a>
  <h1>paper-docx</h1>

[![PyPI](https://img.shields.io/pypi/v/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![License](https://img.shields.io/pypi/l/paper-docx.svg)](LICENSE)
[![CI](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml)

`paper-docx` is an agent-first Python library for safely inspecting, editing,
reviewing, and composing existing Microsoft Word (`.docx`) documents. It is a
strict-superset hard fork of
[`python-docx`](https://github.com/python-openxml/python-docx) and a drop-in
replacement. The distribution is renamed; the import name stays `docx`, so
existing code keeps working unchanged.

```python
import docx                       # the import name is unchanged
doc = docx.Document("contract.docx")
```

## Why it exists

`python-docx` is excellent at *creating* documents. Its lossless package layer,
disciplined XML mapping, and years of absorbed edge cases are why this fork
builds on it.

The harder problem is changing a contract or other real-world document without
losing formatting, revisions, fields, or content outside the body. Hand-edited
XML can produce **silent corruption**: a file that opens fine and is quietly
wrong. An agent cannot eyeball the result, so it needs the document's structure
and every edit outcome as typed, machine-readable data. It also needs the
library to refuse rather than guess.

## Safety contract

Every added operation either does exactly what it claims or refuses atomically.
Gates such as Restrict-Editing are checked before anything is touched, and
compound edits capture the live package first and restore it if they raise. If
an operation cannot proceed safely, it raises a typed `PaperRefusal` and leaves
the document byte-for-byte unchanged in memory and on disk. Callers can catch `PaperRefusal`
separately from programmer errors, which remain plain `ValueError` or
`TypeError`. Comparison and rewrite paths preserve meaningful whitespace,
including trailing spaces inside runs.

## A short example

Create a native Word redline from two document versions:

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

`compare` emits markup Word renders as tracked changes. Before returning, it
accepts and rejects private copies and verifies both outcomes. If a difference
cannot be represented safely as a redline, such as a style or package-part
change, it raises a typed refusal instead of returning an incomplete result.

## What it adds

### Reading and editing one document

- **`docx.story`** traverses the body, headers, footers, footnotes, endnotes,
  comments, tracked insertions, content controls, and text boxes. Callers can
  view the document as it stands, before pending revisions, or all at once.
- **`docx.search`** finds normalized text across Word's run fragmentation. A
  returned `Span` can replace the matched text while preserving unaffected
  runs, emit the replacement as a tracked change, or anchor a comment.
- **`docx.blocks`** inserts, deletes, or replaces whole paragraphs relative to
  a text anchor, plainly or as a tracked change.
- **`docx.tableops` / `docx.numbering`** provide cell, row, and list edits that
  refuse on unsafe structures such as merged cells, nested tables, or undefined
  numbering.
- **`docx.controls`** fills content controls with the correct value type and
  clears placeholder state so Word treats them as filled.
- **`docx.bookmarks` / `docx.fields`** create bookmarks over a span and author
  page numbers, dates, cross-references, captions, and tables of contents as
  fields with placeholder results.
- **`docx.notes` / `docx.links`** anchor real footnotes, endnotes, and
  hyperlinks to a matched span, so they compose with `docx.search` rather than
  needing their own targeting.
- **`Drawing.replace_picture`** swaps the image behind a drawing in place,
  leaving its size, position, and identity alone.
- **`docx.formatting`** resolves effective formatting through document defaults,
  styles, and direct formatting, with provenance for each value.

### Reviewing

- **`doc.revisions`** enumerates and resolves tracked changes across every part:
  insertions, deletions, run and paragraph format changes, table-row revisions,
  and moves. Unresolvable markup is listed by name.
- **`docx.comments` / `docx.commentops`** read comments and delete one by
  identity, keeping the modern comment identity parts consistent.
- **`docx.protection`** reads, sets, and respects Restrict-Editing. Mutating
  operations refuse on a protected document unless the caller explicitly
  overrides, and the setting is preserved in the document.

### Working across documents

- **`docx.package.compare`** generates a native tracked-change redline from two
  documents, with the accept/reject round-trip shown above.
- **`docx.package.patch_save` / `diff_package` / `text_diff`** keeps unchanged
  parts byte-identical and reports changed parts and text. `diagnose` explains
  why an unreadable file cannot be opened.
- **`docx.composition`** copies formatted content between documents, reconciles
  styles, numbering, media, hyperlinks, and bookmarks, and reports every part
  touched.
- **`docx.errors`** exposes typed refusals, distinct from programmer errors.

## Drop-in and name map

Only the distribution and repository are renamed. The importable package stays
`docx`. This is the same distribution/import split as Pillow
(`pip install pillow`, `import PIL`), and it preserves existing code, snippets,
and model priors.

- GitHub repository / PyPI distribution: **`paper-docx`**
- Python import: **`docx`**
- Fork sentinel: `docx.__paper_version__ = "0.2.0"`

## Installation

Install from PyPI:

```bash
python -m pip uninstall -y python-docx paper-docx
python -m pip install paper-docx
```

The clean uninstall is required when migrating from `python-docx`. Both
distributions use the frozen `docx` import package, and pip cannot safely
overlay or uninstall two distributions that own the same files.

Confirm the install:

```bash
paper-docx-doctor
```

Pip does not treat `paper-docx` as satisfying another package's declared
dependency on `python-docx`. That dependency will reinstall upstream and
overwrite shared `docx` files. Replace or remove the dependency, or run that
package in a separate environment.

In a controlled deployment, a constraint containing `python-docx<0` makes pip
reject direct or transitive attempts to install upstream. The constraint must
be applied to every install in that environment.

## Documentation

The Sphinx docs extend the upstream python-docx documentation to cover the
fork's additions: start with `docs/user/paper-additions.rst` and the
`docs/api/paper-*.rst` reference pages. Everything inherited from python-docx
works as documented at the
[python-docx documentation](https://python-docx.readthedocs.io/).

## How it's tested

- Upstream's pytest and behave suites run on every commit to check compatibility
  with existing behavior.
- A frozen, hash-pinned fixture corpus spans generated and LibreOffice-authored
  documents.
- The contract harness checks refusal atomicity and validates the fixture corpus
  with a headless LibreOffice load smoke.

## License

MIT, inherited from python-docx. Original work © Steve Canny and the python-docx
contributors; fork additions © Paper Instruments, Inc. This fork preserves the
upstream license and attribution. See [`LICENSE`](LICENSE).
