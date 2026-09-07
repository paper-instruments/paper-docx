<div align="center">
  <a href="https://github.com/paper-instruments/paper-docx">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/paper-instruments/paper-docx/main/.github/assets/logo-dark.svg">
      <img alt="paper-docx logo" src="https://raw.githubusercontent.com/paper-instruments/paper-docx/main/.github/assets/logo-light.svg" height="128">
    </picture>
  </a>
  <h1>paper-docx</h1>

[![PyPI](https://img.shields.io/pypi/v/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![Python versions](https://img.shields.io/pypi/pyversions/paper-docx.svg)](https://pypi.org/project/paper-docx/)
[![Test](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml/badge.svg)](https://github.com/paper-instruments/paper-docx/actions/workflows/test.yml)

</div>

**An import-compatible, agent-safe fork of python-docx designed to prevent silent corruption when editing existing Word documents.**

`paper-docx` is an import-compatible hard fork of [python-docx](https://github.com/python-openxml/python-docx) for working with existing Microsoft Word (`.docx`) documents. It keeps python-docx's package layer, XML mapping, and object model, and adds guarded inspection and editing APIs that refuse unsupported operations instead of guessing.

```python
import docx   # the import name is unchanged; see "Import compatibility"
```

---

## Why paper-docx exists

`python-docx` is excellent at *creating* documents. Its lossless package layer, disciplined XML mapping, and years of absorbed edge cases are why this fork builds on it.

The harder problem is changing a contract or other real-world document without losing formatting, revisions, fields, or content outside the body. Hand-edited XML can produce **silent corruption**: a file that opens fine and is quietly wrong. An agent cannot eyeball the result, so it needs the document's structure and every edit outcome as typed, machine-readable data. It also needs the library to refuse rather than guess.

## Quick start

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

`compare` emits markup Word renders as tracked changes. Before returning, it accepts and rejects private copies and verifies both outcomes. If it cannot represent a difference as tracked revisions, it raises a typed refusal instead of returning an incomplete redline.

## What paper-docx adds

`paper-docx` adds guarded APIs for editing existing documents and handling review markup. It also supports document comparison and composition. See [Paper additions](docs/user/paper-additions.rst) for supported operations and refusal conditions.

## Safety contract

Paper mutating APIs validate before changing the document or restore the package if a later step fails. A `PaperRefusal` leaves the document unchanged and identifies the condition that prevented the operation. Invalid arguments and I/O failures use ordinary Python exceptions.

## Import compatibility

The distribution and repository use the name `paper-docx`; Python code continues to import `docx`. Existing import statements do not change. Paper adds package validation and guarded edits, so it can reject some inputs and operations that python-docx accepts.

- GitHub repository / PyPI distribution: **`paper-docx`**
- Python import: **`docx`**
- Fork sentinel: `docx.__paper_version__ = "0.2.0"`

## Installation

Install from PyPI:

```bash
python -m pip uninstall -y python-docx paper-docx
python -m pip install paper-docx
```

The clean uninstall is required when migrating from `python-docx`. Both distributions use the frozen `docx` import package, and pip cannot safely overlay or uninstall two distributions that own the same files.

Confirm the install:

```bash
paper-docx-doctor
```

Pip does not treat `paper-docx` as satisfying another package's declared dependency on `python-docx`. That dependency will reinstall upstream and overwrite shared `docx` files. Replace or remove the dependency, or run that package in a separate environment.

In a controlled deployment, a constraint containing `python-docx<0` makes pip reject direct or transitive attempts to install upstream. The constraint must be applied to every install in that environment.

## Documentation

The Sphinx docs cover Paper-specific behavior and stricter validation: start with `docs/user/paper-additions.rst` and the `docs/api/paper-*.rst` reference pages. For inherited APIs, see the [python-docx documentation](https://python-docx.readthedocs.io/).

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](https://github.com/paper-instruments/paper-docx/blob/main/CONTRIBUTING.md) for development setup and change requirements.

## Community

- **Bugs and feature requests**: [GitHub Issues](https://github.com/paper-instruments/paper-docx/issues)
- **Questions and ideas**: [GitHub Discussions](https://github.com/paper-instruments/paper-docx/discussions)
- **Security**: see [SECURITY.md](https://github.com/paper-instruments/paper-docx/blob/main/SECURITY.md)

## Acknowledgments

paper-docx exists because python-docx's package layer and XML mapping are excellent. Thanks to Steve Canny and the python-docx contributors for the work this project builds on. Upstream python-docx lives at [github.com/python-openxml/python-docx](https://github.com/python-openxml/python-docx).

## Citation

If you reference paper-docx in research or writing:

```bibtex
@software{paper_docx,
  title   = {paper-docx: an agent-first structure editor for Word documents},
  author  = {{Paper Instruments, Inc.}},
  year    = {2026},
  url     = {https://github.com/paper-instruments/paper-docx}
}
```

Cite it as a fork of *python-docx* by Steve Canny and contributors.

## License

MIT, inherited from python-docx. Original work © Steve Canny and the python-docx contributors; fork additions © Paper Instruments, Inc. This fork preserves the upstream license and attribution. See [LICENSE](https://github.com/paper-instruments/paper-docx/blob/main/LICENSE).
