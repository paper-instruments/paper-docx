# paper-docx

`paper-docx` is Paper Instruments' hard fork of
[`python-docx`](https://github.com/python-openxml/python-docx), based on the
upstream `v1.2.0` release.

This fork is intended to remain a strict superset of upstream behavior. Existing
code such as `from docx import Document` must continue to work unchanged.

## Naming

There are four names to keep distinct:

- GitHub repository: `paper-docx`
- PyPI distribution: `paper-docx`
- Python import package: `docx`
- Fork sentinel: `docx.__paper_version__`

Built wheel files are named `paper_docx-*`, while the import remains `docx`.
That mismatch is intentional and follows the normal distribution/import split
used by packages such as Pillow/PIL. Do not rename `src/docx` to `src/paper_docx`.

## Installation

This repository is private and publication to PyPI is intentionally gated. For
now, install from Git:

```bash
pip install "paper-docx @ git+https://github.com/The-LLM-Data-Company/paper-docx.git@main"
```

Verify the fork sentinel:

```bash
python -c "import docx; print(docx.__paper_version__)"
```

See `PAPER.md` for fork lineage, baseline test results, and merge policy.
