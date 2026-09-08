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

**An import-compatible, agent-safe fork of python-docx for creating and editing Word documents without silent corruption.**

`paper-docx` is an import-compatible hard fork of [python-docx](https://github.com/python-openxml/python-docx) for creating and editing Microsoft Word (`.docx`) documents. It keeps the `docx` import name and upstream object model, and adds guarded inspection and editing APIs that refuse unsupported operations instead of guessing.

## Installation

```bash
python -m pip uninstall -y python-docx paper-docx
python -m pip install paper-docx
paper-docx-doctor
```

Both distributions provide the `docx` import package. Do not install `python-docx` and `paper-docx` in the same environment.

## Quick start

Create a Word redline from two document versions:

```python
import docx
from docx.package import compare

original = docx.Document()
original.add_paragraph("Payment is due within thirty calendar days.")
original.save("original.docx")

revised = docx.Document()
revised.add_paragraph("Payment is due within thirty business days.")
revised.save("revised.docx")

result = compare("original.docx", "revised.docx", author="Reviewer")
result.document.save("redline.docx")
print(result.revision_count)  # 2
```

`compare` writes differences as tracked changes and refuses changes it cannot represent safely.

## Documentation

Read the [paper-docx documentation](https://docs.paperinstruments.com/).

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](https://github.com/paper-instruments/paper-docx/blob/main/CONTRIBUTING.md).

## Acknowledgments

paper-docx builds on python-docx by Steve Canny and contributors. This fork preserves their API, license, and attribution.

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

MIT, inherited from python-docx. Original work © Steve Canny and the python-docx contributors; fork additions © Paper Instruments, Inc. See [LICENSE](https://github.com/paper-instruments/paper-docx/blob/main/LICENSE).
