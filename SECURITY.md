# Security Policy

paper-docx is regularly used to open untrusted `.docx` files, so we treat
package-parsing vulnerabilities (ZIP handling, OPC relationship resolution,
XML parsing) as high-severity issues.

## Reporting a vulnerability

Please do not report security vulnerabilities through public GitHub issues.

Instead, use GitHub's private vulnerability reporting for this repository:
[Report a vulnerability](https://github.com/paper-instruments/paper-docx/security/advisories/new).

Include, where possible:

- A description of the issue and its impact
- A minimal reproducing file or script
- The version of paper-docx affected (`docx.__paper_version__`)

## What to expect

- We will acknowledge your report promptly and keep you informed as we
  investigate.
- We will not take legal action against good-faith security research that
  respects user data and does not disrupt the service of others.
- Fixed vulnerabilities are credited to reporters in the release notes unless
  you prefer otherwise.

## Supported versions

Security fixes are applied to the latest released version on PyPI.
