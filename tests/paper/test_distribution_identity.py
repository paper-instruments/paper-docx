"""Distribution/import identity regressions."""

from __future__ import annotations

import pytest

from docx import _version


class DescribeDistributionIdentity:
    def it_refuses_when_both_distributions_are_installed(self, monkeypatch):
        monkeypatch.setattr(_version, "distribution", lambda _name: object())
        with pytest.raises(ImportError, match="both installed"):
            _version.assert_distribution_identity()

    def it_allows_a_clean_paper_docx_install(self, monkeypatch):
        def distribution(name):
            if name == "python-docx":
                raise _version.PackageNotFoundError(name)
            return object()

        monkeypatch.setattr(_version, "distribution", distribution)
        _version.assert_distribution_identity()
