"""Distribution/import identity regressions."""

from __future__ import annotations

import importlib
import importlib.util

import pytest

from docx import _guard, _version

PAPER_ONLY_MODULES = (
    "docx.blocks",
    "docx.bookmarks",
    "docx.commentops",
    "docx.composition",
    "docx.controls",
    "docx.errors",
    "docx.fields",
    "docx.formatting",
    "docx.numbering",
    "docx.protection",
    "docx.revision",
    "docx.search",
    "docx.story",
    "docx.tableops",
)


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

    def it_refuses_paper_only_modules_through_the_shared_guard(self, monkeypatch):
        monkeypatch.setattr(_version, "distribution", lambda _name: object())
        with pytest.raises(ImportError, match="both installed"):
            _guard.check_install()

    @pytest.mark.parametrize("module_name", PAPER_ONLY_MODULES)
    def it_refuses_each_paper_only_module_when_both_distributions_are_installed(
        self, monkeypatch, module_name
    ):
        for name in PAPER_ONLY_MODULES:
            importlib.import_module(name)
        loaded = importlib.import_module(module_name)
        monkeypatch.setattr(_version, "distribution", lambda _name: object())
        spec = importlib.util.spec_from_file_location(
            f"_paper_guard_probe_{module_name.replace('.', '_')}",
            loaded.__file__,
        )
        assert spec is not None
        assert spec.loader is not None
        probe = importlib.util.module_from_spec(spec)
        with pytest.raises(ImportError, match="both installed"):
            spec.loader.exec_module(probe)
