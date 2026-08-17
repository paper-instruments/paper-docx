"""Each paper-added module must declare its public surface, accurately.

`__all__` is load-bearing for two consumers that fail quietly when it is wrong. `griffe check`
reads it to decide whether a change breaks the public API, so a name missing from it can be
removed without the gate objecting. The documentation generator in `paper-office-docs` reads it
to decide what to publish, so a name missing from it silently vanishes from the reference
instead of failing a build.

This is the companion to `test_api_surface.py`, not a duplicate of it. That module pins the
exact *signature* of each approved function; this one asserts the module says which names are
public at all. Neither catches what the other does: a function can keep its signature while
falling out of `__all__`, and a name can be exported while its signature drifts.

Both assertions here matter, in opposite directions. A name in `__all__` that does not exist is
a lie. A public name absent from `__all__` is an omission — and `docx.search.normalize_text`
is why that assertion earns its keep: it is a deliberate re-export from `docx._normalize`, so it
is public and approved while not being *defined* in the module that exports it.

`docx.package` is deliberately absent from this test. It is an upstream module that paper
extended - it exists at the `paper-base` tag - so declaring its public surface would be a
change to inherited code. Its paper-added members are named explicitly by the generator instead.
"""

from __future__ import annotations

import ast
import importlib
import pathlib

import pytest

PAPER_MODULES = (
    "blocks",
    "bookmarks",
    "commentops",
    "composition",
    "controls",
    "errors",
    "fields",
    "formatting",
    "numbering",
    "protection",
    "revision",
    "scrubbing",
    "search",
    "story",
    "tableops",
)

SRC = pathlib.Path(__file__).resolve().parents[2] / "src" / "docx"

# -- Names a module exports without defining. `search` re-exports `normalize_text` from
# -- `docx._normalize` on purpose, marked `# noqa: F401 - public re-export`, and
# -- test_api_surface.py pins its signature as approved public API.
REEXPORTS = {"search": {"normalize_text"}}


def _defined_public_names(module_name: str) -> set:
    """Return the public names `module_name` defines at module level, read from source.

    Read statically rather than through `dir()`, which cannot distinguish a name the module
    defines from one it merely imported.
    """
    tree = ast.parse((SRC / f"{module_name}.py").read_text(encoding="utf-8"))
    names = set()
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            if not node.name.startswith("_"):
                names.add(node.name)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id.isupper() and target.id[0] != "_":
                    names.add(target.id)
    return names


class DescribeThePublicSurface:
    """Every paper module declares __all__, and it matches what the module offers."""

    @pytest.mark.parametrize("module_name", PAPER_MODULES)
    def it_declares_all(self, module_name: str):
        module = importlib.import_module(f"docx.{module_name}")
        assert hasattr(module, "__all__"), (
            f"docx.{module_name} declares no __all__. Both `griffe check` and the "
            f"documentation generator read it; without one, neither knows what is public."
        )

    @pytest.mark.parametrize("module_name", PAPER_MODULES)
    def it_exports_only_names_that_resolve(self, module_name: str):
        module = importlib.import_module(f"docx.{module_name}")
        missing = sorted(n for n in module.__all__ if not hasattr(module, n))
        assert not missing, (
            f"docx.{module_name}.__all__ exports {missing}, which the module does not "
            f"provide. Remove them, or restore the names."
        )

    @pytest.mark.parametrize("module_name", PAPER_MODULES)
    def it_leaves_no_public_name_unexported(self, module_name: str):
        module = importlib.import_module(f"docx.{module_name}")
        expected = _defined_public_names(module_name) | REEXPORTS.get(module_name, set())
        unexported = sorted(expected - set(module.__all__))
        assert not unexported, (
            f"docx.{module_name} offers public {unexported} but does not export them. An "
            f"unexported public name is invisible to `griffe check` and absent from the "
            f"generated reference. Add them to __all__, or make them private if they were "
            f"never meant to be public."
        )

    def it_leaves_the_upstream_package_module_undeclared(self):
        """`docx.package` is inherited; its public surface is not ours to declare."""
        import docx.package

        assert not hasattr(docx.package, "__all__"), (
            "docx.package is an upstream module (present at the paper-base tag). Declaring "
            "__all__ there changes inherited code. Its paper-added members are named "
            "explicitly by the documentation generator instead."
        )
