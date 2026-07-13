"""Mechanical enforcement of the approved public API surface.

Each entry pins the exact public signature approved for a module. Modules
that have not shipped yet skip cleanly; once a module lands, its surface is
enforced. Changing a signature means updating this table deliberately —
never silently.

Only parameter names, kinds and defaults are compared (annotations are
deliberately ignored; they are documentation, not surface).
"""

from __future__ import annotations

import importlib
import inspect

import pytest


def _canonical(signature: inspect.Signature) -> str:
    """Render a signature as names/kinds/defaults only, annotation-free."""
    parts = []
    seen_kw_marker = False
    for param in signature.parameters.values():
        if param.kind is inspect.Parameter.KEYWORD_ONLY and not seen_kw_marker:
            parts.append("*")
            seen_kw_marker = True
        if param.kind is inspect.Parameter.VAR_POSITIONAL:
            seen_kw_marker = True
            parts.append(f"*{param.name}")
            continue
        if param.kind is inspect.Parameter.VAR_KEYWORD:
            parts.append(f"**{param.name}")
            continue
        rendered = param.name
        if param.default is not inspect.Parameter.empty:
            rendered += f"={param.default!r}"
        parts.append(rendered)
    return f"({', '.join(parts)})"


def _resolve(module_name: str, attr_path: str):
    """The object at `module_name.attr_path`, skipping unshipped modules.

    A module missing entirely, or present but exposing NONE of its approved
    names (e.g. upstream `docx.package` before the kernel lands), means the
    module hasn't shipped -> skip. A module exposing SOME approved names but
    not the requested one is a broken surface -> fail.
    """
    module = pytest.importorskip(
        module_name, reason=f"{module_name} not yet implemented"
    )
    approved_names = {
        attr.split(".")[0] for mod, attr, _ in APPROVED_SIGNATURES if mod == module_name
    }
    if not any(hasattr(module, name) for name in approved_names):
        pytest.skip(f"{module_name} surface not yet implemented")
    target = module
    for piece in attr_path.split("."):
        if not hasattr(target, piece):
            pytest.fail(f"{module_name}.{attr_path}: {piece!r} missing from approved surface")
        target = getattr(target, piece)
    return target


#: (module, attr-path, approved canonical signature) — self is included for
#: methods accessed on the class.
APPROVED_SIGNATURES = [
    # -- kernel --------------------------------------------------------------
    ("docx.package", "xml_equivalent", "(a, b)"),
    ("docx.package", "diff_package", "(path_a, path_b)"),
    ("docx.package", "patch_save", "(original_path, document, out_path)"),
    # -- story traversal -----------------------------------------------------
    ("docx.story", "story_parts", "(document)"),
    ("docx.story", "iter_blocks", "(document, *, view='current')"),
    ("docx.story", "outline", "(document, *, view='current')"),
    # -- search --------------------------------------------------------------
    ("docx.search", "normalize_text", "(value)"),
    (
        "docx.search",
        "find_text",
        "(document, needle, *, nth=None, near=None, story=None, view='current')",
    ),
    (
        "docx.search",
        "find_one",
        "(document, needle, *, nth=None, near=None, story=None, view='current')",
    ),
    # -- replace -------------------------------------------------------------
    (
        "docx.search",
        "Span.replace",
        "(self, new_text, *, tracked=False, author=None, date=None)",
    ),
    # -- block operations ----------------------------------------------------
    (
        "docx.blocks",
        "insert_section_after",
        "(document, anchor, *, heading, paragraphs, heading_style='Heading2',"
        " body_style=None, tracked=False, author=None, date=None)",
    ),
    (
        "docx.blocks",
        "tracked_delete_paragraphs",
        "(document, start_anchor, *, end_anchor=None, count=1, author, date=None)",
    ),
    (
        "docx.blocks",
        "tracked_replace_paragraphs",
        "(document, start_anchor, replacement_paragraphs, *, end_anchor=None,"
        " count=1, body_style=None, author, date=None)",
    ),
    # -- revisions -----------------------------------------------------------
    ("docx.revision", "Revisions.accept_all", "(self, *, author=None)"),
    ("docx.revision", "Revisions.reject_all", "(self, *, author=None)"),
    # -- tables + numbering --------------------------------------------------
    ("docx.tableops", "find_table", "(document, *, near_text)"),
    (
        "docx.tableops",
        "update_cell",
        "(table, row, column, new_text, *, tracked=False, author=None, date=None)",
    ),
    (
        "docx.tableops",
        "insert_row_after",
        "(table, row, values, *, copy_format_from=None)",
    ),
    ("docx.tableops", "delete_row", "(table, row)"),
    ("docx.numbering", "list_numbering", "(document)"),
    ("docx.numbering", "apply_numbering", "(paragraph, *, num_id, level=0)"),
    ("docx.numbering", "apply_list_style", "(paragraph, style_name)"),
    # -- later additions -----------------------------------------------------
    ("docx.package", "diagnose", "(path)"),
    ("docx.package", "text_diff", "(path_a, path_b, *, view='current')"),
    ("docx.package", "pending_changes", "(path)"),
    (
        "docx.search",
        "replace_all",
        "(document, needle, new_text, *, story=None, view='current',"
        " tracked=False, author=None, date=None)",
    ),
    (
        "docx.search",
        "Span.comment",
        "(self, text, *, author, initials=None, date=None)",
    ),
    ("docx.controls", "list_controls", "(document)"),
    ("docx.controls", "get_control", "(document, *, tag=None, alias=None)"),
    (
        "docx.controls",
        "set_control_value",
        "(document, value, *, tag=None, alias=None)",
    ),
    ("docx.controls", "Control.set_value", "(self, value)"),
    ("docx.numbering", "ensure_bullet_definition", "(document)"),
    ("docx.numbering", "ensure_decimal_definition", "(document)"),
    ("docx.numbering", "restart_numbering", "(document, *, num_id)"),
    (
        "docx.blocks",
        "insert_blocks_after",
        "(document, anchor, *, blocks, tracked=False, author=None, date=None)",
    ),
    ("docx.commentops", "is_resolved", "(document, comment)"),
    ("docx.commentops", "resolve", "(document, comment, *, resolved=True)"),
    (
        "docx.commentops",
        "reply",
        "(document, comment, text, *, author, initials=None, date=None)",
    ),
    ("docx.commentops", "anchored_text", "(document, comment)"),
    ("docx.commentops", "comment_thread", "(document)"),
    ("docx.commentops", "parent_of", "(document, comment)"),
    ("docx.controls", "iter_controls", "(document)"),
    ("docx.revision", "Revisions.remaining_unsupported", "(self)"),
    # -- composition, bookmarks, fields, formatting --------------------------
    ("docx.document", "Document.finalize", "(self, *, revisions='accept')"),
    (
        "docx.document",
        "Document.scrub",
        "(self, *, comments=True, metadata=True, track_changes_setting=True,"
        " rsids=False, hidden_text=False)",
    ),
    ("docx.protection", "protection_status", "(document)"),
    ("docx.protection", "acknowledge_protection", "(document)"),
    (
        "docx.package",
        "compare",
        "(original, revised, *, author, date=None, granularity='word',"
        " materialize=None)",
    ),
    (
        "docx.composition",
        "insert_blocks_from",
        "(document, source, start_anchor, *, anchor, end_anchor=None,"
        " count=1, styles='match_by_name')",
    ),
    (
        "docx.composition",
        "append_document",
        "(document, source, *, section='new_page', styles='match_by_name')",
    ),
    ("docx.bookmarks", "list_bookmarks", "(document)"),
    ("docx.bookmarks", "create_bookmark", "(document, span, name)"),
    ("docx.bookmarks", "delete_bookmark", "(document, name)"),
    ("docx.fields", "add_page_number_field", "(paragraph)"),
    ("docx.fields", "add_page_count_field", "(paragraph)"),
    ("docx.fields", "add_date_field", "(paragraph, *, date_format=None)"),
    (
        "docx.fields",
        "add_reference_field",
        "(paragraph, *, bookmark, kind='text')",
    ),
    ("docx.fields", "insert_toc_after", "(document, anchor, *, levels=(1, 3))"),
    ("docx.formatting", "format_of", "(target)"),
    ("docx.formatting", "surrounding_format", "(document, anchor)"),
]

_IDS = [f"{module}.{attr}" for module, attr, _ in APPROVED_SIGNATURES]


class DescribeApprovedApiSurface:
    @pytest.mark.parametrize(
        ("module_name", "attr_path", "approved"), APPROVED_SIGNATURES, ids=_IDS
    )
    def it_matches_the_approved_signature(
        self, module_name: str, attr_path: str, approved: str
    ):
        target = _resolve(module_name, attr_path)
        actual = _canonical(inspect.signature(target))
        assert actual == approved, (
            f"{module_name}.{attr_path}: signature {actual} deviates from approved"
            f" {approved}; update this table in the same commit"
        )

    def it_pins_the_refusal_hierarchy(self):
        errors = pytest.importorskip(
            "docx.errors", reason="docx.errors not yet implemented"
        )
        assert issubclass(errors.PaperRefusal, Exception)
        for name in (
            "AmbiguousTargetError",
            "TargetNotFoundError",
            "UnsupportedStructureError",
            "BoundaryViolationError",
            "RelationshipPolicyError",
            "DocumentProtectedError",
        ):
            subclass = getattr(errors, name)
            assert issubclass(subclass, errors.PaperRefusal), name

    def it_exposes_the_kernel_at_the_pinned_public_path(self):
        """The public path is docx.package.*."""
        package = pytest.importorskip("docx.package")
        if not hasattr(package, "patch_save"):
            pytest.skip("kernel not yet implemented")
        for name in ("xml_equivalent", "diff_package", "patch_save", "PackageDiff"):
            assert hasattr(package, name), f"docx.package.{name} missing"
        # the upstream surface must be intact, not shadowed
        assert hasattr(package, "Package") and hasattr(package, "ImageParts")

    def it_keeps_document_revisions_additive(self):
        import docx.document

        if not hasattr(docx.document.Document, "revisions"):
            pytest.skip("Document.revisions not yet implemented")
        assert isinstance(
            inspect.getattr_static(docx.document.Document, "revisions"), property
        )
