"""Contract-harness assertion utilities.

Every mutating paper API must pass five assertions. The helpers here are the
shared machinery for them:

1. Save -> reopen: `save_and_reopen` — never assert on the in-memory object.
2. Intended effect present in the reopened document: the caller's own asserts,
   made against the document `save_and_reopen` returns.
3. Changed-part budget: `assert_changed_parts` — the package diff between
   input and output shows exactly the expected parts changed, nothing else.
4. Independent-loader smoke: `harness.lo.assert_libreoffice_opens`, under the
   `lo_smoke` marker.
5. Refusal atomicity: `assert_refusal_atomic` — the typed refusal is raised
   AND neither the in-memory XML tree nor any file on disk changed.
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict, Iterable, Tuple, Type

import docx
from docx.document import Document

from .pkgdiff import PartsDiff, diff_parts


def save_and_reopen(document: Document, path: Path) -> Document:
    """Save `document` to `path` and return a freshly loaded Document from disk.

    The classic silent failure is an edit that lands in the tree but never
    reaches disk; asserting only on reopened documents rules it out.
    """
    document.save(str(path))
    return docx.Document(str(path))


def snapshot_package(document: Document) -> Dict[str, str]:
    """{partname: blob-sha256} for every part of `document`'s package, plus rels.

    XML part blobs are serialized from the live element tree, so two snapshots
    taken before and after an operation compare equal exactly when the
    operation left no trace in the in-memory package.
    """
    import hashlib

    package = document.part.package
    assert package is not None
    snapshot: Dict[str, str] = {}
    # the package's own relationship collection serializes to _rels/.rels on
    # save; "<package>" cannot collide with a real partname (those start "/")
    if package.rels:
        snapshot["<package>:rels"] = hashlib.sha256(package.rels.xml).hexdigest()
    for part in package.iter_parts():
        snapshot[str(part.partname)] = hashlib.sha256(part.blob).hexdigest()
        rels_xml = part.rels.xml if part.rels else None
        if rels_xml is not None:
            snapshot[str(part.partname) + ":rels"] = hashlib.sha256(rels_xml).hexdigest()
    return snapshot


def assert_changed_parts(
    original: Path,
    modified: Path,
    expected_changed: Iterable[str],
    *,
    expected_added: Iterable[str] = (),
    expected_removed: Iterable[str] = (),
) -> PartsDiff:
    """Assert the semantic package diff is exactly the expected change set.

    `expected_changed` names parts whose content may differ semantically;
    `expected_added`/`expected_removed` name parts that may appear/disappear.
    Anything outside those sets — in either direction — fails the assertion.
    Byte-level churn that is semantically neutral (reserialization noise) is
    tolerated by design until `patch_save` makes it disappear.

    Returns the diff so callers can make further assertions on it.
    """
    diff = diff_parts(original, modified)
    problems = []
    if set(diff.semantic_changed) != set(expected_changed):
        problems.append(
            f"changed parts {sorted(diff.semantic_changed)!r}"
            f" != expected {sorted(expected_changed)!r}"
        )
    if set(diff.added) != set(expected_added):
        problems.append(f"added parts {sorted(diff.added)!r} != expected {sorted(expected_added)!r}")
    if set(diff.removed) != set(expected_removed):
        problems.append(
            f"removed parts {sorted(diff.removed)!r} != expected {sorted(expected_removed)!r}"
        )
    assert not problems, "; ".join(problems)
    return diff


def assert_no_op_roundtrip_is_semantically_clean(original: Path, resaved: Path) -> PartsDiff:
    """Assert an open->save round trip changed nothing semantically.

    (Byte identity is the *patch_save* kernel invariant; upstream `save()`
    reserializes and is not byte-stable.)
    """
    diff = diff_parts(original, resaved)
    assert diff.is_semantically_empty, (
        "no-op round trip changed parts semantically: "
        f"added={diff.added!r} removed={diff.removed!r} changed={diff.semantic_changed!r}"
    )
    return diff


def assert_refusal_atomic(
    document: Document,
    operation: Callable[[Document], object],
    refusal_type: Type[BaseException],
    *,
    on_disk: Tuple[Path, ...] = (),
) -> BaseException:
    """Assert `operation(document)` raises `refusal_type` and mutates nothing.

    "Nothing" means: every part blob (and rels) of the in-memory package is
    byte-identical before and after, and every path in `on_disk` has identical
    bytes before and after. Returns the raised exception so callers can assert
    on its message.
    """
    before_memory = snapshot_package(document)
    before_disk = {path: path.read_bytes() for path in on_disk}

    raised: BaseException | None = None
    try:
        operation(document)
    except refusal_type as exc:
        raised = exc
    assert raised is not None, (
        f"operation completed without raising {refusal_type.__name__};"
        " a documented refusal input must refuse"
    )

    after_memory = snapshot_package(document)
    assert after_memory == before_memory, (
        "refused operation left a trace in the in-memory package; parts differing: "
        f"{sorted(k for k in before_memory if before_memory[k] != after_memory.get(k))!r}"
    )
    for path, before_bytes in before_disk.items():
        assert path.read_bytes() == before_bytes, (
            f"refused operation modified {path} on disk"
        )
    return raised
