"""Part-level package diff for changed-part budget assertions.

This is a *test-side* stand-in for the `docx.package` kernel. The assertion
interface in `contract.py` is intentionally narrow so the harness can ride on
the kernel's `diff_package` once it is available; this module then shrinks or
disappears.

Comparison semantics (mined from the reference `office_helpers/package.py`):

* Parts present on only one side are `added` / `removed`.
* Common parts whose bytes differ are `byte_changed`.
* A byte-changed part is *not* counted as `semantic_changed` when it is an XML
  part and its W3C-canonical form is identical on both sides. Canonicalization
  uses `xml.etree.ElementTree.canonicalize(..., strip_text=False)`: text nodes
  are preserved verbatim, so a meaningful trailing space inside `w:t` makes two
  parts UNEQUAL ("whitespace is content").
* Two OPC bookkeeping part types are *maps*, not documents, and are compared
  as data rather than as XML text (real-world producers differ freely in
  element order, inter-element whitespace, and inert entries):
  - `*.rels`: the set of (Id, Type, Target, TargetMode) tuples, with absent
    TargetMode meaning "Internal".
  - `[Content_Types].xml`: the override map plus the default map restricted
    to extensions actually present in that package — a Default for an
    extension no part uses (LibreOffice declares several) is inert.
* A byte-changed XML part that fails to parse on either side is conservatively
  counted as `semantic_changed` — a changed-and-malformed part can never be
  proven equivalent, and pretending otherwise would be a silent fallback.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Tuple
from xml.etree import ElementTree as ET


def is_xml_part_name(name: str) -> bool:
    """True for package part names that hold XML content by convention."""
    return name.endswith(".xml") or name.endswith(".rels")


def read_parts(path: Path) -> Dict[str, bytes]:
    """Mapping of member-name -> bytes for every part in the package at `path`."""
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def c14n_bytes(data: bytes) -> bytes:
    """W3C-canonical form of `data`, with all text content preserved verbatim.

    `data` is raw part bytes in whatever encoding its XML declaration names —
    expat handles the decode, so a UTF-16 part canonicalizes correctly and
    byte garbage raises plain `ParseError` rather than `UnicodeDecodeError`.

    Raises `xml.etree.ElementTree.ParseError` on malformed XML — callers decide
    how to classify that; this function never guesses.
    """
    out = io.StringIO()
    ET.canonicalize(xml_data=data, out=out, strip_text=False)
    return out.getvalue().encode("utf-8")


def xml_semantically_equal(a: bytes, b: bytes) -> bool:
    """True when `a` and `b` are canonically identical XML (whitespace preserved)."""
    return c14n_bytes(a) == c14n_bytes(b)


_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def relationship_set(data: bytes) -> Tuple:
    """Sorted (Id, Type, Target, TargetMode) tuples a `.rels` part defines.

    A sorted tuple (multiset), not a set: a corrupt part that duplicates an
    Id must not compare equal to a clean one that defines it once.
    """
    root = ET.fromstring(data)
    return tuple(
        sorted(
            (
                rel.attrib.get("Id"),
                rel.attrib.get("Type"),
                rel.attrib.get("Target"),
                rel.attrib.get("TargetMode") or "Internal",
            )
            for rel in root.iter(f"{{{_RELS_NS}}}Relationship")
        )
    )


def effective_content_types(data: bytes, part_names: "frozenset[str] | set") -> Tuple:
    """(active defaults, overrides) defined by a `[Content_Types].xml` blob.

    Defaults are restricted to extensions some part in `part_names` actually
    has — a Default for an unused extension is inert and must not register as
    semantic difference.
    """
    extensions_in_use = {
        name.rsplit(".", 1)[-1].lower() for name in part_names if "." in name
    }
    root = ET.fromstring(data)
    defaults = {
        elm.attrib.get("Extension", "").lower(): elm.attrib.get("ContentType")
        for elm in root.iter(f"{{{_CT_NS}}}Default")
    }
    overrides = {
        elm.attrib.get("PartName"): elm.attrib.get("ContentType")
        for elm in root.iter(f"{{{_CT_NS}}}Override")
    }
    active_defaults = {
        ext: ctype for ext, ctype in defaults.items() if ext in extensions_in_use
    }
    return active_defaults, overrides


def _part_semantically_equal(
    name: str,
    a: bytes,
    b: bytes,
    a_part_names: "set[str]",
    b_part_names: "set[str]",
) -> bool:
    """Part-type-aware semantic comparison; raises ET.ParseError on bad XML."""
    if name == "[Content_Types].xml":
        return effective_content_types(a, a_part_names) == effective_content_types(
            b, b_part_names
        )
    if name.endswith(".rels"):
        return relationship_set(a) == relationship_set(b)
    return xml_semantically_equal(a, b)


@dataclass(frozen=True)
class PartsDiff:
    """Result of a part-by-part comparison of two packages."""

    added: Tuple[str, ...]
    removed: Tuple[str, ...]
    byte_changed: Tuple[str, ...]
    semantic_changed: Tuple[str, ...]

    @property
    def is_semantically_empty(self) -> bool:
        """True when the packages differ in no part, semantically."""
        return not (self.added or self.removed or self.semantic_changed)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "schema": "paper_test_parts_diff",
            "version": 1,
            "added": list(self.added),
            "removed": list(self.removed),
            "byte_changed": list(self.byte_changed),
            "semantic_changed": list(self.semantic_changed),
        }


def diff_parts(original: Path, modified: Path) -> PartsDiff:
    """Part-level diff between the packages at `original` and `modified`."""
    left = read_parts(original)
    right = read_parts(modified)

    added = tuple(sorted(set(right) - set(left)))
    removed = tuple(sorted(set(left) - set(right)))

    byte_changed = []
    semantic_changed = []
    for name in sorted(set(left) & set(right)):
        if left[name] == right[name]:
            continue
        byte_changed.append(name)
        if is_xml_part_name(name):
            try:
                if _part_semantically_equal(
                    name, left[name], right[name], set(left), set(right)
                ):
                    continue
            except ET.ParseError:
                pass
        semantic_changed.append(name)

    return PartsDiff(
        added=added,
        removed=removed,
        byte_changed=tuple(byte_changed),
        semantic_changed=tuple(semantic_changed),
    )
