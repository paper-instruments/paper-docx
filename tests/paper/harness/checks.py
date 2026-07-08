"""Mechanical package-fact checks, ported from the reference `verify_docx.py`.

Only the checks that assert *package facts* are ported (CONVENTIONS §5):
required parts, XML parseability, undefined style references, undefined
numbering references, broken relationship targets, and fake-bullet detection.
The reference's domain/styling checks (placeholder text, minimum paragraph
count, heading requirements) are deliberately NOT ported — they are
example-level policy, not package facts.

Each `find_*` function returns a list of findings (empty == clean); each
`assert_no_*` wrapper raises `AssertionError` with the findings spelled out.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "rel": REL_NS}

REQUIRED_PARTS = frozenset(
    {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
    }
)

#: Literal-bullet detector: paragraph text that starts with a hyphen/asterisk
#: or a typographic bullet character followed by whitespace and content.
FAKE_BULLET_RE = re.compile(r"^\s*(?:[-*]|[•‣▪▫])\s+\S")

#: Style-reference tags checked against the style definitions, tag -> label.
_STYLE_REF_TAGS = (f"{{{W_NS}}}pStyle", f"{{{W_NS}}}rStyle", f"{{{W_NS}}}tblStyle")


def _qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _read_xml_parts(path: Path) -> Tuple[Dict[str, ET.Element], List[Tuple[str, str]]]:
    """Parse every `.xml`/`.rels` part; return (parsed roots, parse failures)."""
    roots: Dict[str, ET.Element] = {}
    failures: List[Tuple[str, str]] = []
    with zipfile.ZipFile(path) as zf:
        for name in sorted(zf.namelist()):
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                roots[name] = ET.fromstring(zf.read(name))
            except ET.ParseError as exc:
                failures.append((name, str(exc)))
    return roots, failures


def paragraph_text(p: ET.Element) -> str:
    """Visible-ish text of a `w:p` element, tracked/field text included.

    Mirrors the reference verifier's `node_text`: `w:t`, `w:delText` and
    `w:instrText` contribute text; `w:tab` -> TAB; `w:br` -> newline.
    """
    parts: List[str] = []
    for child in p.iter():
        if child.tag in {_qn("t"), _qn("delText"), _qn("instrText")}:
            parts.append(child.text or "")
        elif child.tag == _qn("tab"):
            parts.append("\t")
        elif child.tag == _qn("br"):
            parts.append("\n")
    return "".join(parts)


def find_missing_required_parts(path: Path) -> List[str]:
    """Names from REQUIRED_PARTS absent from the package at `path`."""
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
    return sorted(REQUIRED_PARTS - names)


def find_unparseable_xml_parts(path: Path) -> List[Tuple[str, str]]:
    """(part-name, parse-error) for every `.xml`/`.rels` part that fails to parse."""
    _, failures = _read_xml_parts(path)
    return failures


def _story_scope_of(part_name: str) -> Optional[str]:
    """The story scope a `word/` part resolves its style/numbering refs in.

    Desktop Word's glossary document (building blocks, cover pages, control
    placeholder text) is its own story with its own `styles.xml` and
    `numbering.xml` under `word/glossary/`; references there must never be
    checked against the main document's tables.
    """
    if part_name.startswith("word/glossary/"):
        return "word/glossary/"
    if part_name.startswith("word/"):
        return "word/"
    return None


def find_undefined_style_references(path: Path) -> List[Tuple[str, str]]:
    """(part-name, style-id) for style references with no definition in scope.

    Checks `w:pStyle`, `w:rStyle` and `w:tblStyle` in every XML part under
    `word/` (body, headers, footers, footnotes, comments, ...), each against
    its own scope's styles part (`word/styles.xml`, or
    `word/glossary/styles.xml` for the glossary story). A scope with no
    styles part is skipped — nothing to check against.
    """
    roots, _ = _read_xml_parts(path)
    defined_by_scope: Dict[str, set] = {}
    for scope in ("word/", "word/glossary/"):
        styles_root = roots.get(scope + "styles.xml")
        if styles_root is not None:
            defined_by_scope[scope] = {
                style_id
                for style in styles_root.findall(f"{{{W_NS}}}style")
                if (style_id := style.attrib.get(_qn("styleId")))
            }
    findings: List[Tuple[str, str]] = []
    for name, root in sorted(roots.items()):
        scope = _story_scope_of(name)
        if scope is None or name == scope + "styles.xml" or scope not in defined_by_scope:
            continue
        defined = defined_by_scope[scope]
        for node in root.iter():
            if node.tag in _STYLE_REF_TAGS:
                value = node.attrib.get(_qn("val"))
                if value and value not in defined:
                    findings.append((name, value))
    return findings


def find_undefined_numbering_references(path: Path) -> List[Tuple[str, str]]:
    """(part-name, numId) for `w:numPr/w:numId` references with no definition.

    A `w:num` element with matching `w:numId` must exist in the scope's
    numbering part (`word/numbering.xml`, or `word/glossary/numbering.xml`
    for the glossary story). Unlike the reference verifier, a reference with
    *no numbering part at all* is reported too: the reference is dangling
    either way. `numId` value "0" is exempt — ECMA-376 defines it as "no
    numbering" (it removes inherited numbering rather than referencing a
    definition).
    """
    roots, _ = _read_xml_parts(path)
    defined_by_scope: Dict[str, set] = {}
    for scope in ("word/", "word/glossary/"):
        numbering_root = roots.get(scope + "numbering.xml")
        defined_by_scope[scope] = (
            set()
            if numbering_root is None
            else {
                num_id
                for num in numbering_root.findall(f"{{{W_NS}}}num")
                if (num_id := num.attrib.get(_qn("numId")))
            }
        )
    findings: List[Tuple[str, str]] = []
    for name, root in sorted(roots.items()):
        scope = _story_scope_of(name)
        if scope is None or name == scope + "numbering.xml":
            continue
        defined = defined_by_scope[scope]
        for num_id_elm in root.iter(_qn("numId")):
            value = num_id_elm.attrib.get(_qn("val"))
            if value and value != "0" and value not in defined:
                findings.append((name, value))
    return findings


def find_broken_relationship_targets(path: Path) -> List[Tuple[str, str]]:
    """(rels-part-name, target) for internal relationship targets that don't exist."""
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
    roots, _ = _read_xml_parts(path)
    findings: List[Tuple[str, str]] = []
    for rels_name, root in sorted(roots.items()):
        if not rels_name.endswith(".rels"):
            continue
        for rel in root.findall(f"{{{REL_NS}}}Relationship"):
            target = rel.attrib.get("Target")
            if not target or rel.attrib.get("TargetMode") == "External":
                continue
            base = posixpath.dirname(rels_name)
            if posixpath.basename(base) == "_rels":
                base = posixpath.dirname(base)
            normalized = posixpath.normpath(posixpath.join(base, target))
            if normalized.startswith("/"):
                # absolute part URI -> package-root relative member name
                normalized = normalized[1:]
            escapes_root = normalized == ".." or normalized.startswith("../")
            if escapes_root or normalized not in names:
                findings.append((rels_name, target))
    return findings


def find_fake_bullet_paragraphs(path: Path) -> List[int]:
    """1-based indices of word/document.xml paragraphs that fake list bullets.

    A paragraph counts when its text starts with a literal bullet/hyphen marker
    but it carries no real `w:numPr` numbering.
    """
    roots, _ = _read_xml_parts(path)
    document = roots.get("word/document.xml")
    if document is None:
        return []
    findings: List[int] = []
    for index, p in enumerate(document.iter(_qn("p")), start=1):
        has_numbering = p.find(f"{_qn('pPr')}/{_qn('numPr')}") is not None
        if not has_numbering and FAKE_BULLET_RE.match(paragraph_text(p).strip()):
            findings.append(index)
    return findings


def _assert_empty(findings: object, description: str) -> None:
    assert not findings, f"{description}: {findings!r}"


def assert_package_is_wellformed(path: Path) -> None:
    """Required parts present and every XML part parseable."""
    _assert_empty(find_missing_required_parts(path), f"{path.name}: missing required parts")
    _assert_empty(find_unparseable_xml_parts(path), f"{path.name}: unparseable XML parts")


def assert_no_undefined_style_references(path: Path) -> None:
    _assert_empty(
        find_undefined_style_references(path), f"{path.name}: undefined style references"
    )


def assert_no_undefined_numbering_references(path: Path) -> None:
    _assert_empty(
        find_undefined_numbering_references(path),
        f"{path.name}: undefined numbering references",
    )


def assert_no_broken_relationship_targets(path: Path) -> None:
    _assert_empty(
        find_broken_relationship_targets(path),
        f"{path.name}: broken relationship targets",
    )


def assert_package_facts_clean(path: Path) -> None:
    """The full battery of positive package-fact checks."""
    assert_package_is_wellformed(path)
    assert_no_broken_relationship_targets(path)
    assert_no_undefined_style_references(path)
    assert_no_undefined_numbering_references(path)
