"""Package kernel: semantic XML comparison, package diff, and narrow save.

Implementation of the package kernel. The public import path is
`docx.package.{xml_equivalent, diff_package, patch_save}` — that upstream
module re-exports these names; this module exists so its diff stays minimal.

Design constraints honored here:

* **Whitespace is content.** Parsing uses a dedicated lxml parser WITHOUT
  `remove_blank_text`, and text/tail compare verbatim — a preserved trailing
  space inside `w:t` makes two parts non-equivalent.
* **Map-like OPC parts compare as data.** `*.rels` parts compare as the
  multiset of (Id, Type, Target, TargetMode) tuples; `[Content_Types].xml`
  compares as (defaults restricted to extensions in use, overrides) maps.
  Real producers differ freely in element order, inter-element whitespace and
  inert defaults; none of that is semantic change.
* **Compare-based `patch_save`** — no opc-internals changes, no dirty flags.
  Deterministic zip output (candidate entry order, fixed 1980-01-01 entry
  timestamps); all writes go through a temp file + `os.replace`, so a
  mid-write failure leaves any existing output file intact. A save in which
  nothing changed writes a verbatim copy of the original file (no-op round
  trip is byte-identical).
"""

from __future__ import annotations

import hashlib
import io
import os
import stat
import tempfile
import zipfile
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path
from typing import IO, TYPE_CHECKING, Dict, List, Tuple, Union

from lxml import etree

from docx._zipguard import (
    _END_RECORD,
    _END_RECORD_SIGNATURE,
    _MAX_END_COMMENT_BYTES,
    GuardedZipReader,
    preflight_zip,
    read_compressed_bytes,
)
from docx.errors import MalformedPackageError

if TYPE_CHECKING:
    from docx.document import Document

_PathLike = Union[str, "os.PathLike[str]"]

#: Parser for comparison only: keeps whitespace-only text nodes (unlike the
#: oxml parser) and never resolves entities.
_compare_parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_CONTENT_TYPES_PART = "[Content_Types].xml"

#: Fixed zip entry timestamp for deterministic `patch_save` output.
_FIXED_ZIP_DATE = (1980, 1, 1, 0, 0, 0)


class UnsupportedXmlError(ValueError):
    """XML this kernel refuses to compare (fail loudly, never guess)."""


def _parse(data: bytes) -> etree._Element:
    """Parse raw part bytes for comparison. Raises XMLSyntaxError loudly.

    Parts carrying an internal DTD subset are refused: unresolved entity
    references would compare by name while their replacement text differs — a
    false-EQUAL — and ISO 29500 forbids DOCTYPE in package parts anyway.
    """
    root = etree.fromstring(data, _compare_parser)
    docinfo = root.getroottree().docinfo
    if docinfo.internalDTD is not None or docinfo.doctype:
        raise UnsupportedXmlError(
            "part carries a DOCTYPE/internal DTD subset; OPC parts must not"
            " (ISO 29500), and entity semantics cannot be compared safely"
        )
    return root


def _doc_level_nodes(root: etree._Element) -> Tuple:
    """Comments/PIs outside the root element (prolog and epilog), in order.

    An `<?mso-application?>` prolog PI changes how Windows treats the file;
    differences there are semantic and must not be invisible.
    """
    prolog = []
    node = root.getprevious()
    while node is not None:
        prolog.append((str(node.tag), getattr(node, "target", None), node.text or ""))
        node = node.getprevious()
    epilog = []
    node = root.getnext()
    while node is not None:
        epilog.append((str(node.tag), getattr(node, "target", None), node.text or ""))
        node = node.getnext()
    return tuple(reversed(prolog)), tuple(epilog)


def _elements_equal(a: etree._Element, b: etree._Element) -> bool:
    if a.tag != b.tag:
        return False
    # processing instructions share one tag object; the target is identity
    if not isinstance(a.tag, str) and getattr(a, "target", None) != getattr(
        b, "target", None
    ):
        return False
    if dict(a.attrib) != dict(b.attrib):
        return False
    if (a.text or "") != (b.text or ""):
        return False
    a_children = list(a)
    b_children = list(b)
    if len(a_children) != len(b_children):
        return False
    for a_child, b_child in zip(a_children, b_children):
        if (a_child.tail or "") != (b_child.tail or ""):
            return False
        if not _elements_equal(a_child, b_child):
            return False
    return True


def xml_equivalent(a: bytes, b: bytes) -> bool:
    """True when `a` and `b` are structurally identical XML documents.

    Tags and attribute names compare in Clark notation, so namespace *prefix*
    choices never matter while namespace URIs always do. Attribute order is
    insignificant (XML defines attributes as unordered); child order is
    significant; text and tail content compare verbatim, whitespace included
    (a canonicalizer that trims a meaningful trailing space would corrupt
    documents through `patch_save`). Prolog/epilog comments and
    processing instructions are compared too.

    Raises `lxml.etree.XMLSyntaxError` on malformed input and
    ``UnsupportedXmlError`` on DTD-bearing input — this function never guesses.

    Known limit: attribute VALUES holding
    QNames compare textually, so a prefix rebound to a different URI while
    the QName text stays identical is not detected. OOXML producers keep the
    standard prefixes, and the error direction in `diff_package` remains
    conservative for everything else.
    """
    root_a = _parse(a)
    root_b = _parse(b)
    if _doc_level_nodes(root_a) != _doc_level_nodes(root_b):
        return False
    return _elements_equal(root_a, root_b)


def _relationship_multiset(data: bytes) -> Tuple[Tuple[str, ...], ...]:
    root = _parse(data)
    return tuple(
        sorted(
            (
                rel.get("Id") or "",
                rel.get("Type") or "",
                rel.get("Target") or "",
                rel.get("TargetMode") or "Internal",
            )
            for rel in root.iter(f"{{{_RELS_NS}}}Relationship")
        )
    )


def _effective_content_types(
    data: bytes, part_names: "List[str]"
) -> Tuple[Dict[str, str], Dict[str, str]]:
    """(active defaults, overrides) — defaults restricted to in-use extensions."""
    extensions_in_use = {
        name.rsplit(".", 1)[-1].lower() for name in part_names if "." in name
    }
    root = _parse(data)
    defaults = {
        (elm.get("Extension") or "").lower(): elm.get("ContentType") or ""
        for elm in root.iter(f"{{{_CT_NS}}}Default")
    }
    overrides = {
        elm.get("PartName") or "": elm.get("ContentType") or ""
        for elm in root.iter(f"{{{_CT_NS}}}Override")
    }
    active = {ext: ctype for ext, ctype in defaults.items() if ext in extensions_in_use}
    return active, overrides


def is_xml_part_name(name: str) -> bool:
    """True for member names that hold XML content by OPC convention."""
    return name.endswith(".xml") or name.endswith(".rels")


def _parts_semantically_equal(
    name: str,
    a: bytes,
    b: bytes,
    a_names: "List[str]",
    b_names: "List[str]",
) -> bool:
    """Part-type-aware equivalence; malformed XML is conservatively unequal."""
    try:
        if name == _CONTENT_TYPES_PART:
            return _effective_content_types(a, a_names) == _effective_content_types(
                b, b_names
            )
        if name.endswith(".rels"):
            return _relationship_multiset(a) == _relationship_multiset(b)
        return xml_equivalent(a, b)
    except (etree.XMLSyntaxError, UnsupportedXmlError):
        # a changed part that cannot be parsed (or compared safely) can never
        # be proven equivalent
        return False


def _read_zip(
    source: Union[_PathLike, IO[bytes], bytes]
) -> Tuple[Dict[str, bytes], List[str]]:
    """Return validated member bytes and central-directory order."""
    if isinstance(source, (str, os.PathLike)):
        with open(source, "rb") as stream:
            return _read_zip(stream)
    source = io.BytesIO(source) if isinstance(source, bytes) else source
    preflight_zip(source)
    with zipfile.ZipFile(source) as zf:
        return GuardedZipReader(zf).read_all()


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@dataclass(frozen=True)
class PartDiff:
    """One byte-changed part in a package comparison."""

    part: str
    kind: str  # "xml" | "binary"
    before_sha256: str
    after_sha256: str
    semantic_change: bool

    def to_dict(self) -> dict:
        return {
            "part": self.part,
            "kind": self.kind,
            "before_sha256": self.before_sha256,
            "after_sha256": self.after_sha256,
            "semantic_change": self.semantic_change,
        }


@dataclass(frozen=True)
class PackageDiff:
    """Part-by-part comparison of two OPC packages."""

    added: Tuple[str, ...]
    removed: Tuple[str, ...]
    changed: Tuple[PartDiff, ...]
    byte_identical_count: int

    def semantic_changed_parts(self) -> Tuple[str, ...]:
        return tuple(item.part for item in self.changed if item.semantic_change)

    @property
    def is_semantically_empty(self) -> bool:
        """True when the packages hold the same parts with the same meaning."""
        return not (self.added or self.removed or self.semantic_changed_parts())

    def to_dict(self) -> dict:
        return {
            "schema": "paper_package_diff",
            "version": 1,
            "added": list(self.added),
            "removed": list(self.removed),
            "changed": [item.to_dict() for item in self.changed],
            "byte_identical_count": self.byte_identical_count,
            "is_semantically_empty": self.is_semantically_empty,
        }


def diff_package(path_a: _PathLike, path_b: _PathLike) -> PackageDiff:
    """Part-by-part diff of the OPC packages at `path_a` and `path_b`.

    XML parts (`*.xml`, `*.rels`) compare semantically, binary parts by bytes. A
    byte-changed XML part that fails to parse counts as semantically changed, never
    silently equal. Raises `MalformedPackageError` when either package will not open.
    """
    a_parts, a_order = _read_zip(Path(path_a))
    b_parts, b_order = _read_zip(Path(path_b))

    added = tuple(sorted(set(b_parts) - set(a_parts)))
    removed = tuple(sorted(set(a_parts) - set(b_parts)))

    changed: List[PartDiff] = []
    byte_identical = 0
    for name in sorted(set(a_parts) & set(b_parts)):
        before, after = a_parts[name], b_parts[name]
        if before == after:
            byte_identical += 1
            continue
        if is_xml_part_name(name):
            semantic = not _parts_semantically_equal(name, before, after, a_order, b_order)
            kind = "xml"
        else:
            semantic = True
            kind = "binary"
        changed.append(
            PartDiff(
                part=name,
                kind=kind,
                before_sha256=_sha256(before),
                after_sha256=_sha256(after),
                semantic_change=semantic,
            )
        )
    return PackageDiff(
        added=added,
        removed=removed,
        changed=tuple(changed),
        byte_identical_count=byte_identical,
    )


@dataclass(frozen=True)
class PatchSaveResult:
    """Outcome of a `patch_save` call."""

    restored_parts: Tuple[str, ...]
    changed_parts: Tuple[str, ...]
    added_parts: Tuple[str, ...]
    removed_parts: Tuple[str, ...]
    verbatim_copy: bool

    def to_dict(self) -> dict:
        return {
            "schema": "paper_patch_save",
            "version": 1,
            "restored_parts": list(self.restored_parts),
            "changed_parts": list(self.changed_parts),
            "added_parts": list(self.added_parts),
            "removed_parts": list(self.removed_parts),
            "verbatim_copy": self.verbatim_copy,
        }


def _deterministic_zip_bytes(parts: Dict[str, bytes], order: List[str]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            info = zipfile.ZipInfo(name, date_time=_FIXED_ZIP_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            zf.writestr(info, parts[name])
    return buf.getvalue()


def _write_bytes_atomically(out_path: Path, data: bytes, mode: int) -> None:
    """Write via a same-directory temp file + `os.replace`.

    Follows a destination symlink (updates the target, keeps the link).
    `fsync`s the file and directory. A failure before the final rename leaves
    any existing file intact; the temp file is always cleaned up.
    """
    path = os.fspath(out_path)
    link_state = None
    if os.path.islink(path):
        link_stat = os.lstat(path)
        destination = os.path.realpath(path)
        link_state = (link_stat.st_dev, link_stat.st_ino, destination)
    else:
        destination = path
    directory = os.path.dirname(os.path.abspath(destination)) or os.curdir
    os.makedirs(directory, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(
        prefix=os.path.basename(destination) + ".",
        suffix=".partial",
        dir=directory,
    )
    temp_path = Path(temp_name)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.chmod(temp_path, mode)
        if link_state is not None:
            try:
                current = os.lstat(path)
            except OSError as exc:
                raise OSError(
                    "destination symlink changed during save; nothing was replaced"
                ) from exc
            if (
                not stat.S_ISLNK(current.st_mode)
                or (current.st_dev, current.st_ino) != link_state[:2]
                or os.path.realpath(path) != link_state[2]
            ):
                raise OSError(
                    "destination symlink changed during save; nothing was replaced"
                )
        os.replace(temp_path, destination)
        with suppress(OSError):
            directory_fd = os.open(directory, os.O_RDONLY)
            try:
                os.fsync(directory_fd)
            finally:
                os.close(directory_fd)
    finally:
        if temp_path.exists():
            temp_path.unlink()


#: OLE Compound File header — encrypted Office files and legacy .doc binaries
_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _has_end_of_central_directory(path: Path) -> bool:
    """Whether the file ends with a structurally plausible end-of-central-directory record.

    A file carrying one is a ZIP whose archive simply does not start at byte 0, which is a
    different defect from having no ZIP structure at all. The record must *account for the rest
    of the file* — its declared comment length has to reach exactly the end — which is the same
    rule the preflight applies. Testing for the four signature bytes alone would misread any
    non-ZIP that happens to contain them: measured, a text file with those bytes planted in its
    tail was reported as a prefixed archive.
    """
    try:
        size = path.stat().st_size
        with path.open("rb") as stream:
            window = min(size, _END_RECORD.size + _MAX_END_COMMENT_BYTES)
            stream.seek(size - window)
            tail = stream.read(window)
    except OSError:
        return False
    cursor = tail.rfind(_END_RECORD_SIGNATURE)
    while cursor >= 0:
        if cursor + _END_RECORD.size <= len(tail):
            comment_length = _END_RECORD.unpack_from(tail, cursor)[7]
            if cursor + _END_RECORD.size + comment_length == len(tail):
                return True
        cursor = tail.rfind(_END_RECORD_SIGNATURE, 0, cursor)
    return False


_MAIN_PART_KINDS = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": "docx",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml": "dotx",
    "application/vnd.ms-word.document.macroEnabled.main+xml": "docm",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml": "dotm",
}


@dataclass(frozen=True)
class PackageDiagnosis:
    """Typed triage for a file `docx.Document()` may refuse or crash on.

    `readable` means "this package can open it as a WordprocessingML
    document"; `kind` names what the file actually is; `problems` say why it
    is not readable (empty when it is).
    """

    path: str
    readable: bool
    kind: str
    problems: Tuple[str, ...]

    def to_dict(self) -> dict:
        return {
            "schema": "paper_diagnosis",
            "version": 1,
            "path": self.path,
            "readable": self.readable,
            "kind": self.kind,
            "problems": list(self.problems),
        }


def diagnose(path: _PathLike) -> PackageDiagnosis:
    """Why the file at `path` can or cannot be opened, as a `PackageDiagnosis`.

    Returns instead of raising, so you can triage untrusted input before opening it.
    `Document()` refuses corrupt and encrypted packages with `MalformedPackageError` but
    still raises a bare `ValueError` for macro-enabled and template files.
    """
    file_path = Path(path)

    def result(readable: bool, kind: str, *problems: str) -> PackageDiagnosis:
        return PackageDiagnosis(
            path=str(file_path), readable=readable, kind=kind, problems=tuple(problems)
        )

    if not file_path.is_file():
        return result(False, "missing", "file does not exist")
    with file_path.open("rb") as stream:
        header = stream.read(8)
    if header.startswith(_CFB_MAGIC):
        return result(
            False,
            "encrypted-or-legacy-binary",
            "OLE compound file: either a password-protected Office document"
            " (decrypt it first) or a legacy binary .doc (convert to .docx)",
        )
    if not header.startswith(b"PK"):
        # -- a package whose archive does not start at byte 0 is still a ZIP; saying
        # -- "not a ZIP archive" would send the caller looking for the wrong problem
        if _has_end_of_central_directory(file_path):
            return result(
                False,
                "unsafe-archive",
                "a ZIP archive is present but does not begin at the start of the file;"
                " Word cannot open a .docx with anything in front of the package."
                " Extract the archive to a file of its own and open that",
            )
        return result(False, "not-a-zip", "not a ZIP archive, so not an OPC package")
    try:
        parts, _ = _read_zip(file_path)
    except MalformedPackageError as exc:
        return result(False, "unsafe-archive", str(exc))
    except zipfile.BadZipFile as exc:
        return result(False, "corrupt-zip", f"ZIP structure is damaged: {exc}")

    problems = [
        f"required part {name!r} is missing"
        for name in (_CONTENT_TYPES_PART, "_rels/.rels")
        if name not in parts
    ]
    main_type = None
    main_part_name = "word/document.xml"
    if "_rels/.rels" in parts:
        try:
            for rel_id, rel_type, target, mode in _relationship_multiset(
                parts["_rels/.rels"]
            ):
                if rel_type.endswith("/officeDocument") and mode == "Internal":
                    main_part_name = target.lstrip("/")
                    break
        except (etree.XMLSyntaxError, UnsupportedXmlError):
            pass
    if _CONTENT_TYPES_PART in parts:
        try:
            _, overrides = _effective_content_types(
                parts[_CONTENT_TYPES_PART], list(parts)
            )
            main_type = overrides.get("/" + main_part_name)
        except (etree.XMLSyntaxError, UnsupportedXmlError) as exc:
            problems.append(f"[Content_Types].xml is unparseable or unsafe: {exc}")

    if main_part_name not in parts:
        for marker, kind in (("xl/", "xlsx"), ("ppt/", "pptx")):
            if any(name.startswith(marker) for name in parts):
                return result(
                    False, kind, "an OPC package, but not a WordprocessingML one"
                )
        return result(
            False, "opc-unknown", f"no {main_part_name} main document part",
            *problems,
        )

    kind = _MAIN_PART_KINDS.get(main_type or "", "docx" if main_type is None else "opc-unknown")
    if kind in ("docm", "dotm"):
        problems.append(
            f"macro-enabled Office file ({kind}); python-docx opens only"
            " plain .docx — resave without macros"
        )
    elif kind == "dotx":
        problems.append(
            "Word template (.dotx); python-docx opens only .docx — resave as"
            " a document"
        )
    elif kind == "opc-unknown":
        problems.append(f"unrecognized main-part content type {main_type!r}")
    return result(not problems, kind, *problems)


def patch_save(
    original_path: _PathLike, document: "Document", out_path: _PathLike
) -> PatchSaveResult:
    """Save `document` to `out_path`, restoring original bytes wherever possible.

    The document is serialized normally, then every XML part semantically identical to its
    counterpart in `original_path` gets that counterpart's exact original bytes back, a
    narrow save that keeps unrelated parts byte-stable through the open/edit/save cycle.
    When nothing changed, `out_path` becomes a verbatim copy of the original.

    `original_path == out_path` is permitted; the original bytes are read up front and the
    write is atomic (temp file + rename), so the original survives a mid-write failure. An
    existing destination keeps its permission bits; a new one inherits `original_path`'s.
    Creates the destination's parent directory when it is missing.

    Returns a `PatchSaveResult` naming what it restored and what it rewrote. Raises
    `MalformedPackageError` when `original_path`, the serialized document, or the finished
    output fails to reparse, and `OSError` when a destination symlink moves mid-save.
    """
    original_file = Path(original_path)
    output_file = Path(out_path)
    try:
        output_mode = stat.S_IMODE(output_file.stat().st_mode)
    except FileNotFoundError:
        output_mode = stat.S_IMODE(original_file.stat().st_mode)

    original_bytes = read_compressed_bytes(original_file)
    original_parts, _ = _read_zip(io.BytesIO(original_bytes))

    buf = io.BytesIO()
    document.save(buf)
    candidate_bytes = buf.getvalue()

    candidate_parts, candidate_order = _read_zip(io.BytesIO(candidate_bytes))
    original_names = list(original_parts)
    candidate_names = list(candidate_parts)

    restored: List[str] = []
    changed: List[str] = []
    for name in candidate_order:
        if name not in original_parts:
            continue
        before, after = original_parts[name], candidate_parts[name]
        if before == after:
            continue
        if is_xml_part_name(name) and _parts_semantically_equal(
            name, before, after, original_names, candidate_names
        ):
            candidate_parts[name] = before
            restored.append(name)
        else:
            changed.append(name)

    added = tuple(sorted(set(candidate_parts) - set(original_parts)))
    removed = tuple(sorted(set(original_parts) - set(candidate_parts)))

    is_noop = (
        not added
        and not removed
        and all(candidate_parts[name] == original_parts[name] for name in candidate_parts)
    )
    if is_noop:
        output_bytes = original_bytes
    else:
        output_bytes = _deterministic_zip_bytes(candidate_parts, candidate_order)
    _read_zip(output_bytes)
    _write_bytes_atomically(output_file, output_bytes, output_mode)
    return PatchSaveResult(
        restored_parts=tuple(restored),
        changed_parts=tuple(changed),
        added_parts=added,
        removed_parts=removed,
        verbatim_copy=is_noop,
    )
