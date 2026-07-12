"""Objects that implement reading and writing OPC packages."""

from __future__ import annotations

import os
import secrets
import stat
import tempfile
from contextlib import suppress
from typing import IO, TYPE_CHECKING, BinaryIO, Iterator, Optional, cast

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PACKAGE_URI, PackURI
from docx.opc.part import PartFactory
from docx.opc.parts.coreprops import CorePropertiesPart
from docx.opc.pkgreader import PackageReader
from docx.opc.pkgwriter import PackageWriter
from docx.opc.rel import Relationships
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from typing_extensions import Self

    from docx.opc.coreprops import CoreProperties
    from docx.opc.part import Part
    from docx.opc.rel import _Relationship  # pyright: ignore[reportPrivateUsage]


class OpcPackage:
    """Main API class for |python-opc|.

    A new instance is constructed by calling the :meth:`open` class method with a path
    to a package file or file-like object containing one.
    """

    def after_unmarshal(self):
        """Entry point for any post-unmarshaling processing.

        May be overridden by subclasses without forwarding call to super.
        """
        # don't place any code here, just catch call if not overridden by
        # subclass
        pass

    @property
    def core_properties(self) -> CoreProperties:
        """|CoreProperties| object providing read/write access to the Dublin Core
        properties for this document."""
        return self._core_properties_part.core_properties

    def iter_rels(self) -> Iterator[_Relationship]:
        """Generate exactly one reference to each relationship in the package by
        performing a depth-first traversal of the rels graph."""

        def walk_rels(
            source: OpcPackage | Part, visited: list[Part] | None = None
        ) -> Iterator[_Relationship]:
            visited = [] if visited is None else visited
            for rel in source.rels.values():
                yield rel
                if rel.is_external:
                    continue
                part = rel.target_part
                if part in visited:
                    continue
                visited.append(part)
                new_source = part
                for rel in walk_rels(new_source, visited):
                    yield rel

        for rel in walk_rels(self):
            yield rel

    def iter_parts(self) -> Iterator[Part]:
        """Generate exactly one reference to each of the parts in the package by
        performing a depth-first traversal of the rels graph."""

        def walk_parts(source, visited=[]):
            for rel in source.rels.values():
                if rel.is_external:
                    continue
                part = rel.target_part
                if part in visited:
                    continue
                visited.append(part)
                yield part
                new_source = part
                for part in walk_parts(new_source, visited):
                    yield part

        for part in walk_parts(self):
            yield part

    def load_rel(self, reltype: str, target: Part | str, rId: str, is_external: bool = False):
        """Return newly added |_Relationship| instance of `reltype` between this part
        and `target` with key `rId`.

        Target mode is set to ``RTM.EXTERNAL`` if `is_external` is |True|. Intended for
        use during load from a serialized package, where the rId is well known. Other
        methods exist for adding a new relationship to the package during processing.
        """
        return self.rels.add_relationship(reltype, target, rId, is_external)

    @property
    def main_document_part(self):
        """Return a reference to the main document part for this package.

        Examples include a document part for a WordprocessingML package, a presentation
        part for a PresentationML package, or a workbook part for a SpreadsheetML
        package.
        """
        return self.part_related_by(RT.OFFICE_DOCUMENT)

    def next_partname(self, template: str) -> PackURI:
        """Return a |PackURI| instance representing partname matching `template`.

        The returned part-name has the next available numeric suffix to distinguish it
        from other parts of its type. `template` is a printf (%)-style template string
        containing a single replacement item, a '%d' to be used to insert the integer
        portion of the partname. Example: "/word/header%d.xml"
        """
        partnames = {part.partname for part in self.iter_parts()}
        for n in range(1, len(partnames) + 2):
            candidate_partname = template % n
            if candidate_partname not in partnames:
                return PackURI(candidate_partname)

    @classmethod
    def open(cls, pkg_file: str | IO[bytes]) -> Self:
        """Return an |OpcPackage| instance loaded with the contents of `pkg_file`."""
        pkg_reader = PackageReader.from_file(pkg_file)
        package = cls()
        Unmarshaller.unmarshal(pkg_reader, package, PartFactory)
        return package

    def part_related_by(self, reltype: str) -> Part:
        """Return part to which this package has a relationship of `reltype`.

        Raises |KeyError| if no such relationship is found and |ValueError| if more than
        one such relationship is found.
        """
        return self.rels.part_with_reltype(reltype)

    @property
    def parts(self) -> list[Part]:
        """Return a list containing a reference to each of the parts in this package."""
        return list(self.iter_parts())

    def relate_to(self, part: Part, reltype: str):
        """Return rId key of new or existing relationship to `part`.

        If a relationship of `reltype` to `part` already exists, its rId is returned. Otherwise a
        new relationship is created and that rId is returned.
        """
        rel = self.rels.get_or_add(reltype, part)
        return rel.rId

    @lazyproperty
    def rels(self):
        """Return a reference to the |Relationships| instance holding the collection of
        relationships for this package."""
        return Relationships(PACKAGE_URI.baseURI)

    def save(self, pkg_file: str | IO[bytes]):
        """Save this package to `pkg_file`.

        `pkg_file` can be either a file-path or a file-like object.
        """
        parts = self.parts
        for part in parts:
            part.before_marshal()
        parts = self.parts
        if isinstance(pkg_file, (str, os.PathLike)):
            _atomic_package_write(os.fspath(pkg_file), self.rels, parts)
        else:
            _atomic_stream_write(pkg_file, self.rels, parts)

    @property
    def _core_properties_part(self) -> CorePropertiesPart:
        """|CorePropertiesPart| object related to this package.

        Creates a default core properties part if one is not present (not common).
        """
        try:
            return cast(CorePropertiesPart, self.part_related_by(RT.CORE_PROPERTIES))
        except KeyError:
            core_properties_part = CorePropertiesPart.default(self)
            self.relate_to(core_properties_part, RT.CORE_PROPERTIES)
            return core_properties_part


class Unmarshaller:
    """Hosts static methods for unmarshalling a package from a |PackageReader|."""

    @staticmethod
    def unmarshal(pkg_reader, package, part_factory):
        """Construct graph of parts and realized relationships based on the contents of
        `pkg_reader`, delegating construction of each part to `part_factory`.

        Package relationships are added to `pkg`.
        """
        parts = Unmarshaller._unmarshal_parts(pkg_reader, package, part_factory)
        Unmarshaller._unmarshal_relationships(pkg_reader, package, parts)
        for part in parts.values():
            part.after_unmarshal()
        package.after_unmarshal()

    @staticmethod
    def _unmarshal_parts(pkg_reader, package, part_factory):
        """Return a dictionary of |Part| instances unmarshalled from `pkg_reader`, keyed
        by partname.

        Side-effect is that each part in `pkg_reader` is constructed using
        `part_factory`.
        """
        parts = {}
        for partname, content_type, reltype, blob in pkg_reader.iter_sparts():
            parts[partname] = part_factory(partname, content_type, reltype, blob, package)
        return parts

    @staticmethod
    def _unmarshal_relationships(pkg_reader, package, parts):
        """Add a relationship to the source object corresponding to each of the
        relationships in `pkg_reader` with its target_part set to the actual target part
        in `parts`."""
        for source_uri, srel in pkg_reader.iter_srels():
            source = package if source_uri == "/" else parts[source_uri]
            target = srel.target_ref if srel.is_external else parts[srel.target_partname]
            source.load_rel(srel.reltype, target, srel.rId, srel.is_external)


def _validate_serialized_output(source) -> None:
    """Reopen a staged package so malformed output never reaches its destination."""
    reader = PackageReader.from_file(source)
    # Force traversal of the complete serialized relationship graph.
    tuple(reader.iter_sparts())
    tuple(reader.iter_srels())


def _atomic_package_write(path: str, relationships, parts) -> None:
    """Serialize beside ``path``, validate, and atomically replace it."""
    link_state = None
    if os.path.islink(path):
        link_stat = os.lstat(path)
        destination = os.path.realpath(path)
        link_state = (link_stat.st_dev, link_stat.st_ino, destination)
    else:
        destination = path
    directory = os.path.dirname(os.path.abspath(destination)) or os.curdir
    descriptor, temporary = _new_atomic_temp(directory, f".{os.path.basename(destination)}.")
    try:
        if os.path.exists(destination):
            mode = stat.S_IMODE(os.stat(destination).st_mode)
            if hasattr(os, "fchmod"):
                os.fchmod(descriptor, mode)
            else:
                os.chmod(temporary, mode)
        os.close(descriptor)
        descriptor = -1
        PackageWriter.write(temporary, relationships, parts)
        _validate_serialized_output(temporary)
        with open(temporary, "rb") as staged:
            os.fsync(staged.fileno())
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
        os.replace(temporary, destination)
        with suppress(OSError):
            directory_fd = os.open(directory, os.O_RDONLY)
            try:
                os.fsync(directory_fd)
            finally:
                os.close(directory_fd)
    finally:
        if descriptor >= 0:
            os.close(descriptor)
        with suppress(OSError):
            os.unlink(temporary)


def _new_atomic_temp(directory: str, prefix: str) -> "tuple[int, str]":
    """Create a sibling temporary file using normal umask semantics."""
    for _ in range(100):
        path = os.path.join(directory, f"{prefix}{secrets.token_hex(8)}.partial")
        try:
            return os.open(path, os.O_RDWR | os.O_CREAT | os.O_EXCL, 0o666), path
        except FileExistsError:
            continue
    raise FileExistsError("could not allocate a unique atomic-save temp file")


_STREAM_COPY_BYTES = 1024 * 1024
_STREAM_SPOOL_BYTES = 8 * 1024 * 1024


def _atomic_stream_write(stream: IO[bytes], relationships, parts) -> None:
    """Stage output and roll back a readable, seekable destination on failure."""
    start = _stream_position(stream)
    snapshot_record = _snapshot_stream_tail(stream, start)
    if snapshot_record is None:
        if _has_stream_rollback_surface(stream, start):
            raise OSError(
                "destination stream exposes rollback operations but its existing"
                " content could not be snapshotted; nothing was written"
            )
        _write_staged_to_unrestorable_stream(
            stream, relationships, parts, start=start
        )
        return
    snapshot, original_end = snapshot_record
    try:
        with tempfile.SpooledTemporaryFile(max_size=_STREAM_SPOOL_BYTES, mode="w+b") as staged:
            assert start is not None
            _copy_stream_prefix(stream, staged, start)
            PackageWriter.write(staged, relationships, parts)
            _validate_serialized_output(staged)
            staged.seek(start)
            try:
                stream.seek(start)
                _copy_stream(staged, cast(BinaryIO, stream))
                stream.truncate()
                flush = getattr(stream, "flush", None)
                if callable(flush):
                    flush()
            except BaseException:
                try:
                    _restore_stream_tail(stream, start, original_end, snapshot)
                except BaseException as rollback_error:
                    raise RuntimeError(
                        "stream save failed and the destination stream could not be restored"
                    ) from rollback_error
                raise
    finally:
        snapshot.close()


def _has_stream_rollback_surface(stream, start: "Optional[int]") -> bool:
    return start is not None and all(
        callable(getattr(stream, name, None)) for name in ("read", "seek", "tell", "truncate")
    )


def _write_staged_to_unrestorable_stream(
    stream, relationships, parts, *, start: "Optional[int]"
) -> None:
    """Validate first; after commit starts, write-only stream errors are final."""
    if start not in (None, 0):
        raise OSError(
            "cannot validate a write-only stream with an existing prefix;"
            " nothing was written"
        )
    with tempfile.SpooledTemporaryFile(max_size=_STREAM_SPOOL_BYTES, mode="w+b") as staged:
        PackageWriter.write(staged, relationships, parts)
        staged.seek(0)
        _validate_serialized_output(staged)
        staged.seek(0)
        _copy_stream(staged, cast(BinaryIO, stream))
        flush = getattr(stream, "flush", None)
        if callable(flush):
            flush()


def _copy_stream_prefix(stream, destination: BinaryIO, length: int) -> None:
    """Copy the exact preserved prefix into the staged validation stream."""
    stream.seek(0)
    remaining = length
    while remaining:
        chunk = stream.read(min(_STREAM_COPY_BYTES, remaining))
        if not isinstance(chunk, bytes) or not chunk:
            raise OSError(
                "destination stream prefix could not be read; nothing was written"
            )
        destination.write(chunk)
        remaining -= len(chunk)
    stream.seek(length)


def _stream_position(stream) -> "Optional[int]":
    try:
        position = stream.tell()
    except (AttributeError, OSError, TypeError, ValueError):
        return None
    return (
        position
        if isinstance(position, int) and not isinstance(position, bool) and position >= 0
        else None
    )


def _snapshot_stream_tail(stream, start) -> "Optional[tuple[BinaryIO, int]]":
    if not _has_stream_rollback_surface(stream, start):
        return None
    snapshot = tempfile.SpooledTemporaryFile(  # noqa: SIM115 - returned to caller
        max_size=_STREAM_SPOOL_BYTES, mode="w+b"
    )
    try:
        stream.seek(0, os.SEEK_END)
        original_end = stream.tell()
        stream.seek(start)
        remaining = original_end - start
        while remaining:
            chunk = stream.read(min(_STREAM_COPY_BYTES, remaining))
            if not isinstance(chunk, bytes) or not chunk:
                raise OSError("destination stream suffix could not be read")
            snapshot.write(chunk)
            remaining -= len(chunk)
        snapshot.seek(0)
        stream.seek(start)
        return snapshot, original_end
    except (AttributeError, OSError, TypeError, ValueError):
        snapshot.close()
        with suppress(Exception):
            stream.seek(start)
        return None


def _restore_stream_tail(stream, start: int, original_end: int, snapshot: BinaryIO) -> None:
    stream.seek(start)
    snapshot.seek(0)
    _copy_stream(snapshot, cast(BinaryIO, stream))
    stream.truncate(original_end)
    stream.seek(start)


def _copy_stream(source: BinaryIO, destination: BinaryIO) -> None:
    while chunk := source.read(_STREAM_COPY_BYTES):
        offset = 0
        while offset < len(chunk):
            written = destination.write(chunk[offset:])
            if not isinstance(written, int) or isinstance(written, bool) or written <= 0:
                raise OSError("destination stream returned an invalid write count")
            if written > len(chunk) - offset:
                raise OSError("destination stream returned an invalid write count")
            offset += written
