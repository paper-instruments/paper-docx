"""Provides a general interface to a `physical` OPC package, such as a zip file."""

import os
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo, is_zipfile

from docx._zipguard import GuardedZipReader, preflight_zip
from docx.errors import PackageLimitError
from docx.opc.exceptions import PackageNotFoundError
from docx.opc.packuri import CONTENT_TYPES_URI

_FIXED_ZIP_DATE = (1980, 1, 1, 0, 0, 0)


class PhysPkgReader:
    """Factory for physical package reader objects."""

    def __new__(cls, pkg_file):
        # if `pkg_file` is a string, treat it as a path
        if isinstance(pkg_file, (str, os.PathLike)):
            path = os.fspath(pkg_file)
            if os.path.isdir(path):
                reader_cls = _DirPkgReader
            elif is_zipfile(path):
                reader_cls = _ZipPkgReader
            else:
                raise PackageNotFoundError("Package not found at '%s'" % path)
        else:  # assume it's a stream and pass it to Zip reader to sort out
            reader_cls = _ZipPkgReader

        return super(PhysPkgReader, cls).__new__(reader_cls)


class PhysPkgWriter:
    """Factory for physical package writer objects."""

    def __new__(cls, pkg_file):
        return super(PhysPkgWriter, cls).__new__(_ZipPkgWriter)


class _DirPkgReader(PhysPkgReader):
    """Implements |PhysPkgReader| interface for an OPC package extracted into a
    directory."""

    def __init__(self, path):
        """`path` is the path to a directory containing an expanded package."""
        super(_DirPkgReader, self).__init__()
        self._path = os.path.abspath(os.fspath(path))
        if os.path.islink(self._path):
            raise PackageLimitError("expanded package root is a symbolic link")
        names = []
        for root, directories, files in os.walk(self._path, followlinks=False):
            directories.sort()
            files.sort()
            for directory in directories:
                if os.path.islink(os.path.join(root, directory)):
                    raise PackageLimitError("expanded package contains a symbolic link")
            for filename in files:
                file_path = os.path.join(root, filename)
                if os.path.islink(file_path):
                    raise PackageLimitError("expanded package contains a symbolic link")
                member = os.path.relpath(file_path, self._path).replace(
                    os.sep, "/"
                )
                names.append(member)
        self._member_name_by_fold = {}
        for name in names:
            folded = name.casefold()
            if folded in self._member_name_by_fold:
                raise PackageLimitError(
                    f"expanded package contains case-ambiguous member name {name!r}"
                )
            self._member_name_by_fold[folded] = name
        self._member_names = tuple(names)

    def blob_for(self, pack_uri):
        """Return contents of file corresponding to `pack_uri` in package directory."""
        actual_name = self._member_name_by_fold[pack_uri.membername.casefold()]
        path = os.path.join(self._path, *actual_name.split("/"))
        with open(path, "rb") as f:
            blob = f.read()
        return blob

    def partname_for(self, pack_uri):
        actual_name = self._member_name_by_fold[pack_uri.membername.casefold()]
        return type(pack_uri)(f"/{actual_name}")

    @property
    def member_names(self):
        return self._member_names

    def close(self):
        """Provides interface consistency with |ZipFileSystem|, but does nothing, a
        directory file system doesn't need closing."""
        pass

    @property
    def content_types_xml(self):
        """Return the `[Content_Types].xml` blob from the package."""
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """Return rels item XML for source with `source_uri`, or None if the item has no
        rels item."""
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except (IOError, KeyError):
            rels_xml = None
        return rels_xml


class _ZipPkgReader(PhysPkgReader):
    """Implements |PhysPkgReader| interface for a zip file OPC package."""

    def __init__(self, pkg_file):
        super(_ZipPkgReader, self).__init__()
        if isinstance(pkg_file, (str, os.PathLike)):
            with open(pkg_file, "rb") as stream:
                self._load(stream)
                # GuardedZipReader caches every member during construction.
                self._zipf.close()
            return
        self._load(pkg_file)

    def _load(self, source):
        """Preflight and cache one already-open package source."""
        try:
            preflight_zip(source)
            self._zipf = ZipFile(source, "r")
            self._guarded_reader = GuardedZipReader(self._zipf)
            self._member_name_by_fold = {
                name.casefold(): name for name in self._guarded_reader.order
            }
        except Exception:
            zip_file = getattr(self, "_zipf", None)
            if zip_file is not None:
                zip_file.close()
            raise

    def blob_for(self, pack_uri):
        """Return blob corresponding to `pack_uri`.

        Raises |ValueError| if no matching member is present in zip archive.
        """
        actual_name = self._member_name_by_fold[pack_uri.membername.casefold()]
        return self._guarded_reader.read(actual_name)

    def partname_for(self, pack_uri):
        actual_name = self._member_name_by_fold[pack_uri.membername.casefold()]
        return type(pack_uri)(f"/{actual_name}")

    @property
    def member_names(self):
        return self._guarded_reader.order

    def close(self):
        """Close the zip archive, releasing any resources it is using."""
        self._zipf.close()

    @property
    def content_types_xml(self):
        """Return the `[Content_Types].xml` blob from the zip package."""
        return self.blob_for(CONTENT_TYPES_URI)

    def rels_xml_for(self, source_uri):
        """Return rels item XML for source with `source_uri` or None if no rels item is
        present."""
        try:
            rels_xml = self.blob_for(source_uri.rels_uri)
        except KeyError:
            rels_xml = None
        return rels_xml


class _ZipPkgWriter(PhysPkgWriter):
    """Implements |PhysPkgWriter| interface for a zip file OPC package."""

    def __init__(self, pkg_file):
        super(_ZipPkgWriter, self).__init__()
        self._zipf = ZipFile(pkg_file, "w", compression=ZIP_DEFLATED)

    def close(self):
        """Close the zip archive, flushing any pending physical writes and releasing any
        resources it's using."""
        self._zipf.close()

    def write(self, pack_uri, blob):
        """Write `blob` to this zip package with the membername corresponding to
        `pack_uri`."""
        info = ZipInfo(pack_uri.membername, date_time=_FIXED_ZIP_DATE)
        info.compress_type = ZIP_DEFLATED
        info.create_system = 3
        info.external_attr = 0o600 << 16
        self._zipf.writestr(info, blob)
