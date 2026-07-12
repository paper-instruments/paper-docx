"""Verify that the frozen ``docx`` import belongs to ``paper-docx``."""

from __future__ import annotations

import base64
import csv
import hashlib
import hmac
import importlib
import sys
from importlib.metadata import Distribution, PackageNotFoundError, distribution
from io import StringIO
from pathlib import Path, PurePosixPath
from typing import Iterable, Optional, Tuple


class DoctorError(RuntimeError):
    """The installed ``docx`` package cannot be trusted as ``paper-docx``."""


_REMEDY = (
    "python -m pip uninstall -y python-docx paper-docx && "
    "python -m pip install --force-reinstall paper-docx"
)


def verify_install() -> str:
    """Verify distribution ownership, installed bytes, and the fork sentinel.

    Returns the installed ``paper-docx`` version. Raises :class:`DoctorError`
    without importing ``docx`` until its wheel-owned files have been checked.
    """
    paper = _installed_distribution("paper-docx")
    upstream = _installed_distribution("python-docx")
    if paper is None:
        raise DoctorError("paper-docx distribution metadata is missing")
    if upstream is not None:
        raise DoctorError(
            "paper-docx and python-docx are both installed and own the same "
            "docx package"
        )

    docx_root = _verify_docx_record(paper)

    try:
        docx = importlib.import_module("docx")
    except Exception as exc:
        raise DoctorError(f"docx cannot be imported: {exc}") from exc

    _verify_docx_import_path(docx, docx_root)

    sentinel = getattr(docx, "__paper_version__", None)
    if sentinel is None:
        raise DoctorError("docx.__paper_version__ is missing")
    if sentinel != paper.version:
        raise DoctorError(
            "docx.__paper_version__ does not match the installed paper-docx "
            f"version ({sentinel!r} != {paper.version!r})"
        )
    return paper.version


def main() -> int:
    """Console entry point for ``paper-docx-doctor``."""
    try:
        version = verify_install()
    except DoctorError as exc:
        print(f"paper-docx-doctor: FAIL: {exc}", file=sys.stderr)
        print(f"Remedy: {_REMEDY}", file=sys.stderr)
        return 1
    print(f"paper-docx-doctor: OK (paper-docx {version})")
    return 0


def _installed_distribution(name: str) -> Optional[Distribution]:
    try:
        return distribution(name)
    except PackageNotFoundError:
        return None


def _verify_docx_record(dist: Distribution) -> Path:
    record = dist.read_text("RECORD")
    if record is None:
        raise DoctorError("paper-docx RECORD is missing")

    entries = tuple(
        (relative_path, hash_spec)
        for relative_path, hash_spec in _docx_record_entries(record)
        if hash_spec
    )
    if not entries:
        raise DoctorError("paper-docx RECORD has no hashed docx package files")

    for relative_path, hash_spec in entries:
        path = Path(dist.locate_file(relative_path))
        if not path.is_file():
            raise DoctorError(f"paper-docx file is missing: {relative_path}")
        algorithm, expected = _parse_hash(hash_spec, relative_path)
        try:
            actual = _file_digest(path, algorithm)
        except OSError as exc:
            raise DoctorError(
                f"paper-docx file cannot be read: {relative_path}"
            ) from exc
        if not hmac.compare_digest(actual, expected):
            raise DoctorError(f"paper-docx file hash mismatch: {relative_path}")

    try:
        return Path(dist.locate_file(PurePosixPath("docx"))).resolve(strict=True)
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        raise DoctorError("paper-docx owned docx root cannot be resolved") from exc


def _verify_docx_import_path(docx: object, docx_root: Path) -> None:
    module_file = getattr(docx, "__file__", None)
    if not module_file:
        raise DoctorError("docx.__file__ is missing")

    try:
        imported_path = Path(module_file).resolve(strict=True)
    except (OSError, RuntimeError, TypeError, ValueError) as exc:
        raise DoctorError("docx.__file__ cannot be resolved") from exc

    try:
        imported_path.relative_to(docx_root)
    except ValueError:
        raise DoctorError(
            "docx.__file__ is outside the paper-docx owned docx root "
            f"({imported_path} is not within {docx_root})"
        ) from None


def _docx_record_entries(record: str) -> Iterable[Tuple[PurePosixPath, str]]:
    try:
        for row in csv.reader(StringIO(record)):
            if len(row) != 3:
                raise DoctorError("paper-docx RECORD contains a malformed row")
            raw_path, hash_spec, _size = row
            path = PurePosixPath(raw_path)
            if not path.parts or path.parts[0] != "docx":
                continue
            if path.is_absolute() or ".." in path.parts:
                raise DoctorError(
                    f"paper-docx RECORD contains an unsafe path: {raw_path}"
                )
            yield path, hash_spec
    except csv.Error as exc:
        raise DoctorError("paper-docx RECORD contains malformed CSV") from exc


def _parse_hash(hash_spec: str, relative_path: PurePosixPath) -> Tuple[str, str]:
    try:
        algorithm, expected = hash_spec.split("=", 1)
        hashlib.new(algorithm)
    except (TypeError, ValueError):
        raise DoctorError(
            f"paper-docx RECORD has an invalid hash for {relative_path}"
        ) from None
    if not expected:
        raise DoctorError(
            f"paper-docx RECORD has an invalid hash for {relative_path}"
        )
    return algorithm, expected.rstrip("=")


def _file_digest(path: Path, algorithm: str) -> str:
    digest = hashlib.new(algorithm)
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return base64.urlsafe_b64encode(digest.digest()).rstrip(b"=").decode("ascii")


__all__ = ["DoctorError", "main", "verify_install"]
