"""LibreOffice independent-loader smoke oracle.

A document "passes" when headless LibreOffice can open and convert it to PDF
with a zero exit code. This is an openability gate, not a rendering check.

Tests that use this helper carry the `lo_smoke` marker; `conftest.py` skips
them automatically when no `soffice` binary is available.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Tuple


def soffice_binary() -> Optional[str]:
    """Path of the soffice/libreoffice binary, or None when not installed."""
    return shutil.which("soffice") or shutil.which("libreoffice")


def libreoffice_available() -> bool:
    return soffice_binary() is not None


def libreoffice_converts(path: Path, *, timeout: float = 90.0) -> Tuple[bool, str]:
    """(ok, diagnostic) for a headless LibreOffice PDF conversion of `path`.

    Uses an isolated, throwaway LibreOffice user profile so runs cannot clash
    with each other or with the developer's own LibreOffice.
    """
    binary = soffice_binary()
    if binary is None:
        raise RuntimeError("LibreOffice is not installed; gate with libreoffice_available()")
    with tempfile.TemporaryDirectory() as tmp:
        profile_dir = Path(tmp) / "lo-profile"
        completed = subprocess.run(
            [
                binary,
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmp,
                str(path),
            ],
            text=True,
            capture_output=True,
            timeout=timeout,
            check=False,
        )
        produced_pdf = any(Path(tmp).glob("*.pdf"))
    ok = completed.returncode == 0 and produced_pdf
    diagnostic = f"rc={completed.returncode} pdf={produced_pdf} stderr={completed.stderr.strip()}"
    return ok, diagnostic


def assert_libreoffice_opens(path: Path, *, timeout: float = 90.0) -> None:
    ok, diagnostic = libreoffice_converts(path, timeout=timeout)
    assert ok, f"LibreOffice failed to open {path.name}: {diagnostic}"
