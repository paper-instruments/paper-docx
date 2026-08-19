"""Fork identity and distribution-conflict guard."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, distribution

__paper_version__ = "0.2.0"


def assert_distribution_identity() -> None:
    """Fail when both distributions claim the frozen ``docx`` import."""
    try:
        distribution("python-docx")
    except PackageNotFoundError:
        return
    try:
        distribution("paper-docx")
    except PackageNotFoundError:
        return
    raise ImportError(
        "paper-docx and python-docx are both installed, but both own the"
        " same 'docx' import package. Uninstall both distributions, then"
        " install paper-docx in a clean environment"
    )
