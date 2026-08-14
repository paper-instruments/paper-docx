"""Mixed-install check imported by every Paper-only public module."""

from docx._version import assert_distribution_identity


def check_install() -> None:
    """Refuse when paper-docx and python-docx are both installed."""
    assert_distribution_identity()
