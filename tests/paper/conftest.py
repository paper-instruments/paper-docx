"""pytest configuration and shared fixtures for the paper test suite."""

from __future__ import annotations

import pytest

from .harness.clock import FrozenClock
from .harness.lo import libreoffice_available
from .harness.paths import FIXTURES_DIR


def pytest_configure(config: pytest.Config) -> None:
    config.addinivalue_line(
        "markers",
        "lo_smoke: independent-loader smoke test requiring a LibreOffice (soffice) binary;"
        " skipped automatically when none is installed",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if libreoffice_available():
        return
    skip_lo = pytest.mark.skip(reason="LibreOffice (soffice) is not installed")
    for item in items:
        if "lo_smoke" in item.keywords:
            item.add_marker(skip_lo)


@pytest.fixture
def fixtures_dir():
    return FIXTURES_DIR


@pytest.fixture
def frozen_clock() -> FrozenClock:
    return FrozenClock()
