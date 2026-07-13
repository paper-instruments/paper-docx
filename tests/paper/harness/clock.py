"""Frozen clock for anything that stamps dates.

Every paper API that writes a date (`w:date` on revisions, comment
timestamps, ...) takes an injectable clock or an explicit `date=`; tests pass
a `FrozenClock` so output is deterministic.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

#: The fixed instant used across the paper test suite unless a test needs
#: a different one. Arbitrary but memorable; timezone-aware UTC.
PAPER_TEST_INSTANT = dt.datetime(2026, 1, 2, 3, 4, 5, tzinfo=dt.timezone.utc)


@dataclass(frozen=True)
class FrozenClock:
    """A clock whose `now()` always returns the same instant."""

    instant: dt.datetime = PAPER_TEST_INSTANT

    def now(self) -> dt.datetime:
        return self.instant
