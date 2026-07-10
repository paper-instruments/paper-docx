"""Injectable clock for every paper-docx API that stamps a date.

APIs take `date: datetime | None = None`; `None` resolves through `now()`
here, so tests can freeze time by monkeypatching this module — production
callers never need to touch it.
"""

from __future__ import annotations

import datetime as dt


def now() -> dt.datetime:
    """Current UTC time, truncated to whole seconds (OOXML date precision)."""
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0)
