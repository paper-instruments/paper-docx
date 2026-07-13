"""Text normalization used for matching and anchor hashing.

The table is ported exactly from the battle-tested reference helpers
(`docx_helpers/text.py`): typographic characters agents and humans type
interchangeably fold to ASCII equivalents, whitespace runs collapse, and the
result is casefolded. Normalization exists ONLY for search/matching and
anchor hashing — document edits always preserve original characters and
whitespace verbatim.

`docx.search.normalize_text` is the public name (re-exported there).
"""

from __future__ import annotations

import re

#: char -> replacement. Smart quotes, dashes and minus, exotic spaces,
#: soft hyphen (deleted), tab and CR.
NORMALIZE_CHARS = str.maketrans(
    {
        "‘": "'",  # left single quotation mark
        "’": "'",  # right single quotation mark
        "“": '"',  # left double quotation mark
        "”": '"',  # right double quotation mark
        "–": "-",  # en dash
        "—": "-",  # em dash
        "−": "-",  # minus sign
        " ": " ",  # no-break space
        " ": " ",  # figure space
        " ": " ",  # narrow no-break space
        " ": " ",  # thin space
        "­": "",  # soft hyphen
        "\t": " ",
        "\r": "\n",
    }
)

_WHITESPACE_RUN = re.compile(r"\s+")


def normalize_text(value: str) -> str:
    """`value` normalized for matching: folded punctuation, collapsed
    whitespace, casefolded. Never applied to document content on write.

    ANY Unicode whitespace collapses to a single ASCII space (`\\s+`, not just
    the spaces in the table) so needles and document text normalize
    identically no matter which exotic space either side carries.
    """
    value = value.translate(NORMALIZE_CHARS)
    value = _WHITESPACE_RUN.sub(" ", value)
    return value.casefold()
