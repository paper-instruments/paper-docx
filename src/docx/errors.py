"""Typed refusals for paper-docx safe-editing APIs.

Every mutating paper API is validate-fully-then-mutate: a raised
`PaperRefusal` guarantees the in-memory XML tree and any file on disk are
exactly as they were. Callers can therefore catch "safe refusal" distinctly
from "bug" — programmer errors remain `TypeError`/`ValueError`.
"""

from __future__ import annotations

from docx._guard import check_install

__all__ = [
    "AmbiguousTargetError",
    "BoundaryViolationError",
    "DocumentProtectedError",
    "PackageLimitError",
    "PaperRefusal",
    "RelationshipPolicyError",
    "TargetNotFoundError",
    "UnsupportedStructureError",
]

check_install()


class PaperRefusal(Exception):
    """Base for every safe refusal raised by paper-docx APIs.

    A refused operation mutated nothing, in memory or on disk. The message
    states what was found and why it was unsafe to proceed.
    """


class PackageLimitError(PaperRefusal):
    """A package archive is corrupt, encrypted, or malformed.

    Package reads validate ZIP structure before any XML is parsed or output is
    replaced. This exception therefore represents a safe refusal, including for
    encrypted entries, duplicate/noncanonical member
    names, unsupported entry types, and forged ZIP size metadata.
    """


class AmbiguousTargetError(PaperRefusal):
    """The target specification matches more than one location.

    Disambiguate with `nth=`, `near=`, or `story=` rather than letting the
    library guess.
    """


class TargetNotFoundError(PaperRefusal):
    """No location matches the target specification.

    Also raised when a previously-valid anchor or span has gone stale — the
    underlying content changed since it was captured.
    """


class UnsupportedStructureError(PaperRefusal):
    """The target involves structure this operation does not safely support.

    Examples: text inside a tracked deletion or field instruction; a table
    with merged or nested cells; numbering that would require authoring new
    definitions.
    """


class BoundaryViolationError(PaperRefusal):
    """The operation would cross a structural boundary it must respect.

    Examples: a character-level replacement spanning a paragraph boundary or
    entering/leaving a content control; a block operation selecting
    paragraphs that do not share one parent.
    """


class RelationshipPolicyError(PaperRefusal):
    """The operation would create or modify a package relationship unsafely."""


class DocumentProtectedError(PaperRefusal):
    """The document enforces an editing restriction this operation ignores.

    Word honors `w:documentProtection` (read-only, forms-only, comments-only,
    tracked-changes-enforced); silently editing a locked template reports
    false state. Protection is ADVISORY, not security — after reviewing why
    the document is locked, call
    `docx.protection.acknowledge_protection(document)` to proceed. paper-docx
    never strips the protection setting itself.
    """
