"""Document-protection awareness (paper-docx).

Word's Restrict Editing pane writes `w:documentProtection` into
word/settings.xml; Word then blocks or restricts editing. This package's own
mutating APIs check that setting and refuse with |DocumentProtectedError|
rather than silently editing a locked template — the same fail-loudly
principle applied elsewhere in this fork. Upstream python-docx APIs are untouched
(strict superset).

Protection is ADVISORY, not security: the setting is plain XML anyone can
remove. The sanctioned override is one explicit, document-level
acknowledgment — `acknowledge_protection(document)` — after which this
package's APIs treat the document as unlocked for the life of the open
package. There is deliberately NO protection-stripping verb: the setting is
reported by `protection_status` and never removed.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Optional

from docx._guard import check_install
from docx._transaction import rollback_on_error
from docx.errors import DocumentProtectedError
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from docx.document import Document

check_install()

_ACK_ATTR = "_paper_protection_acknowledged"
_TRUTHY = ("1", "true", "on")
_EDIT_MODES = ("readOnly", "comments", "forms", "trackedChanges")
_DOCUMENT_PROTECTION_SUCCESSORS = (
    "w:autoFormatOverride",
    "w:styleLockTheme",
    "w:styleLockQFSet",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:showEnvelope",
    "w:summaryLength",
    "w:clickAndTypeStyle",
    "w:defaultTableStyle",
    "w:evenAndOddHeaders",
    "w:updateFields",
    "w:compat",
)


def _is_on(value: Optional[str]) -> bool:
    return (value or "").lower() in _TRUTHY


@dataclass(frozen=True)
class ProtectionStatus:
    """What `w:documentProtection` declares, read-only.

    `edit` is the raw `w:edit` token ("readOnly", "forms", "comments",
    "trackedChanges", or unrestricted "none") or None when no edit restriction is declared;
    `formatting` reflects the independent format restriction. `enforced`
    reports active enforcement of either kind; `acknowledged` is this
    package's in-memory override flag (never persisted).
    """

    edit: Optional[str]
    enforced: bool
    acknowledged: bool
    formatting: bool = False

    @property
    def blocks_paper_edits(self) -> bool:
        return self.enforced and not self.acknowledged

    def to_dict(self) -> dict:
        return {
            "edit": self.edit,
            "formatting": self.formatting,
            "enforced": self.enforced,
            "acknowledged": self.acknowledged,
        }


def _package_of(obj):
    """The OPC package behind a |Document| or any Parented proxy/part."""
    part = getattr(obj, "part", None) or getattr(obj, "_part", None)
    if part is None:
        raise TypeError(f"cannot resolve a package from {type(obj).__name__}")
    return part.package


def _protection_element(package):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        document_part = package.main_document_part
        settings_part = document_part.part_related_by(RT.SETTINGS)
    except (KeyError, AttributeError):
        return None
    return settings_part.element.find(qn("w:documentProtection"))


def protection_status(document: "Document") -> ProtectionStatus:
    """The document's `w:documentProtection` state (report-only)."""
    package = _package_of(document)
    element = _protection_element(package)
    if element is None:
        return ProtectionStatus(
            edit=None, enforced=False, acknowledged=False, formatting=False
        )
    edit = element.get(qn("w:edit"))
    formatting = _is_on(element.get(qn("w:formatting")))
    enforcement = _is_on(element.get(qn("w:enforcement")))
    return ProtectionStatus(
        edit=edit,
        enforced=(edit not in (None, "none") or formatting) and enforcement,
        acknowledged=bool(getattr(package, _ACK_ATTR, False)),
        formatting=formatting,
    )


def acknowledge_protection(document: "Document") -> ProtectionStatus:
    """Explicitly override protection for this open package (in-memory only).

    The single sanctioned override affordance: after this call, paper-docx
    mutating APIs proceed on this document despite `w:documentProtection`.
    Nothing is written to the file — the protection setting itself is never
    touched. Returns the status being overridden so callers can log it.
    """
    package = _package_of(document)
    status = protection_status(document)
    setattr(package, _ACK_ATTR, True)
    return status


def set_protection(document: "Document", *, edit: str) -> ProtectionStatus:
    """Turn Restrict Editing on for delivery (`w:documentProtection`).

    `edit` is a Word token: `readOnly`, `comments`, `forms`, or
    `trackedChanges`. This writes the setting; it does not strip one.
    Paper mutators then refuse until `acknowledge_protection`.
    """
    if edit not in _EDIT_MODES:
        raise ValueError(f"edit must be one of {_EDIT_MODES}, got {edit!r}")
    with rollback_on_error(document):
        settings = document.settings.element
        element = settings.find(qn("w:documentProtection"))
        if element is None:
            element = OxmlElement("w:documentProtection")
            settings.insert_element_before(element, *_DOCUMENT_PROTECTION_SUCCESSORS)
        element.set(qn("w:edit"), edit)
        element.set(qn("w:enforcement"), "1")
    setattr(_package_of(document), _ACK_ATTR, False)
    return protection_status(document)


def _refuse_if_protected(obj, operation: str) -> None:
    """Typed refusal when `obj`'s document enforces edit/format protection.

    `obj` is whatever the mutating API has in hand: a |Document|, or any
    proxy/part with a `.part` (Table, Paragraph, ...). Called BEFORE any
    mutation (refusal atomicity).
    """
    package = _package_of(obj)
    if getattr(package, _ACK_ATTR, False):
        return
    element = _protection_element(package)
    if element is None:
        return
    edit = element.get(qn("w:edit"))
    formatting = _is_on(element.get(qn("w:formatting")))
    enforcement = _is_on(element.get(qn("w:enforcement")))
    if (edit in (None, "none") and not formatting) or not enforcement:
        return
    restriction = (
        f"{edit!r} editing"
        if edit not in (None, "none")
        else "formatting-only"
    )
    raise DocumentProtectedError(
        f"cannot {operation}: this document enforces {restriction}"
        " protection (w:documentProtection). Protection is advisory, not"
        " security — if editing is intended, call"
        " docx.protection.acknowledge_protection(document) first; paper-docx"
        " never removes the protection setting itself"
    )
