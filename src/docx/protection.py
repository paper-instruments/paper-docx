"""Document-protection awareness (paper-docx).

Word's Restrict Editing pane writes `w:documentProtection` into
word/settings.xml; Word then blocks or restricts editing. This package's own
mutating APIs check that setting and refuse with |DocumentProtectedError|
rather than silently editing a locked template — the same fail-loudly
principle applied elsewhere in this fork. The refusal follows Word's own rules,
which are per-mode AND per-operation-class (see `_PROTECTION_MATRIX`): a
comments-only restriction blocks body edits but still permits commenting.
Upstream python-docx APIs are untouched (strict superset).

Protection is ADVISORY, not security: the setting is plain XML anyone can
remove. The sanctioned override is one explicit, document-level
acknowledgment — `acknowledge_protection(document)` — after which this
package's APIs treat the document as unlocked for the life of the open
package. There is deliberately NO protection-stripping verb: the setting is
reported by `protection_status` and never removed.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Dict, Optional

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

# --- Operation classes -------------------------------------------------------
#
# Word's Restrict Editing rules are per-mode AND per-operation-class: `comments`
# exists precisely to permit commenting while forbidding body edits, and `forms`
# exists to permit filling a form field while locking everything around it. The
# gate therefore needs to know what KIND of operation is being attempted, not
# just that protection is enforced.
OP_COMMENT = "comment"
OP_FORM_FIELD = "form-field-value"
OP_BODY = "body-content"
OPERATION_CLASSES = (OP_COMMENT, OP_FORM_FIELD, OP_BODY)

# The formatting-only row of the matrix: `w:formatting="1"` with no `w:edit`
# restriction (or `w:edit="none"`). The key is a sentinel object rather than a
# string so that no `w:edit` token can ever spell it: a document declaring
# `w:edit="formatting-only"` is an unrecognised restriction and must refuse
# every class, not land on this row. The noun the refusal message uses is a
# separate name for the same reason in reverse — rewording the message must not
# change which matrix row is looked up.
_FORMATTING_ONLY = object()
_FORMATTING_ONLY_NOUN = "formatting-only"

# True = permit, False = refuse. Cells marked (measured) were verified in Word
# for Mac on 2026-08-18 (verdict set `2-refused-operations`); see
# verifying-against-word/WORD-VERDICTS.md, "The protection gate".
#
# TWO CELLS ARE UNMEASURED and ship PERMISSIVE: body content under
# `trackedChanges` and under formatting-only. Word has not been asked about
# either. A refusal with no Word verdict behind it is an unjustified
# regression, and this gate is a policy mirror of Word's UI rather than a
# corruption guard, so permitting cannot produce a file Word refuses. Both are
# pinned by a test naming them unmeasured, so measuring one later changes a
# test deliberately rather than silently.
_PROTECTION_MATRIX: Dict[object, Dict[str, bool]] = {
    #                    comment          form-field value  body content
    "readOnly": {
        OP_COMMENT: False,  # measured: blocked
        OP_FORM_FIELD: False,  # measured: blocked
        OP_BODY: False,  # measured: blocked
    },
    "comments": {
        OP_COMMENT: True,  # measured: ALLOWED
        OP_FORM_FIELD: False,  # unmeasured; the mode is not about form fields
        OP_BODY: False,  # measured: blocked
    },
    "trackedChanges": {
        OP_COMMENT: True,  # measured: ALLOWED
        OP_FORM_FIELD: False,  # unmeasured; the mode is not about form fields
        OP_BODY: True,  # UNMEASURED — ships permissive, see note above
    },
    "forms": {
        OP_COMMENT: False,  # measured: blocked
        OP_FORM_FIELD: True,  # measured: ALLOWED
        OP_BODY: False,  # measured: blocked
    },
    _FORMATTING_ONLY: {
        OP_COMMENT: True,  # measured: ALLOWED
        OP_FORM_FIELD: True,  # unmeasured; no content restriction is declared
        OP_BODY: True,  # UNMEASURED — ships permissive, see note above
    },
}
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
    "w:bookFoldRevPrinting",
    "w:bookFoldPrinting",
    "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing",
    "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery",
    "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin",
    "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin",
    "w:doNotShadeFormData",
    "w:noPunctuationKerning",
    "w:characterSpacingControl",
    "w:printTwoOnOne",
    "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter",
    "w:noLineBreaksBefore",
    "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema",
    "w:saveInvalidXml",
    "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText",
    "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly",
    "w:useXSLTWhenSaving",
    "w:saveThroughXslt",
    "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace",
    "w:updateFields",
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
    "w:rsids",
    "m:mathPr",
    "w:attachedSchema",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures",
    "w:forceUpgrade",
    "w:captions",
    "w:readModeInkLockDown",
    "w:smartTagType",
    "sl:schemaLibrary",
    "w:shapeDefaults",
    "w:doNotEmbedSmartTags",
    "w:decimalSymbol",
    "w:listSeparator",
)


# `set_protection` accepts exactly the `w:edit` tokens the matrix has a row
# for, so the two cannot drift apart. The formatting-only row is keyed by a
# sentinel rather than a string, so it drops out here.
_EDIT_MODES = tuple(mode for mode in _PROTECTION_MATRIX if isinstance(mode, str))


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
        """True when protection is enforced and has not been acknowledged.

        Not a prediction of whether a call will refuse: it ignores the operation class, so it
        reports True under trackedChanges or formatting-only protection where the gate still
        permits body edits and comments. Call `acknowledge_protection` to proceed deliberately.
        """
        return self.enforced and not self.acknowledged

    def to_dict(self) -> dict:
        return {
            "edit": self.edit,
            "formatting": self.formatting,
            "enforced": self.enforced,
            "acknowledged": self.acknowledged,
        }


def _package_of(obj: object):
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


def _refuse_if_protected(
    obj: object, operation: str, *, operation_class: str = OP_BODY
) -> None:
    """Typed refusal when `obj`'s document enforces edit/format protection.

    `obj` is whatever the mutating API has in hand: a |Document|, or any
    proxy/part with a `.part` (Table, Paragraph, ...). Called BEFORE any
    mutation (refusal atomicity).

    `operation_class` says what KIND of edit this is — one of `OP_COMMENT`,
    `OP_FORM_FIELD`, `OP_BODY` — and is consulted against the declared
    restriction mode via `_PROTECTION_MATRIX`. It defaults to `OP_BODY`, the
    strictest class, so a call site added without a class keeps the
    conservative behaviour rather than silently becoming permissive.
    `operation` is unchanged: the class decides *whether* to refuse, the
    operation string still says *what* was refused.
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
    mode: object = edit if edit not in (None, "none") else _FORMATTING_ONLY
    # An unrecognised `w:edit` token has no row: an unknown restriction is not
    # a licence, so every class refuses.
    if _PROTECTION_MATRIX.get(mode, {}).get(operation_class):
        return
    restriction = (
        _FORMATTING_ONLY_NOUN if mode is _FORMATTING_ONLY else f"{edit!r} editing"
    )
    raise DocumentProtectedError(
        f"cannot {operation}: this document is under Restrict Editing"
        f" ({restriction} protection, w:documentProtection), and Word's own"
        " UI would not permit this edit either. The document itself is fine —"
        " this is a policy restriction, not a defect in the file. Protection"
        " is advisory, not security: if the edit is intended, call"
        " docx.protection.acknowledge_protection(document) first; paper-docx"
        " never removes the protection setting itself"
    )
