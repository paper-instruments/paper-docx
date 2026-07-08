"""Finalize and scrub — the compliance verbs (paper-docx, v0.11 Phase 3).

`finalize` totally resolves every tracked revision (or refuses, typed,
naming what blocked it). `scrub` removes the reviewing residue a file
carries into the outside world — comments, metadata, the track-changes
switch, optionally RSIDs and hidden text — and returns a |ScrubReport|
itemizing exactly what was removed, so the changed-part budget is
report-matches-diff, never trust-me.

Document protection is REPORTED, never removed (see `docx.protection`).
Both verbs check protection before mutating.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, List, Optional

from docx.errors import UnsupportedStructureError
from docx.oxml.ns import qn
from docx.protection import _refuse_if_protected, protection_status
from docx.revision import _remaining_markup
from docx.story import _story_elements

if TYPE_CHECKING:
    from docx.document import Document

#: relationship-type suffixes of the comment part family (base comments part
#: plus Word's extended/threading/people side-parts)
_COMMENT_RELTYPE_SUFFIXES = (
    "/comments",
    "/commentsExtended",
    "/commentsIds",
    "/commentsExtensible",
    "/people",
)

_COMMENT_ANCHOR_TAGS = (
    qn("w:commentRangeStart"),
    qn("w:commentRangeEnd"),
    qn("w:commentReference"),
)

_RSID_ATTRS = tuple(
    qn(name)
    for name in (
        "w:rsidR",
        "w:rsidRPr",
        "w:rsidRDefault",
        "w:rsidP",
        "w:rsidSect",
        "w:rsidDel",
        "w:rsidTr",
    )
)

#: core-properties string fields cleared by a metadata scrub
_CORE_STRING_FIELDS = (
    "author",
    "category",
    "comments",
    "content_status",
    "identifier",
    "keywords",
    "language",
    "last_modified_by",
    "subject",
    "title",
    "version",
)


@dataclass
class ScrubReport:
    """Everything a `scrub` call removed — itemized, deterministic, goldenable.

    The honest contract: a saved scrubbed file's package diff must be
    explained entirely by this report (report-matches-diff).
    """

    removed_parts: List[str] = field(default_factory=list)
    comment_anchors_removed: int = 0
    metadata_fields_cleared: List[str] = field(default_factory=list)
    track_changes_setting_removed: bool = False
    rsids_element_removed: bool = False
    rsid_attributes_removed: int = 0
    hidden_runs_removed: int = 0
    #: protection is REPORTED here and never removed
    document_protection: Optional[dict] = None

    def to_dict(self) -> dict:
        return {
            "schema": "paper_scrub_report",
            "version": 1,
            "removed_parts": sorted(self.removed_parts),
            "comment_anchors_removed": self.comment_anchors_removed,
            "metadata_fields_cleared": self.metadata_fields_cleared,
            "track_changes_setting_removed": self.track_changes_setting_removed,
            "rsids_element_removed": self.rsids_element_removed,
            "rsid_attributes_removed": self.rsid_attributes_removed,
            "hidden_runs_removed": self.hidden_runs_removed,
            "document_protection": self.document_protection,
        }


def finalize(document: "Document", *, revisions: str = "accept") -> int:
    """Totally resolve every tracked revision; refuse (typed) if anything
    cannot be — never a file that LOOKS final while markup remains.

    Returns the number of revisions resolved. After a successful return, a
    rescan of every story finds zero revision markup of any kind.
    """
    if revisions not in ("accept", "reject"):
        raise ValueError(
            f"revisions must be 'accept' or 'reject', got {revisions!r}"
        )
    _refuse_if_protected(document, "finalize the document")
    snapshot = document.revisions
    resolved = (
        snapshot.accept_all() if revisions == "accept" else snapshot.reject_all()
    )
    leftover = _remaining_markup(document)
    if leftover:
        raise UnsupportedStructureError(
            "finalize resolved the enumerated revisions but revision markup"
            f" remains: {leftover}; the document is NOT final. Resolve the"
            " remainder in Word"
        )
    return resolved


def scrub(
    document: "Document",
    *,
    comments: bool = True,
    metadata: bool = True,
    track_changes_setting: bool = True,
    rsids: bool = False,
    hidden_text: bool = False,
) -> ScrubReport:
    """Remove reviewing residue before a file leaves the building.

    Every target is individually toggleable; the returned |ScrubReport|
    itemizes exactly what was removed. A metadata scrub refuses while
    tracked revisions are pending (their author/date attributions would
    survive it) — `finalize` first. Document protection settings are
    reported, never removed.
    """
    from docx.revision import _fallback_markup

    _refuse_if_protected(document, "scrub the document")
    if metadata and (len(document.revisions) or _fallback_markup(document)):
        raise UnsupportedStructureError(
            "cannot scrub metadata while tracked revisions are pending"
            " (including markup inside mc:AlternateContent fallbacks) —"
            " their author and date attributions would remain in the"
            " document; finalize(revisions=...) first, or pass"
            " metadata=False. Nothing was changed"
        )
    report = ScrubReport()
    status = protection_status(document)
    if status.edit is not None:
        report.document_protection = {
            "edit": status.edit,
            "enforced": status.enforced,
            "note": "reported, never removed (docx.protection)",
        }
    if comments:
        _scrub_comment_parts(document, report)
        _scrub_comment_anchors(document, report)
    if metadata:
        _scrub_metadata(document, report)
    if track_changes_setting:
        _scrub_track_changes_setting(document, report)
    if rsids:
        _scrub_rsids(document, report)
    if hidden_text:
        _scrub_hidden_text(document, report)
    return report


def _scrub_comment_parts(document: "Document", report: ScrubReport) -> None:
    """Drop the comment part family and report EVERY part that leaves the
    package with it — parts reachable only through a dropped part (comment
    media, the dropped part's own .rels file) cascade out of the saved zip,
    and an unexplained removal breaks report-matches-diff."""
    document_part = document.part
    package = document_part.package
    parts_before = {part.partname: part for part in package.iter_parts()}
    dropped_any = False
    for r_id, rel in list(document_part.rels.items()):
        if rel.is_external:
            continue
        if any(rel.reltype.endswith(sfx) for sfx in _COMMENT_RELTYPE_SUFFIXES):
            document_part.drop_rel(r_id)
            dropped_any = True
    if not dropped_any:
        return
    parts_after = {part.partname for part in package.iter_parts()}
    for partname, part in sorted(parts_before.items()):
        if partname in parts_after:
            continue
        name = str(partname).lstrip("/")
        report.removed_parts.append(name)
        if len(getattr(part, "rels", ())):
            directory, _, filename = name.rpartition("/")
            report.removed_parts.append(f"{directory}/_rels/{filename}.rels")


def _scrub_comment_anchors(document: "Document", report: ScrubReport) -> None:
    for _story, root in _story_elements(document):
        for node in list(root.iter(*_COMMENT_ANCHOR_TAGS)):
            parent = node.getparent()
            parent.remove(node)
            report.comment_anchors_removed += 1
            # a run that held only the commentReference is Word residue too
            if parent.tag == qn("w:r") and not any(
                child.tag != qn("w:rPr") for child in parent
            ):
                parent.getparent().remove(parent)


def _scrub_metadata(document: "Document", report: ScrubReport) -> None:
    package = document.part.package
    has_core_part = any(
        not rel.is_external and rel.reltype.endswith("/core-properties")
        for rel in package.rels.values()
    )
    if has_core_part:  # never FABRICATE a core part just to clear it
        core = document.core_properties
        for name in _CORE_STRING_FIELDS:
            if getattr(core, name):
                setattr(core, name, "")
                report.metadata_fields_cleared.append(f"core:{name}")
    for r_id, rel in list(package.rels.items()):
        if rel.is_external:
            continue
        if rel.reltype.endswith("/custom-properties") or rel.reltype.endswith(
            "/thumbnail"
        ):
            report.removed_parts.append(str(rel.target_part.partname).lstrip("/"))
            del package.rels[r_id]
        elif rel.reltype.endswith("/extended-properties"):
            _clear_app_properties(rel.target_part, report)


def _clear_app_properties(app_part, report: ScrubReport) -> None:
    """Blank Company/Manager in docProps/app.xml (kept as a generic part)."""
    from lxml import etree

    try:
        root = etree.fromstring(app_part.blob)
    except etree.XMLSyntaxError:
        return
    changed = False
    for local in ("Company", "Manager"):
        for element in root.iter(f"{{*}}{local}"):
            if element.text:
                element.text = ""
                report.metadata_fields_cleared.append(f"app:{local}")
                changed = True
    if changed:
        app_part._blob = etree.tostring(  # noqa: SLF001 - generic Part storage
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )


def _settings_element(document: "Document"):
    """The settings root, or None — scrubbing must never FABRICATE a
    settings part just to remove things from it."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    try:
        part = document.part.part_related_by(RT.SETTINGS)
    except KeyError:
        return None
    return part.element


def _scrub_track_changes_setting(document: "Document", report: ScrubReport) -> None:
    settings = _settings_element(document)
    if settings is None:
        return
    node = settings.find(qn("w:trackRevisions"))
    if node is not None:
        settings.remove(node)
        report.track_changes_setting_removed = True


def _scrub_rsids(document: "Document", report: ScrubReport) -> None:
    settings = _settings_element(document)
    if settings is not None:
        node = settings.find(qn("w:rsids"))
        if node is not None:
            settings.remove(node)
            report.rsids_element_removed = True
    for _story, root in _story_elements(document):
        for element in root.iter():
            for attr in _RSID_ATTRS:
                if element.get(attr) is not None:
                    del element.attrib[attr]
                    report.rsid_attributes_removed += 1


def _scrub_hidden_text(document: "Document", report: ScrubReport) -> None:
    """Remove runs carrying `w:vanish` (explicit opt-in — this DELETES
    content Word merely hides)."""
    vanish = qn("w:vanish")
    val = qn("w:val")
    for _story, root in _story_elements(document):
        for run in list(root.iter(qn("w:r"))):
            r_pr = run.find(qn("w:rPr"))
            if r_pr is None:
                continue
            node = r_pr.find(vanish)
            if node is None:
                continue
            if (node.get(val) or "true").lower() in ("0", "false", "off"):
                continue
            parent = run.getparent()
            if parent is not None:
                parent.remove(run)
                report.hidden_runs_removed += 1
