"""Low-level, read-only API to a serialized Open Packaging Convention (OPC) package."""

from lxml import etree

from docx._contenttypes import content_type_matches
from docx._rels import is_relationship_type, is_xml_id
from docx._zipguard import _parse_content_types
from docx.errors import PackageLimitError
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.oxml import parse_xml
from docx.opc.packuri import PACKAGE_URI, PackURI
from docx.opc.phys_pkg import PhysPkgReader
from docx.opc.shared import CaseInsensitiveDict


class PackageReader:
    """Provides access to the contents of a zip-format OPC package via its
    :attr:`serialized_parts` and :attr:`pkg_srels` attributes."""

    def __init__(self, content_types, pkg_srels, sparts):
        super(PackageReader, self).__init__()
        self._pkg_srels = pkg_srels
        self._sparts = sparts

    @staticmethod
    def from_file(pkg_file):
        """Return a |PackageReader| instance loaded with contents of `pkg_file`."""
        phys_reader = PhysPkgReader(pkg_file)
        try:
            content_types = _ContentTypeMap.from_xml(phys_reader.content_types_xml)
            pkg_srels = PackageReader._srels_for(phys_reader, PACKAGE_URI)
            _validate_package_relationships(pkg_srels)
            sparts = PackageReader._load_serialized_parts(phys_reader, pkg_srels, content_types)
            return PackageReader(content_types, pkg_srels, sparts)
        finally:
            phys_reader.close()

    def iter_sparts(self):
        """Generate a 4-tuple `(partname, content_type, reltype, blob)` for each of the
        serialized parts in the package."""
        for s in self._sparts:
            yield (s.partname, s.content_type, s.reltype, s.blob)

    def iter_srels(self):
        """Generate a 2-tuple `(source_uri, srel)` for each of the relationships in the
        package."""
        for srel in self._pkg_srels:
            yield (PACKAGE_URI, srel)
        for spart in self._sparts:
            for srel in spart.srels:
                yield (spart.partname, srel)

    @staticmethod
    def _load_serialized_parts(phys_reader, pkg_srels, content_types):
        """Return a list of |_SerializedPart| instances corresponding to the parts in
        `phys_reader` accessible by walking the relationship graph starting with
        `pkg_srels`."""
        sparts = []
        part_walker = PackageReader._walk_phys_parts(
            phys_reader, pkg_srels, content_types=content_types
        )
        for partname, blob, reltype, srels in part_walker:
            content_type = content_types[partname]
            spart = _SerializedPart(partname, content_type, reltype, blob, srels)
            sparts.append(spart)
        return tuple(sparts)

    @staticmethod
    def _srels_for(phys_reader, source_uri):
        """Return |_SerializedRelationships| instance populated with relationships for
        source identified by `source_uri`."""
        rels_xml = phys_reader.rels_xml_for(source_uri)
        return _SerializedRelationships.load_from_xml(source_uri.baseURI, rels_xml)

    @staticmethod
    def _walk_phys_parts(phys_reader, srels, visited_partnames=None, content_types=None):
        """Generate a 4-tuple `(partname, blob, reltype, srels)` for each of the parts
        in `phys_reader` by walking the relationship graph rooted at srels."""
        if visited_partnames is None:
            visited_partnames = []
        for srel in srels:
            if srel.is_external:
                continue
            partname = srel.target_partname
            if content_types is not None:
                try:
                    phys_reader.partname_for(partname)
                except KeyError as exc:
                    raise PackageLimitError(
                        f"relationship {srel.rId!r} targets missing package part"
                        f" {partname!s}"
                    ) from exc
                try:
                    content_type = content_types[partname]
                except KeyError as exc:
                    raise PackageLimitError(
                        f"relationship {srel.rId!r} targets {partname!s}, which has no"
                        " declared content type"
                    ) from exc
                _validate_relationship_target_content_type(
                    srel.reltype, content_type, partname
                )
            if str(partname).casefold() in {str(item).casefold() for item in visited_partnames}:
                continue
            visited_partnames.append(partname)
            reltype = srel.reltype
            part_srels = PackageReader._srels_for(phys_reader, partname)
            blob = phys_reader.blob_for(partname)
            yield (partname, blob, reltype, part_srels)
            next_walker = PackageReader._walk_phys_parts(
                phys_reader, part_srels, visited_partnames, content_types
            )
            for partname, blob, reltype, srels in next_walker:
                yield (partname, blob, reltype, srels)


class _ContentTypeMap:
    """Value type providing dictionary semantics for looking up content type by part
    name, e.g. ``content_type = cti['/ppt/presentation.xml']``."""

    def __init__(self):
        super(_ContentTypeMap, self).__init__()
        self._overrides = CaseInsensitiveDict()
        self._defaults = CaseInsensitiveDict()

    def __getitem__(self, partname):
        """Return content type for part identified by `partname`."""
        if not isinstance(partname, PackURI):
            tmpl = "_ContentTypeMap key must be <type 'PackURI'>, got %s"
            raise KeyError(tmpl % type(partname))
        if partname in self._overrides:
            return self._overrides[partname]
        if partname.ext in self._defaults:
            return self._defaults[partname.ext]
        tmpl = "no content type for partname '%s' in [Content_Types].xml"
        raise KeyError(tmpl % partname)

    @staticmethod
    def from_xml(content_types_xml):
        """Return a new |_ContentTypeMap| instance populated with the contents of
        `content_types_xml`."""
        _parse_content_types(content_types_xml)
        types_elm = parse_xml(content_types_xml)
        ct_map = _ContentTypeMap()
        for o in types_elm.overrides:
            ct_map._add_override(o.partname, o.content_type)
        for d in types_elm.defaults:
            ct_map._add_default(d.extension, d.content_type)
        return ct_map

    def _add_default(self, extension, content_type):
        """Add the default mapping of `extension` to `content_type` to this content type
        mapping."""
        self._defaults[extension] = content_type

    def _add_override(self, partname, content_type):
        """Add the default mapping of `partname` to `content_type` to this content type
        mapping."""
        self._overrides[partname] = content_type


class _SerializedPart:
    """Value object for an OPC package part.

    Provides access to the partname, content type, blob, and serialized relationships
    for the part.
    """

    def __init__(self, partname, content_type, reltype, blob, srels):
        super(_SerializedPart, self).__init__()
        self._partname = partname
        self._content_type = content_type
        self._reltype = reltype
        self._blob = blob
        self._srels = srels

    @property
    def partname(self):
        return self._partname

    @property
    def content_type(self):
        return self._content_type

    @property
    def blob(self):
        return self._blob

    @property
    def reltype(self):
        """The referring relationship type of this part."""
        return self._reltype

    @property
    def srels(self):
        return self._srels


class _SerializedRelationship:
    """Value object representing a serialized relationship in an OPC package.

    Serialized, in this case, means any target part is referred to via its partname
    rather than a direct link to an in-memory |Part| object.
    """

    def __init__(self, baseURI, rel_elm):
        super(_SerializedRelationship, self).__init__()
        self._baseURI = baseURI
        self._rId = rel_elm.rId
        self._reltype = rel_elm.reltype
        self._target_mode = rel_elm.target_mode
        self._target_ref = rel_elm.target_ref

    @property
    def is_external(self):
        """True if target_mode is ``RTM.EXTERNAL``"""
        return self._target_mode == RTM.EXTERNAL

    @property
    def reltype(self):
        """Relationship type, like ``RT.OFFICE_DOCUMENT``"""
        return self._reltype

    @property
    def rId(self):
        """Relationship id, like 'rId9', corresponds to the ``Id`` attribute on the
        ``CT_Relationship`` element."""
        return self._rId

    @property
    def target_mode(self):
        """String in ``TargetMode`` attribute of ``CT_Relationship`` element, one of
        ``RTM.INTERNAL`` or ``RTM.EXTERNAL``."""
        return self._target_mode

    @property
    def target_ref(self):
        """String in ``Target`` attribute of ``CT_Relationship`` element, a relative
        part reference for internal target mode or an arbitrary URI, e.g. an HTTP URL,
        for external target mode."""
        return self._target_ref

    @property
    def target_partname(self):
        """|PackURI| instance containing partname targeted by this relationship.

        Raises ``ValueError`` on reference if target_mode is ``'External'``. Use
        :attr:`target_mode` to check before referencing.
        """
        if self.is_external:
            msg = (
                "target_partname attribute on Relationship is undefined w"
                'here TargetMode == "External"'
            )
            raise ValueError(msg)
        # lazy-load _target_partname attribute
        if not hasattr(self, "_target_partname"):
            self._target_partname = PackURI.from_rel_ref(self._baseURI, self.target_ref)
        return self._target_partname


class _SerializedRelationships:
    """Read-only sequence of |_SerializedRelationship| instances corresponding to the
    relationships item XML passed to constructor."""

    def __init__(self):
        super(_SerializedRelationships, self).__init__()
        self._srels = []

    def __iter__(self):
        """Support iteration, e.g. 'for x in srels:'."""
        return self._srels.__iter__()

    @staticmethod
    def load_from_xml(baseURI, rels_item_xml):
        """Return |_SerializedRelationships| instance loaded with the relationships
        contained in `rels_item_xml`.

        Returns an empty collection if `rels_item_xml` is |None|.
        """
        srels = _SerializedRelationships()
        if rels_item_xml is not None:
            if isinstance(rels_item_xml, (bytes, str)):
                _validate_relationship_records(rels_item_xml, baseURI)
            rels_elm = parse_xml(rels_item_xml)
            seen_ids = set()
            for rel_elm in rels_elm.Relationship_lst:
                serialized = _SerializedRelationship(baseURI, rel_elm)
                if isinstance(serialized.rId, str) and not is_xml_id(serialized.rId):
                    raise PackageLimitError(
                        f"relationship Id {serialized.rId!r} is not a valid XML ID"
                    )
                if isinstance(serialized.rId, str) and serialized.rId in seen_ids:
                    raise PackageLimitError(
                        f"relationships for {baseURI!r} contain duplicate Id {serialized.rId!r}"
                    )
                if isinstance(serialized.rId, str):
                    seen_ids.add(serialized.rId)
                if isinstance(serialized.target_mode, str) and serialized.target_mode not in (
                    RTM.INTERNAL,
                    RTM.EXTERNAL,
                ):
                    raise PackageLimitError(
                        f"relationship {serialized.rId!r} has invalid TargetMode"
                    )
                if not serialized.is_external:
                    try:
                        serialized.target_partname
                    except (IndexError, TypeError, ValueError) as exc:
                        raise PackageLimitError(
                            f"relationship {serialized.rId!r} has an invalid target"
                        ) from exc
                srels._srels.append(serialized)
        return srels


def _validate_relationship_records(blob, base_uri) -> None:
    """Check relationship fields before the object model can normalize them."""
    try:
        root = etree.fromstring(blob)
    except etree.XMLSyntaxError as exc:
        raise PackageLimitError(f"relationships for {base_uri!r} are malformed") from exc
    relationships_tag = (
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"
    )
    relationship_tag = (
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    )
    if root.tag != relationships_tag:
        raise PackageLimitError(
            f"relationships for {base_uri!r} have an unexpected root element"
        )
    seen = set()
    for child in root:
        if not isinstance(child.tag, str):
            continue
        if child.tag != relationship_tag:
            raise PackageLimitError(
                f"relationships for {base_uri!r} contain an unexpected element"
            )
        relationship_id = child.get("Id")
        if not is_xml_id(relationship_id):
            raise PackageLimitError(f"relationship Id {relationship_id!r} is not a valid XML ID")
        if relationship_id in seen:
            raise PackageLimitError(
                f"relationships for {base_uri!r} contain duplicate Id {relationship_id!r}"
            )
        seen.add(relationship_id)
        if not child.get("Type") or not child.get("Target"):
            raise PackageLimitError(f"relationship {relationship_id!r} is missing Type or Target")
        if child.get("TargetMode", RTM.INTERNAL) not in (RTM.INTERNAL, RTM.EXTERNAL):
            raise PackageLimitError(f"relationship {relationship_id!r} has invalid TargetMode")


def _validate_package_relationships(relationships) -> None:
    """Reject ambiguous main-part declarations before object construction."""
    office_documents = [
        relationship
        for relationship in relationships
        if is_relationship_type(relationship.reltype, RT.OFFICE_DOCUMENT)
    ]
    if len(office_documents) > 1:
        raise PackageLimitError(
            "package contains multiple officeDocument relationships"
        )
    if office_documents and office_documents[0].is_external:
        raise PackageLimitError("package officeDocument relationship is external")


_OFFICE_DOCUMENT_CONTENT_TYPES = (
    CT.WML_DOCUMENT_MAIN,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "application/vnd.ms-word.document.macroEnabled.main+xml",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    CT.PML_PRESENTATION_MAIN,
    CT.PML_SLIDESHOW_MAIN,
    CT.PML_TEMPLATE_MAIN,
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
    CT.SML_SHEET_MAIN,
    CT.SML_TEMPLATE_MAIN,
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "application/vnd.ms-excel.template.macroEnabled.main+xml",
)
_KNOWN_RELATIONSHIP_CONTENT_TYPES = (
    (RT.CORE_PROPERTIES, (CT.OPC_CORE_PROPERTIES,)),
    (RT.OFFICE_DOCUMENT, _OFFICE_DOCUMENT_CONTENT_TYPES),
    (RT.STYLES, (CT.WML_STYLES, CT.SML_STYLES)),
    (RT.SETTINGS, (CT.WML_SETTINGS,)),
    (RT.NUMBERING, (CT.WML_NUMBERING,)),
    (RT.FONT_TABLE, (CT.WML_FONT_TABLE,)),
    (RT.HEADER, (CT.WML_HEADER,)),
    (RT.FOOTER, (CT.WML_FOOTER,)),
    (RT.FOOTNOTES, (CT.WML_FOOTNOTES,)),
    (RT.ENDNOTES, (CT.WML_ENDNOTES,)),
    (RT.COMMENTS, (CT.WML_COMMENTS, CT.PML_COMMENTS, CT.SML_COMMENTS)),
    (RT.THEME, (CT.OFC_THEME,)),
)


def _validate_relationship_target_content_type(reltype, content_type, partname) -> None:
    """Reject a known role pointing at a part with an incompatible media type."""
    if is_relationship_type(reltype, RT.IMAGE):
        if content_type.partition(";")[0].strip().casefold().startswith("image/"):
            return
        raise PackageLimitError(
            f"image relationship targets {partname!s} with invalid content type {content_type!r}"
        )
    for expected_reltype, expected_types in _KNOWN_RELATIONSHIP_CONTENT_TYPES:
        if not is_relationship_type(reltype, expected_reltype):
            continue
        if any(content_type_matches(content_type, item) for item in expected_types):
            return
        raise PackageLimitError(
            f"relationship type {reltype!r} targets {partname!s} with invalid content type"
            f" {content_type!r}"
        )
