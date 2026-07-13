"""Small relationship helpers used by OPC validation."""

from lxml import etree

TRANSITIONAL_OFFICE_RELATIONSHIPS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
)
STRICT_OFFICE_RELATIONSHIPS = "http://purl.oclc.org/ooxml/officeDocument/relationships/"


def relationship_type_variants(reltype: str) -> tuple[str, ...]:
    """Return transitional and strict spellings for an Office relationship."""
    if reltype.startswith(TRANSITIONAL_OFFICE_RELATIONSHIPS):
        suffix = reltype[len(TRANSITIONAL_OFFICE_RELATIONSHIPS) :]
        return reltype, STRICT_OFFICE_RELATIONSHIPS + suffix
    if reltype.startswith(STRICT_OFFICE_RELATIONSHIPS):
        suffix = reltype[len(STRICT_OFFICE_RELATIONSHIPS) :]
        return TRANSITIONAL_OFFICE_RELATIONSHIPS + suffix, reltype
    return (reltype,)


def is_relationship_type(actual: str, expected: str) -> bool:
    return actual in relationship_type_variants(expected)


def is_xml_id(value: object) -> bool:
    """Whether ``value`` has the XML Schema ID/NCName lexical form."""
    if not isinstance(value, str) or not value:
        return False
    try:
        name = etree.QName(value)
    except ValueError:
        return False
    return name.namespace is None and name.localname == value
