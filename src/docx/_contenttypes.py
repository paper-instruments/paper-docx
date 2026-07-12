"""Small content-type helpers used by OPC validation."""


def media_type(content_type: str) -> str:
    """Return a normalized MIME type without parameters."""
    return content_type.partition(";")[0].strip().casefold()


def content_type_matches(actual: str, expected: str) -> bool:
    return media_type(actual) == media_type(expected)
