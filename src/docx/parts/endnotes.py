"""|EndnotesPart|, the story part containing a document's endnotes."""

from __future__ import annotations

from docx.parts.story import StoryPart


class EndnotesPart(StoryPart):
    """Proxy for the endnotes part (`word/endnotes.xml`) of a document.

    Registered for its content type so endnote content loads as a live XML
    part (visible to `docx.story` traversal) instead of an opaque blob.
    paper-docx v0 reads endnotes; it does not create the part.
    """
