"""|FootnotesPart|, the story part containing a document's footnotes."""

from __future__ import annotations

from docx.parts.story import StoryPart


class FootnotesPart(StoryPart):
    """Proxy for the footnotes part (`word/footnotes.xml`) of a document.

    Registered for its content type so footnote content loads as a live XML
    part (visible to `docx.story` traversal) instead of an opaque blob.
    paper-docx v0 reads footnotes; it does not create the part.
    """
