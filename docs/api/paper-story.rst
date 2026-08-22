
.. _paper_story_api:

Story traversal
===============

*paper-docx addition.* Traverse every story and region. Where
``Document.paragraphs`` sees only the body and skips tracked insertions,
content controls, text boxes, and notes, these functions walk every story part
(body, headers, footers, footnotes, endnotes, comments) and every region that
standard traversal misses. A chosen *view* selects ``"current"``,
``"original"``, or ``"all"``. Blocks are live, owner-bound targets; separate
:class:`docx.story.BlockLocator` values carry portable exact evidence. Generic |Anchor| values
are inert historical location evidence, and |Outline| reports regions it
could not read.

.. currentmodule:: docx.story


.. autofunction:: story_parts

.. autofunction:: iter_blocks

.. autofunction:: outline


|Outline| objects
-----------------

.. autoclass:: Outline()
   :members:
   :undoc-members:
   :member-order: bysource


|Block| objects
---------------

.. autoclass:: Block()
   :members:
   :undoc-members:
   :member-order: bysource

Story traversal returns live blocks tied to the exact OOXML element and
document that produced them. A live block follows that element when preceding
blocks change its index, but becomes stale when the element is detached or
replaced. ``Block.to_dict()`` is only an inspection snapshot and never
serializes live identity. Outline schema version 3 keeps the historical
``anchor`` value but labels it with
``anchor_role="legacy_inert_location_evidence"``; the separate ``locator``
payload is the portable form.


``BlockLocator`` objects
------------------------

.. autoclass:: BlockLocator()
   :members:
   :undoc-members:
   :member-order: bysource

``block.locator`` is the portable form. Its strict ``paper_block_locator``
version-1 payload records story, capture view, kind, exact case/whitespace/
Unicode-sensitive content, table cell topology, structural facts, immediate
previous/next evidence (including start/end boundaries), an index hint, and an
optional pre-existing Word paragraph ID. ``BlockLocator.from_dict()`` accepts
only that complete current schema; it never upgrades a three-field |Anchor|,
normalizes text, resolves a document, or creates Word IDs.

Resolution scans every block in the recorded story and view. Zero complete
matches are stale, one succeeds, and multiple complete matches are ambiguous.
The index and Word ID are supporting evidence only. Adjacent edits,
third-party rewrites, or copied repeated regions can therefore make a locator
stale or ambiguous; reacquire a live block or a new locator in that case.


|TableShape| objects
--------------------

.. autoclass:: TableShape()
   :members:
   :undoc-members:
   :member-order: bysource


|Anchor| objects
----------------

.. autoclass:: Anchor()
   :members:
   :undoc-members:
   :member-order: bysource

``Anchor`` remains readable and serializable for search/revision location
evidence, but it cannot authorize a block mutation.
