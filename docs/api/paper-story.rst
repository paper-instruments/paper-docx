
.. _paper_story_api:

Story traversal
===============

*paper-docx addition.* Traverse every story and region. Where
``Document.paragraphs`` sees only the body and skips tracked insertions,
content controls, text boxes, and notes, these functions walk every story part
(body, headers, footers, footnotes, endnotes, comments) and every region that
standard traversal misses. A chosen *view* selects ``"current"``,
``"original"``, or ``"all"``. Blocks are live, owner-bound targets; separate
|Anchor| values are inert historical location evidence, and |Outline| reports
regions it could not read.

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
``anchor_role="legacy_inert_location_evidence"``. Neither ``Block.to_dict()``
nor an |Anchor| can be used to reconstruct mutation authority. Reacquire a
fresh live block in a later session.

Captured view and live identity are separate. Blocks from ``"original"`` and
``"all"`` remain valid inspection values, but block mutation APIs require a
target reacquired from ``view="current"``. A historical projection is not
stale merely because it cannot authorize a mutation.


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
