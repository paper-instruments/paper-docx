
.. _paper_blocks_api:

Block operations
================

*paper-docx addition.* Make clause-level edits relative to a content anchor.
Insert, delete, or replace whole paragraphs, plainly or as a tracked redline
that stamps paragraph marks so Word accepts or rejects them exactly.
``insert_blocks_after`` takes typed |RichParagraph|/|ListBlock|/|TableBlock|
blocks.

Targets may be an exact string, a live |Span|, a live |Block|, or a current
portable :class:`docx.story.BlockLocator`. Live blocks validate document ownership and exact
element attachment. Locators require one complete exact content/topology and
adjacent-context match; missing evidence raises |TargetNotFoundError| and
multiple matches raise |AmbiguousTargetError|. Tables remain blocks but are
not paragraph anchors, so table blocks and table locators raise
|UnsupportedStructureError|.

Legacy |Anchor| values are inert location evidence and are refused with
migration guidance. Reacquire a live block with ``iter_blocks()``/``outline()``,
pass an exact string or live span, or deserialize a current ``BlockLocator``.
Do not use an old index/hash payload as a mutation target.

.. currentmodule:: docx.blocks


.. autofunction:: insert_section_after

.. autofunction:: insert_blocks_after

.. autofunction:: tracked_delete_paragraphs

.. autofunction:: tracked_replace_paragraphs


|TextRun| objects
-----------------

.. autoclass:: TextRun()
   :members:
   :undoc-members:
   :member-order: bysource


|RichParagraph| objects
-----------------------

.. autoclass:: RichParagraph()
   :members:
   :undoc-members:
   :member-order: bysource


|ListBlock| objects
-------------------

.. autoclass:: ListBlock()
   :members:
   :undoc-members:
   :member-order: bysource


|TableBlock| objects
--------------------

.. autoclass:: TableBlock()
   :members:
   :undoc-members:
   :member-order: bysource


|BlockEditResult| objects
-------------------------

.. autoclass:: BlockEditResult()
   :members:
   :undoc-members:
   :member-order: bysource
