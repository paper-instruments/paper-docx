
.. _paper_blocks_api:

Block operations
================

*paper-docx addition.* Make clause-level edits relative to a content anchor.
Insert, delete, or replace whole paragraphs, plainly or as a tracked redline
that stamps paragraph marks so Word accepts or rejects them exactly.
``insert_blocks_after`` takes typed |RichParagraph|/|ListBlock|/|TableBlock|
blocks.

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
