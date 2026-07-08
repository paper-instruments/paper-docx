
.. _paper_composition_api:

Cross-document composition
==========================

Copy formatted content between documents without corruption: style
reconciliation (adopt the house look or import renamed definitions),
numbering remap, media and hyperlink recreation, and bookmark rename
with cross-reference remap. The returned |CompositionReport| declares
every part the operation touched.

.. currentmodule:: docx.composition


.. autofunction:: insert_blocks_from

.. autofunction:: append_document


|CompositionReport| objects
---------------------------

.. autoclass:: CompositionReport()
   :members:
   :undoc-members:
   :member-order: bysource


|CompositionFinding| objects
----------------------------

.. autoclass:: CompositionFinding()
   :members:
   :undoc-members:
   :member-order: bysource
