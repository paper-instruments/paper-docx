
.. _paper_composition_api:

Cross-document composition
==========================

*paper-docx addition.* Copy formatted content between documents while
reconciling styles (use destination styles or import renamed definitions),
remapping numbering, recreating media and hyperlinks, and renaming bookmarks
with cross-reference remapping. The returned |CompositionReport| declares every
part the operation touched.

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
