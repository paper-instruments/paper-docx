
.. _paper_composition_api:

Cross-document composition
==========================

*paper-docx addition.* Copy formatted content between documents while
reconciling styles (use destination styles or import renamed definitions),
remapping numbering, recreating media and hyperlinks, and renaming bookmarks
with cross-reference remapping. The returned |CompositionReport| declares every
part the operation touched.

Range endpoints
---------------

``start_anchor`` and ``end_anchor`` address top-level source body blocks. Both
are included by default. With an end anchor, the four combinations are:

.. list-table::
   :header-rows: 1
   :widths: 25 25 50

   * - ``include_start``
     - ``include_end``
     - Selected range
   * - ``True``
     - ``True``
     - Start through end, including both blocks.
   * - ``False``
     - ``True``
     - The block after start through end.
   * - ``True``
     - ``False``
     - Start through the block before end.
   * - ``False``
     - ``False``
     - Only blocks strictly between the anchors.

With no ``end_anchor``, ``count`` selects exactly that many blocks beginning
at the start block, or at the next block when ``include_start=False``.
``include_end=False`` therefore requires an end anchor and otherwise raises
:exc:`ValueError`. When an end anchor is present, ``count`` must still be at
least one but does not limit the anchor-bounded range. An exclusion that leaves
no blocks raises |TargetNotFoundError| before the destination changes.

For example, a source containing ``Opening boundary``, two content blocks, and
``Closing boundary`` can copy only its interior without first discovering the
interior text:

.. code-block:: python

   insert_blocks_from(
       destination,
       source,
       "Opening boundary",
       end_anchor="Closing boundary",
       include_start=False,
       include_end=False,
       anchor="Insert after this paragraph",
   )

Top-level content controls count as one logical block. The selected logical
endpoints are translated to the existing physical body slice, so unsupported
children between them still reach composition preflight and refuse rather
than disappearing silently.

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
