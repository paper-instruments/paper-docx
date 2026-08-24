
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

``start_anchor`` and ``end_anchor`` identify source paragraphs whose containing
top-level body blocks form the range endpoints. A direct body paragraph or a
paragraph in a top-level content control can serve as an endpoint. Paragraphs
in tables, table cells, and text boxes cannot. Both endpoints are included by
default. With an end anchor, the four combinations are:

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

Every source endpoint and the destination insertion target must resolve wholly
inside one paragraph. A multi-paragraph search span remains useful for
inspection, but composition raises |BoundaryViolationError| instead of
inferring its first or last block. Select exact text inside one paragraph, or
pass an explicit live block for each endpoint.

Source endpoints are read-only evidence and may be live blocks or spans from
any supported inspection view. A live destination authorizes mutation only
when it was captured from ``view="current"``; reacquire a historical target
from the current view rather than relying on matching text or position.

Composition also refuses when insertion after the complete destination block
would remain inside an open complex-field result. This applies equally to a
paragraph destination and to the final table or block content control chosen
by ``append_document()``. Choose a current-view paragraph after the matching
field end, or deliberately close or unlink the field before retrying. The
package does not move the insertion point or repair field markers implicitly.

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
