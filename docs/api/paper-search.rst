
.. _paper_search_api:

Find and replace
================

Normalized, run-fragmentation-tolerant matching. ``find_text`` and
``find_one`` locate visible text the way a person quotes it (smart
quotes, dashes, exotic spaces and case are normalized) and return a
|Span| mapping it back to the exact runs. |Span| ``.replace`` edits
surgically — untouched runs keep their formatting byte-for-byte — or,
with ``tracked=True``, emits a minimal genuine ``w:ins``/``w:del``
redline; |Span| ``.comment`` anchors a comment to exactly the span.

.. currentmodule:: docx.search


.. autofunction:: normalize_text

.. autofunction:: find_text

.. autofunction:: find_one

.. autofunction:: replace_all


|Span| objects
--------------

.. autoclass:: Span()
   :members:
   :undoc-members:
   :member-order: bysource


|ReplaceResult| objects
-----------------------

.. autoclass:: ReplaceResult()
   :members:
   :undoc-members:
   :member-order: bysource


|ReplaceAllResult| objects
--------------------------

.. autoclass:: ReplaceAllResult()
   :members:
   :undoc-members:
   :member-order: bysource
