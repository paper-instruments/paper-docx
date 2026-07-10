
.. _paper_search_api:

Find and replace
================

*paper-docx addition.* Match normalized text across run fragmentation.
``find_text`` and ``find_one`` locate visible text the way a person quotes it,
normalizing smart quotes, dashes, exotic spaces, and case. They return a |Span|
that maps the match back to its exact runs. |Span| ``.replace`` changes only
the matched text; untouched runs keep their formatting byte-for-byte. With
``tracked=True``, it emits a minimal genuine ``w:ins``/``w:del`` redline.
|Span| ``.comment`` anchors a comment to exactly the span.

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
