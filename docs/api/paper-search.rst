
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

Correct an existing insertion
-----------------------------

``span.replace(text, preserve_revision=True)`` corrects a current-view span
wholly owned by one existing ``w:ins`` without reauthoring or restamping it.
The text remains attributed to the insertion's recorded author and date, the
wrapper ID is retained, and accepting or rejecting the revision keeps its
original meaning. Outside revision markup the option behaves like an ordinary
untracked edit, so ``replace_all`` can update base text and insertion-owned
matches in one batch.

The operation refuses deletions, tracked moves, mixed base/insertion text,
multiple or nested revision wrappers, and non-current views. ``tracked=True``
cannot be combined with ``preserve_revision=True``. A fully preflighted no-op
changes nothing and leaves the span reusable. |ReplaceResult|
``preserved_revision_ids`` reports the existing insertion actually retained;
``revision_ids`` continues to report only newly authored revisions. Direct and
batch serialized results retain their schema discriminators and earlier keys.

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
