
.. _paper_search_api:

Find and replace
================

*paper-docx addition.* Match normalized text across run fragmentation.
``find_text`` and ``find_one`` locate visible text the way a person quotes it,
normalizing smart quotes, dashes, exotic spaces, and case. They return a |Span|
that maps the match back to its concrete text nodes. |Span| ``.replace`` covers
five replacement intents; |Span| ``.comment`` anchors a comment to exactly the
span.

Choose a replacement policy
---------------------------

All preservation modes are opt-in. Existing calls retain the ordinary
untracked behavior.

.. list-table::
   :header-rows: 1
   :widths: 23 32 45

   * - Intent
     - Call
     - Contract
   * - Ordinary untracked edit
     - ``span.replace(text)``
     - The replacement takes the start run's formatting. Untouched runs keep
       their formatting, but selected text may be redistributed between runs.
   * - Author a new redline
     - ``span.replace(text, tracked=True, author=...)``
     - Emits a minimal ``w:del``/``w:ins`` pair and consumes the span. A direct
       tracked no-op is refused.
   * - Correct one existing insertion
     - ``span.replace(text, preserve_revision=True)``
     - For a current-view span wholly owned by one ``w:ins``, keeps that
       insertion's id, author, date, and accept/reject projections. Outside
       revision markup, behaves like an ordinary untracked edit.
   * - Preserve exact text topology
     - ``span.replace(text, preserve_structure=True)``
     - Changes only existing ``w:t`` values; preserves their attributes,
       ancestor runs, and intervening transparent markers.
   * - Correct an insertion and preserve topology
     - ``span.replace(text, preserve_revision=True,
       preserve_structure=True)``
     - Applies both contracts to one existing insertion.

``tracked=True`` cannot be combined with either preservation option. Revision
preservation does not reauthor or restamp the existing insertion and does not
support deletions, tracked moves, mixed base/insertion text, or multiple or
nested revision wrappers. The corrected text remains attributed to the
existing insertion's recorded author and date. Its outside-revision behavior
lets ``replace_all`` apply one policy to both base text and insertion-owned
matches. ``preserve_structure=True`` alone does not authorize an edit inside
an insertion; request both guarantees for that case.

Exact topology means structural preservation, not preservation of inferred
formatting intent. Replacement text fills each selected text-node slice from
left to right up to that slice's original capacity; the final selected node
receives the remainder. Empty nodes remain present. Existing field,
content-control, hyperlink, revision, protection, and paragraph-boundary guards
still apply. The operation also refuses when the result would need an
``xml:space`` attribute change, hollow a bookmark, or require placeholder
cleanup. A successful mutation consumes the span because its captured offsets
no longer describe the same text. A no-op still runs the full preflight and
reports preservation evidence, but changes nothing and leaves the span
reusable.

|ReplaceResult| sets ``preserved_structure`` when exact topology was requested
and satisfied. ``preserved_revision_ids`` contains the existing insertion ID
only when an insertion was actually retained; it is empty for base text.
``revision_ids`` continues to identify only newly authored revisions.
``replace_all`` accepts the same options, records each per-match
|PaperRefusal| while continuing independent matches, and retains one batch
transaction. A stale target aborts and rolls back the batch. Matches already
equal to the replacement are skipped rather than producing no-op results.
The direct and batch ``to_dict()`` payloads retain their schema discriminators
and earlier keys.

These contracts describe the edited XML structure. They do not promise raw
byte identity inside a changed package part; use :func:`docx.package.patch_save`
and :func:`docx.package.diff_package` when package-level change evidence
matters.

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
