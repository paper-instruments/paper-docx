
.. _paper_search_api:

Find and replace
================

*paper-docx addition.* Match visible text across run fragmentation.
``find_text`` and ``find_one`` use literal exact matching by default. They
return a |Span| that maps the match back to its concrete text nodes. |Span|
``.replace`` covers five replacement intents; |Span| ``.comment`` anchors a
comment to exactly the span.

Choose a matching policy
------------------------

Exact matching is the mutation-safe default for ``find_text``, ``find_one``,
and ``replace_all``. It compares Unicode codepoints literally, including case,
punctuation, and whitespace. Ordinary Word run boundaries do not insert text,
so an exact match can cross any number of runs in one paragraph. A paragraph
boundary is represented in raw search text by exactly one ``"\n"``; a space or
``"\r"`` does not substitute for it. A candidate consisting only of a paragraph
separator cannot form a live span because that separator does not belong to a
text atom. An inline Word line break is itself a visible text atom and also has
the exact representation ``"\n"``.

Pass ``match="normalized"`` when the caller intentionally wants the previous
convenience behavior. Normalized matching case-folds, maps smart quotes and
dashes to their ASCII forms, maps exotic spaces and tabs to spaces, collapses
whitespace runs, and removes soft hyphens. It still returns the exact document
characters in ``Span.text`` and the resulting live span remains a valid
mutation target. The selected policy applies to both ``needle`` and ``near``.

Search-produced spans expose that evidence as ``Span.match_policy``
(``"exact"`` or ``"normalized"``). Spans derived from already-known live
offsets have no search policy and report ``None``. Empty needles never match;
exact whitespace is literal searchable content, while normalized input that
folds to only whitespace does not match.

Disambiguate without hiding uncertainty
----------------------------------------

``find_text(..., near=...)`` is an inspection operation. It returns every
target candidate, ordered by the character distance from each target's start
to its nearest eligible context occurrence. Target and context use the same
story, view, and match policy. Multiple context occurrences are valid; each
candidate uses the closest one. Equal distances retain stable document order,
and missing context leaves the complete candidate set in ordinary document
order. Neither case means that the first result has unique authority.

``find_one(..., near=...)`` applies the authoritative form of that selector.
It returns a span only when eligible context exists and exactly one target has
the finite minimum distance. Missing context raises |TargetNotFoundError| and
points the caller toward its context text, match policy, view, and story
scope. A tied minimum raises |AmbiguousTargetError| with the tied locations
and distance; use more distinctive context or a narrower story scope.
Context in a different story or excluded by the selected view cannot rank a
target.

Use ``nth=`` only as an explicit 1-based positional selector when ``near`` is
absent. ``near`` and ``nth`` are mutually exclusive and combining them raises
:exc:`ValueError` before the document is searched. To inspect an inconclusive
contextual selection before refining it:

.. code-block:: python

   candidates = find_text(doc, "Payment terms", near="Renewal")
   # all candidates remain visible, even if context is absent or tied

   target = find_one(doc, "Payment terms", near="Renewal")
   # succeeds only for one unique nearest candidate

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
``replace_all`` accepts the same replacement options plus the same exact-by-
default ``match`` policy, records each per-match
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
