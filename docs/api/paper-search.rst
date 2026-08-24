
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

The single-result ``find_one()`` resolver has no contextual-ranking keyword. It
resolves ordinary zero/one/many matching: no candidate raises
|TargetNotFoundError| and multiple candidates raise |AmbiguousTargetError|. Use
a more specific exact target, narrow ``story``, or deliberately use ``nth=``
when position is the intended identity.

On ``find_text()``, ``nth=`` is an explicit 1-based positional selector and is
mutually exclusive with ``near``. Combining them raises :exc:`ValueError`
before the document is searched. To use contextual ranking for inspection:

.. code-block:: python

   candidates = find_text(doc, "Payment terms", near="Renewal")
   # all candidates remain visible, even if context is absent or tied

   for candidate in candidates:
       print(candidate.story, candidate.anchor, candidate.text)

   # After inspecting the evidence, retain the deliberately chosen live span.
   target = candidates[chosen_index]

Choose a replacement policy
---------------------------

Ordinary replacement is preservation-safe by default. It considers every exact
prefix/suffix split that preserves the maximal number of characters, then
changes a residual interval only when all maximal alignments agree on that
interval and one writable formatting/inline-ancestry destination. One changed
text node supplies local evidence; multiple changed nodes must have identical
complete run-property XML and full inline ancestry. A positional marker or
non-text run node cannot sit inside the changed interval.

.. list-table::
   :header-rows: 1
   :widths: 23 32 45

   * - Intent
     - Call
     - Contract
   * - Ordinary untracked edit
     - ``span.replace(text)``
     - Preserves exact unchanged affixes only when maximal alignment identifies
       one changed interval and one formatting/inline-ancestry destination.
       Ambiguous, mixed, scope-crossing, or marker-crossing intent refuses.
   * - Author a new redline
     - ``span.replace(text, tracked=True, author=...)``
     - Uses the same unique maximal affix localization, then emits a minimal
       ``w:del``/``w:ins`` pair only when inserted text has one complete
       formatting/inline-ancestry destination. A direct tracked no-op is
       refused and a successful change consumes the span.
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

Ordinary and tracked replacement consider every non-overlapping exact prefix/suffix split
that preserves the maximal number of characters. It narrows only when all
maximal splits identify the same changed interval. A pure insertion may choose
between adjacent text nodes only when their complete run properties and inline
ancestry agree; otherwise the operation refuses with guidance to re-find and
replace only the intended exact substring. This private narrowing is an
optimization, not a text selector or proof of author intent.

A changed interval inside one text node keeps that node's complete formatting.
When it spans multiple text nodes, their complete canonical ``w:rPr`` and full
inline ancestry must agree. This comparison includes generic wrappers such as
hyperlinks, revisions, controls, smart tags, ``customXml``, and directional
containers; it does not infer equivalence from a partial effective-format model.
For tracked edits that insert text, this proof determines the one run-property
and ancestry outcome for the new ``w:ins``. A deletion-only tracked edit needs
unique localization but may cross differently formatted source runs because
each ``w:del/w:r`` retains its own complete source properties. When the proof
is ambiguous, target a smaller uniform substring or construct the intended
formatting and structure explicitly.

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

|ReplaceResult| sets ``preserved_formatting_regions`` for a successful ordinary
untracked replacement. For a no-op, it records that no formatting changed; a
no-op is not a formatting probe and leaves the span reusable. Tracked edits and
empty-cell creation report false. The flag is
orthogonal to ``preserved_revision_ids``: an authorized correction inside one
existing insertion reports both the preserved insertion ID and formatting-
region evidence. ``preserved_structure`` remains separate, and
``revision_ids`` continues to identify only newly authored revisions.
``replace_all`` accepts the same replacement options plus the same exact-by-
default ``match`` policy, records each per-match
|PaperRefusal| while continuing independent matches, and retains one batch
transaction. A stale target aborts and rolls back the batch. Matches already
equal to the replacement are skipped rather than producing no-op results.
Every successful text-changing direct ``Span.replace()`` consumes the supplied
span. Use the returned result as the mutation evidence and re-find the text
before another operation. A direct no-op, preflight refusal, or rolled-back
mutation changes nothing and leaves the supplied span reusable. ``replace_all``
does not return or expose the private spans it uses for each match. Because a
no-op has no text assignments, it does not apply a hypothetical mutation's
bookmark-hollowing or changed-region checks.

When a replacement refuses because the changed interval is mixed or crosses a
marker, target a smaller span wholly inside one formatting/structural region.
Markers and differently formatted text wholly contained in exact unchanged
affixes remain in their original elements and order.

The direct and batch ``to_dict()`` payloads use schema version 2 and retain
their earlier keys while adding ``preserved_formatting_regions``.

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
