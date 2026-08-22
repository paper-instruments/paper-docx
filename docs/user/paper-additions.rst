.. _paper_additions:

Paper additions
===============

*paper-docx* is an agent-first, strict-superset hard fork of python-docx. The
distribution is renamed; the import name stays ``docx``. Everything documented
in the rest of this guide still works exactly as before, and upstream's own test
suites run green on every commit. This page summarizes the editing and review
APIs the fork adds for existing documents.

Safety contract
---------------

The fork exists to prevent **silent corruption**: a document that opens fine
and is quietly wrong after direct XML edits. Every added operation either does
exactly what it claims or refuses atomically. Mutating operations validate fully
before they touch anything or restore the captured state on failure, so a
raised |PaperRefusal| leaves the document and supplied live proxies unchanged.
Callers can handle the refusal separately from programmer errors.

.. contents::
   :local:
   :depth: 1


Perceive a document
--------------------

Stock ``Document.paragraphs`` sees only the body. It is blind to text inside
tracked insertions, content controls, text boxes, footnotes and endnotes.
:ref:`docx.story <paper_story_api>` walks every story part and every region
standard traversal misses, under a chosen *view* — the document as it stands
(``"current"``), before its pending revisions (``"original"``), or everything
at once (``"all"``). |Outline| reports what it could not read.

.. highlight:: python

::

    from docx.story import outline

    o = outline(doc)
    o.blind_region_counts                    # {"tracked_insertions": 2, "text_boxes": 1, ...}
    [b.text for b in o.blocks if b.in_text_box]

Blocks in an outline are live targets bound to the exact document element.
Use them directly during the current in-memory editing session:

::

    from docx.story import iter_blocks

    block = next(b for b in iter_blocks(doc) if b.text == "Termination")
    insert_section_after(doc, block, heading="Notice", paragraphs=[])

A live block follows its exact element across harmless index shifts and becomes
stale if that element is detached, replaced, or structurally moved. Reacquire a
fresh block from the document in a later session.

The view that produced a live block or span is also part of its editing
authority. ``"original"`` and ``"all"`` projections remain available for
inspection, but paragraph and TOC mutations require a target captured from
``view="current"``. When a historical target is refused, inspect it as needed,
then deliberately reacquire the intended current target; the package does not
silently remap it by text or position.

The older ``block.anchor`` and ``revision.anchor`` values remain inert result
locations for compatibility. They cannot authorize paragraph mutation. Use a
live block/span or an exact string instead.

:ref:`docx.search <paper_search_api>` matches literal visible text by default,
across the multiple runs Word fragments text into. Callers can explicitly opt
into normalization of smart quotes, dashes, exotic spaces, whitespace and
case. The returned |Span| maps the exact document text back to the runs that
hold it.
:ref:`docx.formatting <paper_formatting_api>` answers the complementary
question: what formatting does this text *actually* carry? It resolves through
document defaults, the style chain and direct formatting, with every value
naming the layer it came from. String targets use exact search and report the
formatting at the matched text, rather than a representative run elsewhere in
the paragraph. A live span keeps its already-selected location, including one
obtained through explicit normalized search:

::

    from docx.formatting import surrounding_format
    from docx.search import find_one

    exact = surrounding_format(doc, "Payment terms")
    normalized_span = find_one(doc, "payment terms", match="normalized")
    normalized = surrounding_format(doc, normalized_span)

A span touching multiple runs reports ``mixed`` for disagreeing properties and
preserves provenance when equal values come from different layers. A live
paragraph block or current ``BlockLocator`` carries paragraph identity but no
inline position, so it reports the paragraph's style/document defaults even
when the paragraph is empty; it never samples the first existing run. A table
block has no implied paragraph target and refuses. Select exact text or a live
span inside the intended cell paragraph instead.


Edit one document
-----------------

|Span| replaces matched text while untouched runs retain their formatting.

::

    from docx.search import find_one, find_text

    span = find_one(doc, "rate: $75–100/hr")
    span.replace("rate: $85-110/hr")             # formatting intact

Exact targeting is the mutation-safe default. To migrate a convenience lookup
that intentionally ignores typography or case, opt in once and pass the
resulting live span to the mutation API:

::

    span = find_one(
        doc,
        'rate: $75-100/hr on a "full-service" basis',
        match="normalized",
    )
    span.text                                    # preserves the smart punctuation
    span.replace("rate: $85-110/hr")

The normalized span is an intentional mutation target, not an inspection-only
result. APIs outside ``docx.search`` do not repeat the ``match`` keyword; use
this explicit preselection workflow when they need normalized targeting.

Search can also return spans that cross paragraph boundaries, representing the
boundary as a literal ``"\\n"``. Those spans are valid for inspection. APIs
that need one paragraph—block edits, composition endpoints, TOC insertion,
and ``surrounding_format()``—raise |BoundaryViolationError| instead of choosing
the first or last paragraph. Refine the text to one paragraph or pass an
explicit live block endpoint.

When repeated text needs context, use ``near`` to rank the complete candidate
set for inspection:

::

    candidates = find_text(doc, "Payment terms", near="Renewal")
    # inspection stays complete and proximity-ranked, including ties

    for candidate in candidates:
        print(candidate.story, candidate.anchor, candidate.text)

Context in a different story or excluded by the selected view cannot rank a
target. Missing or tied context leaves all candidates visible in stable order;
the first result is not automatically authoritative. After inspecting the
evidence, retain the deliberately chosen live span, or refine an ordinary
``find_one`` call with a more specific exact target, ``story``, or explicit
``nth``. On ``find_text``, ``nth`` cannot be combined with ``near``.

Choose the option that matches the edit's preservation contract:

- Use the default for an ordinary untracked correction. The replacement takes
  the start run's formatting.
- Use ``tracked=True`` with ``author=...`` to author a new ``w:ins``/``w:del``
  redline.
- Use ``preserve_revision=True`` only to correct current-view text wholly
  inside one existing insertion while retaining that insertion's attribution
  and accept/reject behavior. The correction remains attributed to the
  recorded author and date. Base text encountered by ``replace_all`` still
  uses the ordinary untracked behavior.
- Use ``preserve_structure=True`` when the existing text-node and run topology
  must remain exact. This mode changes only text values and may refuse text
  whose whitespace cannot be represented without changing ``xml:space``.
- Combine the two preservation options when both contracts apply. Neither can
  be combined with ``tracked=True``.

These are text-structure contracts, not a promise that the serialized bytes of
the changed XML part remain identical. The full refusal rules and result
evidence are in :ref:`docx.search <paper_search_api>`.

:ref:`docx.blocks <paper_blocks_api>` does the clause-level equivalent (insert,
delete or replace whole paragraphs). :ref:`docx.tableops <paper_tableops_api>` and
:ref:`docx.numbering <paper_numbering_api>` cover validated table edits and
list numbering. :ref:`docx.controls <paper_controls_api>` fills content
controls, clearing placeholder state so Word treats them as filled.
:ref:`docx.bookmarks <paper_bookmarks_api>` and :ref:`docx.fields
<paper_fields_api>` create bookmarks and author fields (page numbers, dates,
cross-references, a table of contents), always as *formulas* with placeholder
results. Word computes the displayed values when it opens the document.


Review and resolve
------------------

``Document.revisions`` (:ref:`docx.revision <paper_revisions_api>`) enumerates
every tracked change across every story part and resolves them — insertions,
deletions, format changes, table-row revisions, and moves as paired units:

::

    for rev in doc.revisions:
        print(rev.revision_type, rev.author, repr(rev.text))
    doc.revisions.reject_all(author="Bob Reviewer")     # or accept_all()

Document protection is honored throughout:
:ref:`docx.protection <paper_protection_api>` makes fork mutating APIs refuse
with |DocumentProtectedError| wherever Word's own UI would refuse the edit,
rather than silently editing a locked template. The gate follows Word's per-mode
rules, so a comments-only restriction still permits commenting.


Save a bounded edit
-------------------

Use ordinary ``Document.save()`` for a new document or when whole-package
normalization is intentional. For a bounded edit to an existing file,
``patch_save(original_path, document, out_path)`` is the recommended path when
unrelated package-part bytes must remain untouched. It restores original bytes
for semantically unchanged XML parts; it does not preserve byte ranges inside
an XML part that the edit changed or the original ZIP container bytes after a
non-no-op save.

Use ``diff_package(original_path, out_path)`` as package-level evidence of the
changed-part budget, and reopen ``out_path`` as evidence that the saved document
remains semantically readable.


Compose across documents
------------------------

:ref:`docx.package <paper_package_api>` is the package inspection and compare
layer. ``compare`` generates a native tracked-change redline
transforming one document into another. Accepting it yields the revised
document; rejecting it yields the original. Before returning, ``compare``
proves both outcomes on private copies. Style, relationship, or package-part
changes it cannot express as tracked revisions produce a typed refusal.
Fine-grained word or cell edits require one unambiguous old/new block pair;
larger changed regions use coarse block revisions rather than similarity-based
pairing.

::

    from docx.package import compare

    result = compare("original.docx", "revised.docx", author="Reviewer")
    result.document.save("redline.docx")         # native w:ins/w:del markup

:ref:`docx.composition <paper_composition_api>` copies formatted content
between documents without corruption. It reconciles styles, numbering, media,
hyperlinks and bookmarks, then returns a |CompositionReport| listing every
part it touched. Historical live blocks and spans remain valid read-only source
range evidence, but a live destination must be reacquired from
``view="current"``. Composition refuses rather than guessing when insertion
after the destination paragraph, table, or block content control would remain
inside an open field result; target a current block after the matching field
end, or deliberately close or unlink that field first.


Refusal handling
----------------

Every refusal is a typed member of the :ref:`docx.errors <paper_errors_api>`
hierarchy, so a caller can tell a *safe refusal* apart from a bug:

::

    from docx.errors import PaperRefusal, AmbiguousTargetError

    try:
        find_one(doc, "the")                     # matches everywhere
    except AmbiguousTargetError:
        ...                                      # use context, story scope, or inspect candidates
    except PaperRefusal:
        ...                                      # any safe refusal, distinct from bugs

Programmer mistakes still raise the ordinary :exc:`TypeError`/:exc:`ValueError`.
The ``docx.api.paper-*`` reference pages document the complete API; the
repository README summarizes the fork and its relationship to python-docx.
