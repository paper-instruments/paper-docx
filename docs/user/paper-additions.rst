.. _paper_additions:

Paper additions
===============

*paper-docx* is a strict-superset fork of python-docx. Everything documented
in the rest of this guide still works exactly as before — the import name is
still ``docx`` and upstream's own test suites run green on every commit. This
page is the on-ramp to what the fork adds on top: the editing surface that
real, agent-driven document work needs.

The whole design traces to one enemy. Production systems have historically
reached past python-docx's editing gaps with raw-XML surgery, whose dominant
failure mode is **silent corruption** — a file that opens fine and is quietly
wrong. Every fork API replaces that surgery with a first-class operation whose
failure mode is instead a loud, typed refusal: it validates fully *before* it
touches anything, so a raised |PaperRefusal| leaves the document — in memory
and on disk — exactly as it was. A refused edit is a success mode; a quietly
wrong file is the failure this fork exists to eliminate.

.. contents::
   :local:
   :depth: 1


Perceive a document
--------------------

Stock ``Document.paragraphs`` sees only the body and is blind to text inside
tracked insertions, content controls, text boxes, footnotes and endnotes.
:ref:`docx.story <paper_story_api>` walks every story part and every region
standard traversal misses, under a chosen *view* — the document as it stands
(``"current"``), before its pending revisions (``"original"``), or everything
at once (``"all"``). |Outline| even reports what it could not read.

.. highlight:: python

::

    from docx.story import outline

    o = outline(doc)
    o.blind_region_counts                    # {"tracked_insertions": 2, "text_boxes": 1, ...}
    [b.text for b in o.blocks if b.in_text_box]

:ref:`docx.search <paper_search_api>` finds visible text the way a person
quotes it — smart quotes, dashes, exotic spaces and case are normalized, and
matching assembles across the multiple runs Word fragments text into. The
returned |Span| maps that text back to the exact runs that hold it.
:ref:`docx.formatting <paper_formatting_api>` answers the complementary
question — what formatting does this text *actually* carry, resolved through
document defaults, the style chain and direct formatting, with every value
naming the layer it came from.


Edit one document
-----------------

|Span| edits are surgical: untouched runs keep their formatting byte-for-byte.

::

    from docx.search import find_one

    span = find_one(doc, "rate: $75-100/hr")     # matches “rate: $75–100/hr”
    span.replace("rate: $85-110/hr")             # formatting intact

The same call, with ``tracked=True``, emits a minimal genuine ``w:ins``/
``w:del`` redline that Word renders natively; :ref:`docx.blocks
<paper_blocks_api>` does the clause-level equivalent (insert, delete or
replace whole paragraphs). :ref:`docx.tableops <paper_tableops_api>` and
:ref:`docx.numbering <paper_numbering_api>` cover guarded table edits and
list numbering; :ref:`docx.controls <paper_controls_api>` fills content
controls (clearing placeholder state so Word treats them as genuinely filled);
:ref:`docx.bookmarks <paper_bookmarks_api>` and :ref:`docx.fields
<paper_fields_api>` create bookmarks and author fields (page numbers, dates,
cross-references, a table of contents) — always as *formulas* with placeholder
results, never computed values.


Review, resolve and finalize
----------------------------

``Document.revisions`` (:ref:`docx.revision <paper_revisions_api>`) enumerates
every tracked change across every story part and resolves them — insertions,
deletions, format changes, table-row revisions, and moves as paired units:

::

    for rev in doc.revisions:
        print(rev.revision_type, rev.author, repr(rev.text))
    doc.revisions.reject_all(author="Bob Reviewer")     # or accept_all()

When a file is ready to leave the building, ``Document.finalize`` resolves
every revision (or refuses, naming what blocked it) and ``Document.scrub``
removes reviewing residue — comments, metadata, RSIDs — returning a
|ScrubReport| itemizing exactly what left the package (:ref:`docx.scrubbing
<paper_scrubbing_api>`). Document protection is honored throughout:
:ref:`docx.protection <paper_protection_api>` makes every fork mutating API
refuse with |DocumentProtectedError| on a Restrict-Editing setting rather than
silently editing a locked template.


Work across documents
----------------------

:ref:`docx.package <paper_package_api>` is the corruption-proof package layer.
``patch_save`` writes a compare-based narrow save — parts you did not
semantically change keep their original bytes, so a file-level diff shows your
edit and nothing else — and ``compare`` generates a native tracked-change
redline transforming one document into another, with a tested algebra
(accepting it yields the revised document; rejecting it yields the original).

::

    from docx.package import compare

    result = compare("original.docx", "revised.docx", author="Reviewer")
    result.document.save("redline.docx")         # native w:ins/w:del markup

:ref:`docx.composition <paper_composition_api>` copies formatted content
between documents without corruption — reconciling styles, numbering, media,
hyperlinks and bookmarks — and returns a |CompositionReport| declaring every
part it touched.


When it refuses
---------------

Every refusal is a typed member of the :ref:`docx.errors <paper_errors_api>`
hierarchy, so a caller can tell a *safe refusal* apart from a bug:

::

    from docx.errors import PaperRefusal, AmbiguousTargetError

    try:
        find_one(doc, "the")                     # matches everywhere
    except AmbiguousTargetError:
        ...                                      # disambiguate: nth=, near=, story=
    except PaperRefusal:
        ...                                      # any safe refusal, distinct from bugs

Programmer mistakes still raise the ordinary :exc:`TypeError`/:exc:`ValueError`.
For the full lineage, shipped surface and design principles, see ``PAPER.md``
and ``API-PROPOSAL.md`` in the repository.
