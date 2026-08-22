
.. _paper_formatting_api:

Effective-format resolver
=========================

*paper-docx addition.* Resolve what formatting text *actually* carries through
document defaults, the paragraph-style chain, the character style, and direct
formatting, with correct toggle-property semantics. The operation is read-only.
Every value in the returned |EffectiveFormat| names its source layer; anything
the resolver cannot determine is reported as unresolved.

``surrounding_format()`` has two target classes. A string is found with exact
matching and retains the resulting inline |Span|; a supplied live span retains
the location the caller already selected. Their result is the effective format
of every touched run, including ``mixed`` values when properties disagree and
``agreeing_layers`` when equal values have different provenance. To select with
normalization, first call ``find_one(..., match="normalized")`` and pass that
span. This function has no separate matching-policy option.

A live paragraph |Block| or current ``BlockLocator`` identifies no character
position, so it returns paragraph properties plus the run defaults implied by
the paragraph style and document defaults. Existing text runs are not sampled;
this is also the result for an empty paragraph. A table-kind block or locator
has no implied paragraph and raises |UnsupportedStructureError|. Select exact
text or a live span inside a particular table-cell paragraph instead.

Every inline target must be wholly inside one paragraph. Multi-paragraph spans
remain valid search results for inspection, but this lookup raises
|BoundaryViolationError| rather than choosing the first or last paragraph.
Foreign targets also raise |BoundaryViolationError|; stale or missing targets
raise |TargetNotFoundError|, and duplicate exact strings or locator candidates
raise |AmbiguousTargetError|.

.. currentmodule:: docx.formatting


.. autofunction:: format_of

.. autofunction:: surrounding_format


|EffectiveFormat| objects
-------------------------

.. autoclass:: EffectiveFormat()
   :members:
   :undoc-members:
   :member-order: bysource


|ResolvedValue| objects
-----------------------

.. autoclass:: ResolvedValue()
   :members:
   :undoc-members:
   :member-order: bysource
