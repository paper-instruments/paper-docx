
.. _paper_formatting_api:

Effective-format resolver
=========================

*paper-docx addition.* Resolve what formatting text *actually* carries through
document defaults, the paragraph-style chain, the character style, and direct
formatting, with correct toggle-property semantics. The operation is read-only.
Every value in the returned |EffectiveFormat| names its source layer; anything
the resolver cannot determine is reported as unresolved.

``surrounding_format()`` requires its string or live-span target to resolve
wholly inside one paragraph. Multi-paragraph spans remain valid search results
for inspection, but this paragraph-level lookup raises
|BoundaryViolationError| rather than choosing the first or last paragraph.

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
