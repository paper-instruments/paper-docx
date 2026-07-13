
.. _paper_formatting_api:

Effective-format resolver
=========================

*paper-docx addition.* Resolve what formatting text *actually* carries through
document defaults, the paragraph-style chain, the character style, and direct
formatting, with correct toggle-property semantics. The operation is read-only.
Every value in the returned |EffectiveFormat| names its source layer; anything
the resolver cannot determine is reported as unresolved.

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
