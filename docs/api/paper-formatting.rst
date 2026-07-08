
.. _paper_formatting_api:

Effective-format resolver
=========================

Read-only: what formatting does this text *actually* carry, resolved
through document defaults, the paragraph-style chain, the character
style and direct formatting — with correct toggle-property semantics.
Every value in the returned |EffectiveFormat| names the layer it came
from, and what the resolver cannot determine is declared, never guessed.

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
