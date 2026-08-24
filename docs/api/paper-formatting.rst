
.. _paper_formatting_api:

Effective-format resolver
=========================

*paper-docx addition.* Resolve what formatting text *actually* carries through
document defaults, the paragraph-style chain, the character style, and direct
formatting, with correct toggle-property semantics. The operation is read-only.
Every value in the returned |EffectiveFormat| names its source layer; anything
the resolver cannot determine is reported as unresolved.

``format_of()`` accepts a :class:`~docx.text.run.Run`, a
:class:`~docx.text.paragraph.Paragraph`, or a live |Span|. Runs resolve their
supported run properties. Paragraphs resolve paragraph properties plus the run
defaults implied by their style chain. Spans resolve every run they touch,
reporting ``mixed`` where values disagree and ``agreeing_layers`` where equal
values have different provenance.

Selection is explicit. Use ``find_one()`` with the desired exact or normalized
matching policy, then pass the resulting span to ``format_of()``:

::

    from docx.formatting import format_of
    from docx.search import find_one

    span = find_one(doc, "payment terms", match="normalized")
    effective = format_of(span)

When a caller already holds a paragraph proxy, pass it directly instead. The
resolver does not search strings or resolve blocks. A stale span refuses during
freshness validation; unsupported target types raise ``TypeError``.

.. currentmodule:: docx.formatting


.. autofunction:: format_of


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
