
.. _paper_story_api:

Story traversal
===============

Visibility-complete traversal. Where ``Document.paragraphs`` sees only
the body and skips tracked insertions, content controls, text boxes and
notes, these functions walk every story part (body, headers, footers,
footnotes, endnotes, comments) and every region standard traversal is
blind to, under a chosen *view* (``"current"``, ``"original"`` or
``"all"``). Blocks carry stable |Anchor| identities usable as edit
targets, and |Outline| confesses what it could not read.

.. currentmodule:: docx.story


.. autofunction:: story_parts

.. autofunction:: iter_blocks

.. autofunction:: outline


|Outline| objects
-----------------

.. autoclass:: Outline()
   :members:
   :undoc-members:
   :member-order: bysource


|Block| objects
---------------

.. autoclass:: Block()
   :members:
   :undoc-members:
   :member-order: bysource


|TableShape| objects
--------------------

.. autoclass:: TableShape()
   :members:
   :undoc-members:
   :member-order: bysource


|Anchor| objects
----------------

.. autoclass:: Anchor()
   :members:
   :undoc-members:
   :member-order: bysource
