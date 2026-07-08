
.. _paper_numbering_api:

Numbering
=========

Report the list definitions a document already carries, apply an
existing definition to a paragraph, or author a real bullet/decimal
definition on demand. ``apply_list_style`` refuses rather than produce a
fake bullet (a list-styled paragraph that renders no marker).

.. currentmodule:: docx.numbering


.. autofunction:: list_numbering

.. autofunction:: apply_numbering

.. autofunction:: apply_list_style

.. autofunction:: ensure_bullet_definition

.. autofunction:: ensure_decimal_definition

.. autofunction:: restart_numbering


|NumberingReport| objects
-------------------------

.. autoclass:: NumberingReport()
   :members:
   :undoc-members:
   :member-order: bysource


|NumberingDefinition| objects
-----------------------------

.. autoclass:: NumberingDefinition()
   :members:
   :undoc-members:
   :member-order: bysource


|NumberingLevel| objects
------------------------

.. autoclass:: NumberingLevel()
   :members:
   :undoc-members:
   :member-order: bysource


|NumberedParagraph| objects
---------------------------

.. autoclass:: NumberedParagraph()
   :members:
   :undoc-members:
   :member-order: bysource
