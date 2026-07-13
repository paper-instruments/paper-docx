
.. _paper_tableops_api:

Table operations
================

*paper-docx addition.* Edit cells and rows with typed refusals for structures
that cannot be handled safely, including merged cells and nested tables. Cell
text routes through the |Span| machinery, so ``tracked=True`` produces a real
revision.

.. currentmodule:: docx.tableops


.. autofunction:: find_table

.. autofunction:: update_cell

.. autofunction:: insert_row_after

.. autofunction:: delete_row
