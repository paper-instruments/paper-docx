
.. _paper_tableops_api:

Table operations
================

Cell and row edits that refuse loudly on structures they cannot handle
safely — merged cells, nested tables — instead of guessing. Cell text
routes through the |Span| machinery, so ``tracked=True`` produces a
real revision.

.. currentmodule:: docx.tableops


.. autofunction:: find_table

.. autofunction:: update_cell

.. autofunction:: insert_row_after

.. autofunction:: delete_row
