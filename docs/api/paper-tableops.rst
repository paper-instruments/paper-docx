
.. _paper_tableops_api:

Table operations
================

*paper-docx addition.* Edit cells and rows with typed refusals for structures
that cannot be handled safely, including merged cells and nested tables. Cell
text routes through the |Span| machinery, so ``tracked=True`` produces a real
revision.

``find_table()`` searches top-level tables in the main document body. It uses
literal ``match="exact"`` matching by default; pass ``match="normalized"`` to
fold case, typography, and whitespace explicitly. The requested text must occur
within one physical cell: paragraph boundaries inside that cell remain literal
newlines in exact mode (and may fold as whitespace in normalized mode), while a
match never crosses between cells. No matching table raises
|TargetNotFoundError|, one matching table is returned, and multiple matching
tables raise |AmbiguousTargetError| with guidance to use more specific text.
An empty or normalization-empty ``near_text`` raises |TargetNotFoundError|
instead of matching every table.

.. currentmodule:: docx.tableops


.. autofunction:: find_table

.. autofunction:: update_cell

.. autofunction:: insert_row_after

.. autofunction:: delete_row
