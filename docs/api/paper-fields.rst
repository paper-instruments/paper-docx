
.. _paper_fields_api:

Field authoring
===============

Author fields — page number and count, date, cross-references to a
bookmark, and a table of contents. Every inserted field carries
placeholder result text and arms the update-fields-on-open flag: this
package authors *formulas*, and never computes their values (pagination
is a renderer's job).

.. currentmodule:: docx.fields


.. autofunction:: add_page_number_field

.. autofunction:: add_page_count_field

.. autofunction:: add_date_field

.. autofunction:: add_reference_field

.. autofunction:: insert_toc_after
