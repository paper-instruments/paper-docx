
.. _paper_fields_api:

Field authoring
===============

*paper-docx addition.* Author page number and count fields, dates,
cross-references to a bookmark, and tables of contents. Every inserted field
carries placeholder result text and sets the update-fields-on-open flag. This
package authors the field formulas. Word computes their displayed values when
it opens and paginates the document.

.. currentmodule:: docx.fields


.. autofunction:: add_page_number_field

.. autofunction:: add_page_count_field

.. autofunction:: add_date_field

.. autofunction:: add_reference_field

.. autofunction:: insert_toc_after
