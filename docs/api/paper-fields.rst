
.. _paper_fields_api:

Field authoring
===============

*paper-docx addition.* Author page number and count fields, dates,
cross-references to a bookmark, and tables of contents. Every inserted field
carries placeholder result text and sets the update-fields-on-open flag. This
package authors the field formulas. Word computes their displayed values when
it opens and paginates the document.

``insert_toc_after()`` accepts exact strings and live paragraph targets. A
live |Block| or |Span| destination must be captured from ``view="current"``;
historical projections remain inspection-only and must be reacquired before
TOC insertion.

The target passed to ``insert_toc_after()`` must resolve wholly inside one
paragraph. A multi-paragraph string or live span raises
|BoundaryViolationError|; choose exact text within the intended paragraph or
pass an explicit live block.

.. currentmodule:: docx.fields


.. autofunction:: add_page_number_field

.. autofunction:: add_page_count_field

.. autofunction:: add_date_field

.. autofunction:: add_reference_field

.. autofunction:: insert_toc_after
