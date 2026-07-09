
paper-docx
==========

Release v\ |version| (:ref:`Installation <install>`)

*paper-docx* is an agent-first, strict-superset hard fork of `python-docx`_ for
safely inspecting, editing, reviewing, and composing existing Word documents.
The distribution is renamed; the import name stays ``docx``, so existing code
keeps working unchanged. It adds complete document traversal, normalized
find-and-replace, native tracked changes and their resolution, compare, scrub,
and cross-document composition. Every added operation either does exactly what
it claims or refuses atomically instead of risking silent corruption. See
:ref:`Paper additions <paper_additions>` for the added APIs.

.. _python-docx: https://github.com/python-openxml/python-docx

The rest of this page is inherited from python-docx and covers the shared
foundation for creating and updating Microsoft Word (.docx) files.


What it can do
--------------

.. |img| image:: /_static/img/example-docx-01.png

Here's an example of what |docx| can do:

============================================  ===============================================================
|img|                                         ::

                                                from docx import Document
                                                from docx.shared import Inches

                                                document = Document()

                                                document.add_heading('Document Title', 0)

                                                p = document.add_paragraph('A plain paragraph having some ')
                                                p.add_run('bold').bold = True
                                                p.add_run(' and some ')
                                                p.add_run('italic.').italic = True

                                                document.add_heading('Heading, level 1', level=1)
                                                document.add_paragraph('Intense quote', style='Intense Quote')

                                                document.add_paragraph(
                                                    'first item in unordered list', style='List Bullet'
                                                )
                                                document.add_paragraph(
                                                    'first item in ordered list', style='List Number'
                                                )

                                                document.add_picture('monty-truth.png', width=Inches(1.25))

                                                records = (
                                                    (3, '101', 'Spam'),
                                                    (7, '422', 'Eggs'),
                                                    (4, '631', 'Spam, spam, eggs, and spam')
                                                )

                                                table = document.add_table(rows=1, cols=3)
                                                hdr_cells = table.rows[0].cells
                                                hdr_cells[0].text = 'Qty'
                                                hdr_cells[1].text = 'Id'
                                                hdr_cells[2].text = 'Desc'
                                                for qty, id, desc in records:
                                                    row_cells = table.add_row().cells
                                                    row_cells[0].text = str(qty)
                                                    row_cells[1].text = id
                                                    row_cells[2].text = desc

                                                document.add_page_break()

                                                document.save('demo.docx')
============================================  ===============================================================


User Guide
----------

.. toctree::
   :maxdepth: 1

   user/install
   user/quickstart
   user/paper-additions
   user/documents
   user/tables
   user/text
   user/sections
   user/hdrftr
   user/api-concepts
   user/styles-understanding
   user/styles-using
   user/comments
   user/shapes


API Documentation
-----------------

.. toctree::
   :maxdepth: 2

   api/document
   api/settings
   api/style
   api/text
   api/table
   api/section
   api/comments
   api/shape
   api/dml
   api/shared
   api/enum/index


Paper additions — API reference
-------------------------------

Reference pages for the APIs the fork adds. See
:ref:`Paper additions <paper_additions>` for the narrative overview.

.. toctree::
   :maxdepth: 1

   api/paper-story
   api/paper-search
   api/paper-blocks
   api/paper-revisions
   api/paper-tableops
   api/paper-numbering
   api/paper-controls
   api/paper-commentops
   api/paper-package
   api/paper-scrubbing
   api/paper-protection
   api/paper-composition
   api/paper-bookmarks
   api/paper-fields
   api/paper-formatting
   api/paper-errors


Contributor Guide
-----------------

.. toctree::
   :maxdepth: 1

   dev/analysis/index
