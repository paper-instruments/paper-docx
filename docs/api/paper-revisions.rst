
.. _paper_revisions_api:

Tracked-change resolution
=========================

*paper-docx addition.* ``Document.revisions`` returns a |Revisions| snapshot
over every story part. It resolves insertions, deletions, run and paragraph
format changes, table-row revisions, and moves as paired units. Exotic markup
is enumerated and refused *by name*. ``accept_all``/``reject_all`` validate the
whole selected set before applying any changes.

.. currentmodule:: docx.revision



|Revision| objects
------------------

.. autoclass:: Revision()
   :members:
   :undoc-members:
   :member-order: bysource


|Revisions| objects
-------------------

.. autoclass:: Revisions()
   :members:
   :member-order: bysource
