
.. _paper_revisions_api:

Tracked-change resolution
=========================

``Document.revisions`` returns a |Revisions| snapshot over every story
part. Insertions, deletions, run/paragraph format changes, table-row
revisions and moves (as paired units) resolve; exotic markup is
enumerated and refused *by name*. ``accept_all``/``reject_all`` validate
the whole selected set first and never half-resolve.

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
