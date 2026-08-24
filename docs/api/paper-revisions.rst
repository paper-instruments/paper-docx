
.. _paper_revisions_api:

Tracked-change resolution
=========================

*paper-docx addition.* ``Document.revisions`` returns a |Revisions| snapshot
over every story part. It resolves insertions, deletions, run and paragraph
format changes, table-row revisions, and moves as paired units. Exotic markup
is enumerated and refused *by name*. ``accept_all``/``reject_all`` validate the
whole selected set before applying any changes.

Each |Revision| exposes ``revision.anchor`` as inert compatibility/location
evidence. Revision schema version 4 makes that status machine-readable with
``anchor_role="legacy_inert_location_evidence"``.
Body-level section-property changes belong to no traversed block; their
synthetic story-level Anchor cannot serve as a paragraph mutation target.
Revision accept/reject freshness continues to use the live revision element,
not the serialized value.

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
