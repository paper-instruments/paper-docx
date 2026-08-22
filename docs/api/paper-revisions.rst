
.. _paper_revisions_api:

Tracked-change resolution
=========================

*paper-docx addition.* ``Document.revisions`` returns a |Revisions| snapshot
over every story part. It resolves insertions, deletions, run and paragraph
format changes, table-row revisions, and moves as paired units. Exotic markup
is enumerated and refused *by name*. ``accept_all``/``reject_all`` validate the
whole selected set before applying any changes.

Each block-contained |Revision| exposes two deliberately different values:
``revision.anchor`` is inert compatibility/location evidence, while
``revision.block_locator`` is the containing block's portable exact locator.
Revision schema version 4 makes that status machine-readable with
``anchor_role="legacy_inert_location_evidence"``.
Body-level section-property changes belong to no traversed block and therefore
serialize ``block_locator`` as ``None``; their synthetic story-level Anchor
cannot serve as a paragraph mutation target. Revision accept/reject freshness
continues to use the live revision element, not either serialized value.

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
