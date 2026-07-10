
.. _paper_scrubbing_api:

Finalize and scrub
==================

*paper-docx addition.* Finalize and scrub a document through
``Document.finalize`` / ``Document.scrub``. ``finalize`` resolves every tracked
revision or refuses and names what blocked it. ``scrub`` removes reviewing
residue (comments, metadata, the track-changes setting, optional RSIDs, and
hidden text) and returns a |ScrubReport| itemizing exactly what left the
package. Document protection is reported and preserved.

.. currentmodule:: docx.scrubbing


.. autofunction:: finalize

.. autofunction:: scrub


|ScrubReport| objects
---------------------

.. autoclass:: ScrubReport()
   :members:
   :undoc-members:
   :member-order: bysource
