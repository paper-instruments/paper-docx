
.. _paper_package_api:

Package kernel
==============

*paper-docx addition.* Compare packages semantically and save edits narrowly.
``patch_save`` keeps original package-part bytes wherever it can prove that
serialized XML is semantically unchanged. ``diff_package`` and ``text_diff``
report what changed. ``diagnose`` triages an unopenable file into a typed
verdict. ``compare`` generates a native tracked-change redline that transforms
one document into another. It verifies acceptance and rejection on private
copies before returning and refuses differences it cannot encode without loss.
These names are the pinned public path
(``docx.package.*``); the implementation lives in private modules.

Choose a save path
------------------

Use ordinary ``Document.save()`` for a new document or when normalizing the
whole package is intentional. For a bounded edit to an existing package, use
``patch_save(original_path, document, out_path)`` when unrelated part bytes
must remain untouched. It serializes normally, restores the original bytes of
semantically unchanged XML parts, and reports restored, changed, added, and
removed parts.

This is a package-part boundary: ``patch_save`` does not promise byte-range
identity inside an XML part whose semantics changed, nor does it preserve the
original ZIP container bytes unless the entire save is a no-op. Use
``diff_package`` to record package-level change evidence, then reopen the
output to establish that the saved document is semantically readable.

.. currentmodule:: docx.package


.. autofunction:: xml_equivalent

.. autofunction:: diff_package

.. autofunction:: patch_save

.. autofunction:: diagnose

.. autofunction:: text_diff

.. autofunction:: pending_changes

.. autofunction:: compare


|PackageDiff| objects
---------------------

.. autoclass:: PackageDiff()
   :members:
   :undoc-members:
   :member-order: bysource


|PartDiff| objects
------------------

.. autoclass:: PartDiff()
   :members:
   :undoc-members:
   :member-order: bysource


|PatchSaveResult| objects
-------------------------

.. autoclass:: PatchSaveResult()
   :members:
   :undoc-members:
   :member-order: bysource


|PackageDiagnosis| objects
--------------------------

.. autoclass:: PackageDiagnosis()
   :members:
   :undoc-members:
   :member-order: bysource


|TextDiff| objects
------------------

.. autoclass:: TextDiff()
   :members:
   :undoc-members:
   :member-order: bysource


|StoryTextDiff| objects
-----------------------

.. autoclass:: StoryTextDiff()
   :members:
   :undoc-members:
   :member-order: bysource


|CompareResult| objects
-----------------------

.. autoclass:: CompareResult()
   :members:
   :undoc-members:
   :member-order: bysource


|CompareFinding| objects
------------------------

.. autoclass:: CompareFinding()
   :members:
   :undoc-members:
   :member-order: bysource
