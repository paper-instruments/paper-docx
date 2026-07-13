
.. _paper_package_api:

Package kernel
==============

*paper-docx addition.* Compare packages semantically and save edits narrowly.
``patch_save`` keeps the original bytes for parts you did not semantically
change, so a file-level diff shows your edit and nothing else. ``diff_package``
and ``text_diff`` report what changed. ``diagnose`` triages an unopenable file
into a typed verdict. ``compare`` generates a native tracked-change redline
that transforms one document into another. It verifies acceptance and
rejection on private copies before returning and refuses differences it cannot
encode without loss. These names are the pinned public path
(``docx.package.*``); the implementation lives in private modules.

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
