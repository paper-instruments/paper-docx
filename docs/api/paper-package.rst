
.. _paper_package_api:

Package kernel
==============

The corruption-proof package layer. ``patch_save`` writes a compare-based
narrow save — parts you did not semantically change keep their original
bytes, so a file-level diff shows your edit and nothing else. ``diff_package``
and ``text_diff`` prove what changed; ``diagnose`` triages an unopenable
file into a typed verdict; ``compare`` generates a native tracked-change
redline transforming one document into another. These names are the
pinned public path (``docx.package.*``); the implementation lives in
private modules.

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
