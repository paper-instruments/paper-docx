
.. _paper_protection_api:

Document protection
===================

*paper-docx addition.* Word's Restrict-Editing setting
(``w:documentProtection``) is advisory, not security. paper-docx mutating APIs
raise |DocumentProtectedError| wherever Word's own UI would refuse the edit: the
gate mirrors Word's per-mode rules, so a comments-only restriction still permits
commenting while it blocks body edits. ``acknowledge_protection`` is the one
explicit override. The setting itself is preserved, and upstream APIs are
untouched.

.. currentmodule:: docx.protection


.. autofunction:: protection_status

.. autofunction:: acknowledge_protection


|ProtectionStatus| objects
--------------------------

.. autoclass:: ProtectionStatus()
   :members:
   :undoc-members:
   :member-order: bysource
