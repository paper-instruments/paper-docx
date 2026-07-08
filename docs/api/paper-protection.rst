
.. _paper_protection_api:

Document protection
===================

Word's Restrict-Editing setting (``w:documentProtection``) is advisory,
not security, but silently editing a locked template is exactly the
fail-loudly violation this fork exists to prevent. Every paper-docx
mutating API refuses with |DocumentProtectedError| on an enforced
setting; ``acknowledge_protection`` is the one explicit override. The
setting itself is never stripped, and upstream APIs are untouched.

.. currentmodule:: docx.protection


.. autofunction:: protection_status

.. autofunction:: acknowledge_protection


|ProtectionStatus| objects
--------------------------

.. autoclass:: ProtectionStatus()
   :members:
   :undoc-members:
   :member-order: bysource
