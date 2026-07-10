
.. _paper_protection_api:

Document protection
===================

*paper-docx addition.* Word's Restrict-Editing setting
(``w:documentProtection``) is advisory, not security. Every paper-docx mutating
API raises |DocumentProtectedError| on an enforced setting;
``acknowledge_protection`` is the one explicit override. The setting itself is
preserved, and upstream APIs are untouched.

.. currentmodule:: docx.protection


.. autofunction:: protection_status

.. autofunction:: acknowledge_protection


|ProtectionStatus| objects
--------------------------

.. autoclass:: ProtectionStatus()
   :members:
   :undoc-members:
   :member-order: bysource
