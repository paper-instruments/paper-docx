
.. _paper_errors_api:

Refusal hierarchy
=================

*paper-docx addition.* Every added operation either does exactly what it claims
or refuses atomically. Mutating operations validate fully before they touch the
document, so a raised |PaperRefusal| leaves the in-memory tree and any file on
disk byte-for-byte unchanged. Programmer mistakes remain ordinary
:exc:`TypeError`/:exc:`ValueError`. Callers can catch |PaperRefusal| separately.

.. currentmodule:: docx.errors



|PaperRefusal|
--------------

.. autoexception:: PaperRefusal
   :show-inheritance:


|MalformedPackageError|
-----------------------

.. autoexception:: MalformedPackageError
   :show-inheritance:


|AmbiguousTargetError|
----------------------

.. autoexception:: AmbiguousTargetError
   :show-inheritance:


|TargetNotFoundError|
---------------------

.. autoexception:: TargetNotFoundError
   :show-inheritance:


|UnsupportedStructureError|
---------------------------

.. autoexception:: UnsupportedStructureError
   :show-inheritance:


|BoundaryViolationError|
------------------------

.. autoexception:: BoundaryViolationError
   :show-inheritance:


|RelationshipPolicyError|
-------------------------

.. autoexception:: RelationshipPolicyError
   :show-inheritance:


|DocumentProtectedError|
------------------------

.. autoexception:: DocumentProtectedError
   :show-inheritance:
