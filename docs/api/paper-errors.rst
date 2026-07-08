
.. _paper_errors_api:

Refusal hierarchy
=================

Every mutating paper-docx API validates fully *before* it touches the
document, so a raised |PaperRefusal| guarantees the in-memory tree and
any file on disk are exactly as they were. Programmer mistakes still
raise the ordinary :exc:`TypeError`/:exc:`ValueError`; a |PaperRefusal|
is a *safe* refusal, catchable distinctly from a bug.

.. currentmodule:: docx.errors



|PaperRefusal|
--------------

.. autoexception:: PaperRefusal
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
