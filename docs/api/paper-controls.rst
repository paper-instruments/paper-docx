
.. _paper_controls_api:

Content controls
================

Enumerate structured document tags (content controls) and set their
values type-correctly. Filling a control that still shows placeholder
text clears the placeholder state so Word treats it as genuinely
filled; data-bound, nested and unsupported controls refuse.

.. currentmodule:: docx.controls


.. autofunction:: list_controls

.. autofunction:: iter_controls

.. autofunction:: get_control

.. autofunction:: set_control_value


|ControlInfo| objects
---------------------

.. autoclass:: ControlInfo()
   :members:
   :undoc-members:
   :member-order: bysource


|Control| objects
-----------------

.. autoclass:: Control()
   :members:
   :member-order: bysource
