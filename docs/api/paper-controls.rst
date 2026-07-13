
.. _paper_controls_api:

Content controls
================

*paper-docx addition.* Enumerate structured document tags (content controls)
and set their values with the correct type. Filling a control that still shows
placeholder text clears that state so Word treats it as filled.
Data-bound, nested, and unsupported controls refuse.

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
