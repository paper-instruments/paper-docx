
.. _paper_bookmarks_api:

Bookmarks
=========

*paper-docx addition.* Enumerate bookmarks, create one that wraps exactly a
|Span|, or delete one by name. Ids are globally unique. Deletion keeps the text
and refuses while a cross-reference still points at the name.

.. currentmodule:: docx.bookmarks


.. autofunction:: list_bookmarks

.. autofunction:: create_bookmark

.. autofunction:: delete_bookmark


|BookmarkInfo| objects
----------------------

.. autoclass:: BookmarkInfo()
   :members:
   :undoc-members:
   :member-order: bysource
