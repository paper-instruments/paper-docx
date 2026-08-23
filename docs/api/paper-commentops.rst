
.. _paper_commentops_api:

Comment operations
==================

*paper-docx addition.* Work with the comment threads Word writes in the w15
extended vocabulary. Read a comment's anchored text, walk reply threads, reply
to a comment, and mark a thread resolved or reopened.

A reply has its own ``commentRangeStart``, ``commentRangeEnd``, and
``commentReference`` around the same selected text. Its comment-body paragraph
has its own paragraph ID, while ``w15:paraIdParent`` in ``commentsExtended``
stores the parent comment's final paragraph ID. The reply does not reuse the
parent's marker ID. Consequently, adding a reply intentionally changes the main
document story as well as the comment parts.

.. currentmodule:: docx.commentops


.. autofunction:: is_resolved

.. autofunction:: resolve

.. autofunction:: reply

.. autofunction:: anchored_text

.. autofunction:: comment_thread

.. autofunction:: parent_of
