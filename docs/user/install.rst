.. _install:

Installing
==========

Install ``paper-docx`` from PyPI and continue importing ``docx``. The
distribution name changed; the Python import name did not.

Migrating from python-docx
--------------------------

``paper-docx`` and ``python-docx`` cannot coexist in one environment because
both distributions own the same ``docx`` package files. Pip does not model
that conflict. Remove both distributions before installing the fork::

    python -m pip uninstall -y python-docx paper-docx
    python -m pip install paper-docx

Confirm that the fork, rather than the upstream distribution, owns the import::

    paper-docx-doctor

The doctor verifies the installed ``docx`` files against ``paper-docx``'s wheel
record and checks ``docx.__paper_version__``. Existing imports remain unchanged::

    import docx
    document = docx.Document("contract.docx")

Pip does not treat the fork as satisfying a dependency declared on
``python-docx``. A package with ``Requires-Dist: python-docx`` will reinstall
upstream and overwrite shared files. Replace or remove that dependency, or run
the package in a separate environment.

Controlled deployments can place this line in a constraints file::

    python-docx<0

Apply that constraint to every pip install in the environment. Pip will then
reject direct and transitive attempts to install upstream.

Dependencies
------------

* Python 3.9 or later
* lxml 3.1.0 or later
* typing-extensions 4.9.0 or later
