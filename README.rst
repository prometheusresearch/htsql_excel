********************
HTSQL_EXCEL Overview
********************

The ``htsql_excel`` package is an extension for `HTSQL`_ that adds basic
support for the Excel file format.

.. _`HTSQL`: http://htsql.org/


Installation
============

Install this package like you would any other Python package::

    $ pip install htsql_excel

Then, enable the ``htsql_excel`` extension when you invoke HTSQL. E.g.::

    $ htsql-ctl -E htsql_excel ...

Or, in your HTSQL configuration YAML file::

    htsql:
      db: ...
    htsql_excel:


Formatters
==========

This extension adds two formatter functions to HTSQL: ``/:xls`` and ``/:xlsx``.
They are tabular formatters (like ``/:csv``) that will output the results in
in either XLS format (the binary format used prior to Excel 2007) or XLSX (the
Office Open XML format introduced with Excel 2007).


License/Copyright
=================

This project is licensed under the GNU Affero General Public License, version
3. See the accompanying ``LICENSE.rst`` file for details.

Copyright (c) 2016, Prometheus Research, LLC

