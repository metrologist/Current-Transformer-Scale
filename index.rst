.. Transformer Buildup documentation master file, created by
   sphinx-quickstart on Wed May 30 15:48:37 2018.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to Transformer Buildup's documentation
===============================================

Contents:

.. toctree::
   :maxdepth: 2

Introduction
=======
.. include:: preface.txt

Modules
=======

twostage
--------

.. automodule:: twostage
.. autoclass:: TWOSTAGE
   :members:

ctscale_mod
-----------

.. automodule:: ctscale_mod
.. autoclass:: BUILDUP
   :members:
   
ExcelPython 
-----------

.. automodule:: ExcelPython
.. autoclass:: CALCULATOR
   :members:

Report Generation
=================
Process
-------
.. include:: CalReport.txt   

Code
====

twostage.py
-----------
.. include:: twostage.txt
   
.. literalinclude:: twostage.py
   :language: python
   :linenos:
   :lines: 1-560

main_ct_scale.py
----------------
.. include:: main_ct_scale.txt      
.. literalinclude:: main_ct_scale.py
   :language: python
   :linenos:
   :lines: 1-220

ctscale_mod.py
----------------
.. include:: ctscale_mod.txt      
.. literalinclude:: ctscale_mod.py
   :language: python
   :linenos:
   :lines: 1-235

ExcelPython.py
----------------
.. include:: ExcelPython.txt      
.. literalinclude:: ExcelPython.py
   :language: python
   :linenos:
   :lines: 1-153

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

