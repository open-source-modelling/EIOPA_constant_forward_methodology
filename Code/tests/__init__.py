"""Unit tests for calculation.py, validation.py and data_io.py.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code

This folder is a package so that both runners put the Code folder on the import path:
the tests import the modules as ``calculation``, ``validation`` and ``data_io``, and share the
workbook example as ``tests.test_calculation.example_inputs``.
"""
