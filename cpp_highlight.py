#!/usr/bin/env python3
"""
C++ Excel Highlighter

A command-line tool that automatically detects C++ code in Excel cells
and applies syntax highlighting using the Atom One Light color theme.

This file provides backward compatibility for direct script execution.
For package usage, use: python -m cpp_highlight
"""

from cpp_highlight.cli import main

if __name__ == "__main__":
    main()
