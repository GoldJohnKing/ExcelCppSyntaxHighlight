#!/usr/bin/env python3
"""ExcelCppSyntaxHighlight - C++ syntax highlighting for Excel cells."""

__version__ = "1.0.0"

from cpp_highlight.config import FontSettings, ThemeConfig
from cpp_highlight.core import (
    is_cpp_code,
    CellHighlighter,
    calculate_required_height,
)
from cpp_highlight.core.detection import C_DETECTORS_HIGH, C_DETECTORS_MEDIUM
from cpp_highlight.models import TextBlock


# Legacy function for backward compatibility
def highlight_cell(cell):
    """Apply syntax highlighting to a cell containing C++ code.

    Returns:
        tuple: (success: bool, required_height: float or None)
    """
    from openpyxl.styles import Alignment

    text = cell.value
    if not isinstance(text, str):
        return False, None

    if not is_cpp_code(text):
        return False, None

    highlighter = CellHighlighter()
    rich_text, required_height = highlighter.highlight(text)

    if rich_text is None:
        return False, None

    cell.value = rich_text
    cell.alignment = Alignment(wrap_text=True, vertical="top")

    return True, required_height


# Legacy import for backward compatibility
from cpp_highlight.config.theme import TOKEN_TYPE_NAMES, DEFAULT_THEME_COLORS


__all__ = [
    "__version__",
    "FontSettings",
    "ThemeConfig",
    "is_cpp_code",
    "CellHighlighter",
    "calculate_required_height",
    "highlight_cell",
    "TextBlock",
    "C_DETECTORS_HIGH",
    "C_DETECTORS_MEDIUM",
    "TOKEN_TYPE_NAMES",
    "DEFAULT_THEME_COLORS",
]
