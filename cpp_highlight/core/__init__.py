"""Core module for ExcelCppSyntaxHighlight."""

from .detection import is_cpp_code, C_DETECTORS_HIGH, C_DETECTORS_MEDIUM
from .highlighter import CellHighlighter, calculate_required_height

__all__ = [
    "is_cpp_code",
    "C_DETECTORS_HIGH",
    "C_DETECTORS_MEDIUM",
    "CellHighlighter",
    "calculate_required_height",
]
