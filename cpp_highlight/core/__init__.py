"""Core module for ExcelCppSyntaxHighlight."""

from .detection import is_cpp_code, C_DETECTORS_HIGH, C_DETECTORS_MEDIUM

__all__ = [
    "is_cpp_code",
    "C_DETECTORS_HIGH",
    "C_DETECTORS_MEDIUM",
]
