"""C++ code detection logic."""

import re
from typing import List

# C++ detection patterns - high confidence (single match = detected)
C_DETECTORS_HIGH: List[str] = [
    r'#include\s*[<"]',
    r"using\s+namespace\s+\w+",
    r"int\s+main\s*\(",
    r"std::",
    r"::\s*\w+\s*\(",
    r"template\s*<",
]

# C++ detection patterns - medium confidence (multiple matches needed)
C_DETECTORS_MEDIUM: List[str] = [
    r"\b(class|struct|enum)\s+\w+",
    r"\b(int|char|float|double|void|bool|auto|const|constexpr|mutable)\b",
    r"^\s*(public|private|protected):",
    r"^\s*#\s*(define|ifdef|ifndef|endif|pragma)",
    r"\b(for|while|if|else|switch|case|break|continue|return)\b",
    r"\b(cout|cin|endl|printf|scanf)\b",
    r"<<|>>",
    r"\b(string|vector|map|set|array)\b",
    r"//",
    r"/\*.*?\*/",
]


def is_cpp_code(
    text: str,
    high_patterns: List[str] = None,
    medium_patterns: List[str] = None,
) -> bool:
    """Detect if text contains C++ code."""
    if not text or not isinstance(text, str):
        return False

    high = high_patterns or C_DETECTORS_HIGH
    medium = medium_patterns or C_DETECTORS_MEDIUM

    high_confidence = 0
    medium_confidence = 0

    for pattern in high:
        matches = re.findall(pattern, text, re.MULTILINE)
        high_confidence += len(matches)

    for pattern in medium:
        flags = re.MULTILINE
        if "/*" in pattern:
            flags |= re.DOTALL
        matches = re.findall(pattern, text, flags)
        medium_confidence += len(matches)

    if high_confidence >= 1:
        return True
    if medium_confidence >= 3:
        return True
    if medium_confidence >= 2 and len(text) > 100:
        return True

    return False
