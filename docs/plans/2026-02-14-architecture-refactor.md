# Architecture Refactor (Phase 1-2) Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Refactor ExcelCppSyntaxHighlight from single-file monolith to modular, testable architecture with proper separation of concerns.

**Architecture:** Extract modules following single-responsibility principle: detection, theming, highlighting, and CLI. Use dataclasses for configuration, dependency injection to eliminate global state, and pytest for comprehensive testing.

**Tech Stack:** Python 3.6+, pytest, openpyxl, Pygments, dataclasses

---

## Task 1: Setup Test Infrastructure

**Files:**
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Modify: `requirements.txt`

**Step 1: Add pytest to requirements**

```python
# Add to requirements.txt
pytest>=7.0.0
```

**Step 2: Create tests/__init__.py**

```python
# Empty file for package discovery
```

**Step 3: Create tests/conftest.py with shared fixtures**

```python
"""Shared pytest fixtures for ExcelCppSyntaxHighlight tests."""
import pytest
from pygments.token import Token


@pytest.fixture
def sample_cpp_code():
    """Sample C++ code for testing."""
    return """#include <iostream>

int main() {
    std::cout << "Hello, World!" << std::endl;
    return 0;
}"""


@pytest.fixture
def sample_tokens():
    """Sample tokenized C++ code."""
    return [
        (Token.Comment.Preproc, "#include"),
        (Token.Comment.PreprocFile, " <iostream>"),
        (Token.Text, "\n\n"),
        (Token.Keyword.Type, "int"),
        (Token.Text, " "),
        (Token.Name.Function, "main"),
        (Token.Punctuation, "()"),
        (Token.Text, " {\n    "),
        (Token.Name.Builtin, "std"),
        (Token.Operator, "::"),
        (Token.Name, "cout"),
        (Token.Text, " "),
        (Token.Operator, "<<"),
        (Token.Text, " "),
        (Token.String.Double, '"Hello, World!"'),
        (Token.Text, " "),
        (Token.Operator, "<<"),
        (Token.Text, " "),
        (Token.Name.Builtin, "std"),
        (Token.Operator, "::"),
        (Token.Name, "endl"),
        (Token.Punctuation, ";"),
        (Token.Text, "\n    "),
        (Token.Keyword, "return"),
        (Token.Text, " "),
        (Token.Literal.Number.Integer, "0"),
        (Token.Punctuation, ";"),
        (Token.Text, "\n}"),
    ]
```

**Step 4: Run tests to verify setup**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/ -v`
Expected: "collected 0 items" (no tests yet, but pytest runs)

**Step 5: Commit**

```bash
git add tests/__init__.py tests/conftest.py requirements.txt
git commit -m "test: add pytest infrastructure"
```

---

## Task 2: Add Unit Tests for `is_cpp_code()`

**Files:**
- Create: `tests/test_detection.py`

**Step 1: Write failing tests for high confidence patterns**

```python
"""Tests for C++ code detection logic."""
import pytest


class TestCppDetectionHighConfidence:
    """Tests for high-confidence C++ detection patterns."""

    def test_include_directive(self):
        """#include is high-confidence C++ indicator."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code("#include <vector>") is True
        assert is_cpp_code('#include "myfile.h"') is True

    def test_std_namespace(self):
        """std:: is high-confidence C++ indicator."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code("std::cout << hello;") is True
        assert is_cpp_code("auto x = std::make_unique<int>();") is True

    def test_int_main(self):
        """int main() is high-confidence C++ indicator."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code("int main() { return 0; }") is True
        assert is_cpp_code("int main(int argc, char* argv[])") is True

    def test_template_declaration(self):
        """template< is high-confidence C++ indicator."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code("template<typename T>") is True
        assert is_cpp_code("template<class T, class U>") is True

    def test_using_namespace(self):
        """using namespace is high-confidence C++ indicator."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code("using namespace std;") is True


class TestCppDetectionMediumConfidence:
    """Tests for medium-confidence C++ detection patterns."""

    def test_single_medium_pattern_insufficient(self):
        """Single medium pattern match is insufficient."""
        from cpp_highlight import is_cpp_code
        
        # Only 1 match for 'int'
        assert is_cpp_code("int x") is False

    def test_two_medium_patterns_short_text_insufficient(self):
        """Two medium patterns in short text are insufficient."""
        from cpp_highlight import is_cpp_code
        
        # 2 matches but short text
        assert is_cpp_code("int x; int y") is False

    def test_two_medium_patterns_long_text_sufficient(self):
        """Two medium patterns in long text (>100 chars) are sufficient."""
        from cpp_highlight import is_cpp_code
        
        # 2 matches + long text
        long_code = "int x; int y; " + "x" * 100
        assert is_cpp_code(long_code) is True

    def test_three_medium_patterns_sufficient(self):
        """Three medium pattern matches are sufficient."""
        from cpp_highlight import is_cpp_code
        
        # 3 occurrences of 'int'
        assert is_cpp_code("int x; int y; int z") is True

    def test_class_declaration(self):
        """class declaration counts as medium pattern."""
        from cpp_highlight import is_cpp_code
        
        # 3 patterns: class, int, int
        code = "class MyClass { int x; int y; };"
        assert is_cpp_code(code) is True

    def test_for_loop(self):
        """for loop counts as medium pattern."""
        from cpp_highlight import is_cpp_code
        
        # 3 patterns: for, int, <<
        code = "for (int i = 0; i < 10; ++i) { cout << i; }"
        assert is_cpp_code(code) is True


class TestCppDetectionEdgeCases:
    """Tests for edge cases in C++ detection."""

    def test_empty_string(self):
        """Empty string is not C++ code."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code("") is False

    def test_none_input(self):
        """None is not C++ code."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code(None) is False

    def test_non_string_input(self):
        """Non-string input is not C++ code."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code(123) is False
        assert is_cpp_code(["code"]) is False

    def test_plain_text(self):
        """Plain English text is not detected as C++."""
        from cpp_highlight import is_cpp_code
        
        assert is_cpp_code("Hello, this is just plain text.") is False
        assert is_cpp_code("The quick brown fox jumps.") is False

    def test_code_with_comments(self):
        """C++ with comments is detected."""
        from cpp_highlight import is_cpp_code
        
        # High confidence from #include
        code = "#include <vector>\n// This is a comment"
        assert is_cpp_code(code) is True

    def test_multiline_code(self):
        """Multiline C++ code is detected."""
        from cpp_highlight import is_cpp_code
        
        code = """#include <iostream>
int main() {
    std::cout << "Hello";
    return 0;
}"""
        assert is_cpp_code(code) is True
```

**Step 2: Run tests to verify they pass**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/test_detection.py -v`
Expected: All tests PASS

**Step 3: Commit**

```bash
git add tests/test_detection.py
git commit -m "test: add comprehensive tests for is_cpp_code detection"
```

---

## Task 3: Add Unit Tests for Color Resolution

**Files:**
- Create: `tests/test_theme.py`

**Step 1: Write tests for get_color function**

```python
"""Tests for theme color resolution."""
import pytest
from pygments.token import Token


class TestColorResolution:
    """Tests for token-to-color mapping."""

    def test_exact_keyword_match(self):
        """Exact Keyword token type matches."""
        from cpp_highlight import get_color, TOKEN_TYPE_NAMES
        
        # Verify Keyword is in mapping
        assert Token.Keyword in TOKEN_TYPE_NAMES
        color = get_color(Token.Keyword)
        assert color == "A626A4"

    def test_exact_string_match(self):
        """Exact String token type matches."""
        from cpp_highlight import get_color
        
        color = get_color(Token.String)
        assert color == "50A14F"

    def test_exact_comment_match(self):
        """Exact Comment token type matches."""
        from cpp_highlight import get_color
        
        color = get_color(Token.Comment)
        assert color == "A0A1A7"

    def test_default_fallback(self):
        """Unknown token type returns default color."""
        from cpp_highlight import get_color
        
        # Token.Other is not in default theme, should fall back
        color = get_color(Token.Other)
        assert color == "383A42"

    def test_number_color(self):
        """Number tokens get correct color."""
        from cpp_highlight import get_color
        
        color = get_color(Token.Literal.Number.Integer)
        assert color == "986801"


class TestTokenTypeNames:
    """Tests for token type name mapping."""

    def test_keyword_type_mapping(self):
        """Keyword.Type is mapped correctly."""
        from cpp_highlight import TOKEN_TYPE_NAMES
        
        assert TOKEN_TYPE_NAMES[Token.Keyword.Type] == "Keyword.Type"

    def test_string_single_mapping(self):
        """String.Single is mapped correctly."""
        from cpp_highlight import TOKEN_TYPE_NAMES
        
        assert TOKEN_TYPE_NAMES[Token.String.Single] == "String.Single"

    def test_name_function_mapping(self):
        """Name.Function is mapped correctly."""
        from cpp_highlight import TOKEN_TYPE_NAMES
        
        assert TOKEN_TYPE_NAMES[Token.Name.Function] == "Name.Function"
```

**Step 2: Run tests to verify they pass**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/test_theme.py -v`
Expected: All tests PASS

**Step 3: Commit**

```bash
git add tests/test_theme.py
git commit -m "test: add tests for theme color resolution"
```

---

## Task 4: Extract FontSettings Dataclass

**Files:**
- Create: `cpp_highlight/config/__init__.py`
- Create: `cpp_highlight/config/settings.py`
- Modify: `cpp_highlight.py` (will import from config)

**Step 1: Create config package**

```python
# cpp_highlight/config/__init__.py
"""Configuration module for ExcelCppSyntaxHighlight."""
from .settings import FontSettings

__all__ = ["FontSettings"]
```

**Step 2: Create FontSettings dataclass**

```python
# cpp_highlight/config/settings.py
"""Font and layout settings for syntax highlighting."""
from dataclasses import dataclass


@dataclass
class FontSettings:
    """Configuration for code font in Excel cells.
    
    Attributes:
        name: Font family name (e.g., 'Consolas', 'Courier New')
        size: Font size in points
        base_height: Base row height in points for single-line code
        line_height: Additional height per line in points
    """
    name: str = "Consolas"
    size: int = 11
    base_height: float = 16.0
    line_height: float = 16.0

    @classmethod
    def default(cls) -> "FontSettings":
        """Create default font settings."""
        return cls()
```

**Step 3: Update cpp_highlight.py to use FontSettings**

In `highlight_cell()` function, replace hardcoded font values:

```python
# OLD (line 284):
font = InlineFont(color=color_obj, rFont="Consolas", sz=11)

# NEW:
from cpp_highlight.config import FontSettings
_font_settings = FontSettings.default()
font = InlineFont(color=color_obj, rFont=_font_settings.name, sz=_font_settings.size)
```

In `calculate_required_height()` function:

```python
# OLD (lines 322-323):
base_height = 16.0
line_height = 16.0

# NEW:
from cpp_highlight.config import FontSettings
_settings = FontSettings.default()
base_height = _settings.base_height
line_height = _settings.line_height
```

**Step 4: Run tests to verify no regression**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/ -v`
Expected: All tests PASS

**Step 5: Commit**

```bash
git add cpp_highlight/config/__init__.py cpp_highlight/config/settings.py cpp_highlight.py
git commit -m "refactor: extract FontSettings dataclass for font configuration"
```

---

## Task 5: Add Unit Tests for Height Calculation

**Files:**
- Create: `tests/test_height.py`

**Step 1: Write tests for calculate_required_height**

```python
"""Tests for row height calculation."""
import pytest
from pygments.token import Token


class TestHeightCalculation:
    """Tests for calculate_required_height function."""

    def test_single_line(self):
        """Single line of code returns base height."""
        from cpp_highlight import calculate_required_height
        
        tokens = [(Token.Text, "int x = 42;")]
        height = calculate_required_height(tokens)
        assert height == 16.0  # base_height

    def test_two_lines(self):
        """Two lines of code returns base + line height."""
        from cpp_highlight import calculate_required_height
        
        tokens = [
            (Token.Text, "int x = 1;\n"),
            (Token.Text, "int y = 2;"),
        ]
        height = calculate_required_height(tokens)
        assert height == 32.0  # base_height + line_height

    def test_three_lines(self):
        """Three lines of code returns correct height."""
        from cpp_highlight import calculate_required_height
        
        tokens = [
            (Token.Text, "line1\n"),
            (Token.Text, "line2\n"),
            (Token.Text, "line3"),
        ]
        height = calculate_required_height(tokens)
        assert height == 48.0  # base + 2 * line

    def test_multiple_newlines_in_single_token(self):
        """Multiple newlines in single token are counted."""
        from cpp_highlight import calculate_required_height
        
        tokens = [(Token.Text, "line1\nline2\nline3\nline4")]
        height = calculate_required_height(tokens)
        assert height == 64.0  # base + 3 * line

    def test_empty_tokens(self):
        """Empty tokens list returns base height."""
        from cpp_highlight import calculate_required_height
        
        height = calculate_required_height([])
        assert height == 16.0

    def test_trailing_newline_ignored(self):
        """The trailing newline token handling is applied before this function."""
        from cpp_highlight import calculate_required_height
        
        # Note: In actual usage, trailing \n token is removed before calling
        tokens = [(Token.Text, "code")]
        height = calculate_required_height(tokens)
        assert height == 16.0
```

**Step 2: Run tests to verify they pass**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/test_height.py -v`
Expected: All tests PASS

**Step 3: Commit**

```bash
git add tests/test_height.py
git commit -m "test: add tests for row height calculation"
```

---

## Task 6: Create ThemeConfig Class

**Files:**
- Create: `cpp_highlight/config/theme.py`
- Modify: `cpp_highlight/config/__init__.py`

**Step 1: Create ThemeConfig class**

```python
# cpp_highlight/config/theme.py
"""Theme configuration for syntax highlighting."""
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional

from pygments.token import Token


# Token type to string mapping for JSON config lookup
TOKEN_TYPE_NAMES: Dict[Token, str] = {
    Token.Keyword: "Keyword",
    Token.Keyword.Type: "Keyword.Type",
    Token.Keyword.Namespace: "Keyword.Namespace",
    Token.Keyword.Constant: "Keyword.Constant",
    Token.Keyword.Declaration: "Keyword.Declaration",
    Token.String: "String",
    Token.String.Single: "String.Single",
    Token.String.Double: "String.Double",
    Token.String.Doc: "String.Doc",
    Token.Comment: "Comment",
    Token.Comment.Single: "Comment.Single",
    Token.Comment.Multiline: "Comment.Multiline",
    Token.Comment.Preproc: "Comment.Preproc",
    Token.Comment.PreprocFile: "Comment.PreprocFile",
    Token.Number: "Number",
    Token.Number.Integer: "Number.Integer",
    Token.Number.Float: "Number.Float",
    Token.Number.Hex: "Number.Hex",
    Token.Number.Bin: "Number.Bin",
    Token.Number.Oct: "Number.Oct",
    Token.Name.Function: "Name.Function",
    Token.Name.Class: "Name.Class",
    Token.Name.Namespace: "Name.Namespace",
    Token.Name.Builtin: "Name.Builtin",
    Token.Name.Builtin.Pseudo: "Name.Builtin.Pseudo",
    Token.Name.Decorator: "Name.Decorator",
    Token.Name: "Name",
    Token.Name.Attribute: "Name.Attribute",
    Token.Name.Tag: "Name.Tag",
    Token.Name.Variable: "Name.Variable",
    Token.Name.Variable.Class: "Name.Variable.Class",
    Token.Name.Variable.Global: "Name.Variable.Global",
    Token.Name.Variable.Instance: "Name.Variable.Instance",
    Token.Name.Variable.Magic: "Name.Variable.Magic",
    Token.Name.Other: "Name.Other",
    Token.Operator: "Operator",
    Token.Operator.Word: "Operator.Word",
    Token.Punctuation: "Punctuation",
    Token.Text: "Text",
    Token.Text.Whitespace: "Text.Whitespace",
    Token.Error: "Error",
    Token.Other: "Other",
}

# Default theme colors (Atom One Light)
DEFAULT_THEME_COLORS: Dict[str, str] = {
    "Keyword": "A626A4",
    "Keyword.Type": "A626A4",
    "Keyword.Namespace": "A626A4",
    "Keyword.Constant": "A626A4",
    "Keyword.Declaration": "A626A4",
    "String": "50A14F",
    "String.Single": "50A14F",
    "String.Double": "50A14F",
    "String.Doc": "A0A1A7",
    "Comment": "A0A1A7",
    "Comment.Single": "A0A1A7",
    "Comment.Multiline": "A0A1A7",
    "Comment.Preproc": "A626A4",
    "Comment.PreprocFile": "A626A4",
    "Number": "986801",
    "Number.Integer": "986801",
    "Number.Float": "986801",
    "Number.Hex": "986801",
    "Number.Bin": "986801",
    "Number.Oct": "986801",
    "Name.Function": "4078F2",
    "Name.Class": "C18401",
    "Name.Namespace": "C18401",
    "Name.Builtin": "C18401",
    "Name.Builtin.Pseudo": "C18401",
    "Name.Decorator": "C18401",
    "Name": "E45649",
    "Name.Attribute": "E45649",
    "Name.Tag": "E45649",
    "Name.Variable": "E45649",
    "Name.Variable.Class": "E45649",
    "Name.Variable.Global": "E45649",
    "Name.Variable.Instance": "E45649",
    "Name.Variable.Magic": "E45649",
    "Name.Other": "E45649",
    "Operator": "383A42",
    "Operator.Word": "A626A4",
    "Punctuation": "383A42",
    "Text": "383A42",
    "Text.Whitespace": "383A42",
    "Error": "FF0000",
    "Other": "383A42",
}


def get_config_path() -> Path:
    """Get the path to the theme configuration file."""
    if getattr(sys, "frozen", False):
        # Running in a bundled executable
        base_path = Path(sys.executable).parent
    else:
        # Running as a Python script
        base_path = Path(__file__).parent.parent.parent
    return base_path / "theme.json"


@dataclass
class ThemeConfig:
    """Theme configuration for syntax highlighting.
    
    Attributes:
        colors: Mapping from token type names to hex color codes
        token_names: Mapping from Token types to config string names
    """
    colors: Dict[str, str] = field(default_factory=lambda: DEFAULT_THEME_COLORS.copy())
    token_names: Dict[Token, str] = field(default_factory=lambda: TOKEN_TYPE_NAMES.copy())
    default_color: str = "383A42"

    @classmethod
    def from_json(cls, path: Optional[Path] = None) -> "ThemeConfig":
        """Load theme configuration from JSON file.
        
        Args:
            path: Path to theme.json. If None, uses default location.
            
        Returns:
            ThemeConfig instance with loaded colors.
        """
        if path is None:
            path = get_config_path()
        
        colors = DEFAULT_THEME_COLORS.copy()
        
        if not path.exists():
            # Create default config file
            cls._create_default_config(path)
        else:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    colors.update(loaded)
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Failed to load theme config: {e}", file=sys.stderr)
        
        return cls(colors=colors)

    @staticmethod
    def _create_default_config(path: Path) -> None:
        """Create default theme configuration file."""
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_THEME_COLORS, f, indent=2)
        except IOError as e:
            print(f"Warning: Failed to create default config: {e}", file=sys.stderr)

    def get_color(self, token_type: Token) -> str:
        """Get color for a token type.
        
        Args:
            token_type: Pygments token type
            
        Returns:
            Hex color code (6 characters, no # prefix)
        """
        # Try exact match first
        type_name = self.token_names.get(token_type)
        if type_name and type_name in self.colors:
            return self.colors[type_name]
        
        # Try parent type matching
        for ttype, name in self.token_names.items():
            if token_type in ttype and name in self.colors:
                return self.colors[name]
        
        return self.default_color
```

**Step 2: Update config/__init__.py**

```python
# cpp_highlight/config/__init__.py
"""Configuration module for ExcelCppSyntaxHighlight."""
from .settings import FontSettings
from .theme import ThemeConfig, TOKEN_TYPE_NAMES, DEFAULT_THEME_COLORS

__all__ = [
    "FontSettings",
    "ThemeConfig",
    "TOKEN_TYPE_NAMES",
    "DEFAULT_THEME_COLORS",
]
```

**Step 3: Run tests to verify**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/ -v`
Expected: All tests PASS

**Step 4: Commit**

```bash
git add cpp_highlight/config/theme.py cpp_highlight/config/__init__.py
git commit -m "refactor: create ThemeConfig class for theme management"
```

---

## Task 7: Create Detection Module

**Files:**
- Create: `cpp_highlight/core/__init__.py`
- Create: `cpp_highlight/core/detection.py`

**Step 1: Create core package**

```python
# cpp_highlight/core/__init__.py
"""Core module for ExcelCppSyntaxHighlight."""
from .detection import is_cpp_code, C_DETECTORS_HIGH, C_DETECTORS_MEDIUM

__all__ = [
    "is_cpp_code",
    "C_DETECTORS_HIGH",
    "C_DETECTORS_MEDIUM",
]
```

**Step 2: Create detection module**

```python
# cpp_highlight/core/detection.py
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
    r"//",  # Single line comment
    r"/\*.*?\*/",  # Multi-line comment (non-greedy)
]


def is_cpp_code(
    text: str,
    high_patterns: List[str] = None,
    medium_patterns: List[str] = None,
) -> bool:
    """Detect if text contains C++ code.
    
    Uses a two-tier confidence scoring system:
    - High confidence patterns: 1+ match = detected
    - Medium confidence patterns: 3+ matches or 2+ with long text = detected
    
    Args:
        text: Text to analyze
        high_patterns: Optional custom high-confidence patterns
        medium_patterns: Optional custom medium-confidence patterns
        
    Returns:
        True if text is detected as C++ code
    """
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
```

**Step 3: Run tests to verify**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/test_detection.py -v`
Expected: Tests need update to import from new module

**Step 4: Update tests to use new module path**

```python
# In tests/test_detection.py, update imports:
# OLD:
from cpp_highlight import is_cpp_code

# NEW:
from cpp_highlight.core.detection import is_cpp_code
```

**Step 5: Run tests again**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/test_detection.py -v`
Expected: All tests PASS

**Step 6: Commit**

```bash
git add cpp_highlight/core/__init__.py cpp_highlight/core/detection.py tests/test_detection.py
git commit -m "refactor: extract detection logic to core.detection module"
```

---

## Task 8: Create Highlighter Module

**Files:**
- Create: `cpp_highlight/core/highlighter.py`
- Modify: `cpp_highlight/core/__init__.py`

**Step 1: Create highlighter module**

```python
# cpp_highlight/core/highlighter.py
"""Cell highlighting logic."""
from typing import List, Tuple, Optional

from openpyxl.cell.rich_text import CellRichText
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Color
from pygments import lex
from pygments.lexers import CppLexer
from pygments.token import Token

from cpp_highlight.config import FontSettings, ThemeConfig
from cpp_highlight.models import TextBlock


def calculate_required_height(
    tokens: List[Tuple[Token, str]],
    font_settings: FontSettings = None,
) -> float:
    """Calculate the required row height for given tokens.

    Args:
        tokens: List of (token_type, value) tuples
        font_settings: Font configuration for height calculation

    Returns:
        Required row height in points.
    """
    if font_settings is None:
        font_settings = FontSettings.default()
    
    # Count lines in the content
    line_count = 1
    for token_type, value in tokens:
        if "\n" in value:
            line_count += value.count("\n")

    return font_settings.base_height + (line_count - 1) * font_settings.line_height


class CellHighlighter:
    """Highlights Excel cells with syntax-highlighted C++ code."""

    def __init__(
        self,
        theme: ThemeConfig = None,
        font: FontSettings = None,
        lexer: type = None,
    ):
        """Initialize the highlighter.
        
        Args:
            theme: Theme configuration for colors
            font: Font settings for text formatting
            lexer: Pygments lexer class to use
        """
        self.theme = theme or ThemeConfig.from_json()
        self.font = font or FontSettings.default()
        self.lexer = (lexer or CppLexer)()

    def highlight(self, text: str) -> Tuple[Optional[CellRichText], Optional[float]]:
        """Apply syntax highlighting to C++ code text.
        
        Args:
            text: C++ code string to highlight
            
        Returns:
            Tuple of (CellRichText or None, required_height or None)
        """
        try:
            tokens = list(lex(text, self.lexer))

            # Remove trailing newline token added by CppLexer
            if tokens and tokens[-1][0] == Token.Text.Whitespace and tokens[-1][1] == "\n":
                tokens = tokens[:-1]

            blocks = []
            for token_type, value in tokens:
                color_hex = self.theme.get_color(token_type)
                color_obj = Color(rgb=color_hex)
                font = InlineFont(
                    color=color_obj,
                    rFont=self.font.name,
                    sz=self.font.size,
                )
                block = TextBlock(text=value, font=font)
                blocks.append(block)

            rich_text = CellRichText(*blocks)
            required_height = calculate_required_height(tokens, self.font)

            return rich_text, required_height

        except Exception as e:
            print(f"Warning: Failed to highlight text: {e}", file=__import__("sys").stderr)
            return None, None

    def apply_to_cell(self, cell) -> bool:
        """Apply syntax highlighting to an Excel cell.
        
        Args:
            cell: openpyxl cell object
            
        Returns:
            True if highlighting was applied, False otherwise
        """
        text = cell.value
        if not isinstance(text, str):
            return False
        
        rich_text, required_height = self.highlight(text)
        
        if rich_text is None:
            return False
        
        cell.value = rich_text
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        
        return True
```

**Step 2: Create models module**

```python
# cpp_highlight/models/__init__.py
"""Models module for ExcelCppSyntaxHighlight."""
from .text_block import TextBlock

__all__ = ["TextBlock"]
```

```python
# cpp_highlight/models/text_block.py
"""Custom TextBlock for whitespace preservation."""
from xml.etree.ElementTree import Element

from openpyxl.cell.rich_text import TextBlock as _OrigTextBlock


class TextBlock(_OrigTextBlock):
    """TextBlock that properly preserves whitespace in Excel."""

    def to_tree(self):
        """Convert to XML tree, ensuring whitespace is preserved."""
        el = Element("r")
        if self.font:
            el.append(self.font.to_tree(tagname="rPr"))
        t = Element("t")
        t.text = self.text

        # Always set xml:space="preserve" if text contains any whitespace
        if self.text and (self.text != self.text.strip() or not self.text.strip()):
            XML_NS = "http://www.w3.org/XML/1998/namespace"
            t.set("{%s}space" % XML_NS, "preserve")

        el.append(t)
        return el
```

**Step 3: Update core/__init__.py**

```python
# cpp_highlight/core/__init__.py
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
```

**Step 4: Run tests to verify**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/ -v`
Expected: All tests PASS

**Step 5: Commit**

```bash
git add cpp_highlight/core/highlighter.py cpp_highlight/models/__init__.py cpp_highlight/models/text_block.py cpp_highlight/core/__init__.py
git commit -m "refactor: extract CellHighlighter class to core.highlighter module"
```

---

## Task 9: Update Main Module to Use New Architecture

**Files:**
- Create: `cpp_highlight/__init__.py`
- Create: `cpp_highlight/__main__.py`
- Create: `cpp_highlight/cli.py`
- Create: `cpp_highlight/processor.py`
- Modify: Update imports in all modules

**Step 1: Create package __init__.py**

```python
# cpp_highlight/__init__.py
"""ExcelCppSyntaxHighlight - C++ syntax highlighting for Excel cells."""
__version__ = "1.0.0"

from cpp_highlight.config import FontSettings, ThemeConfig
from cpp_highlight.core import (
    is_cpp_code,
    CellHighlighter,
    calculate_required_height,
)

__all__ = [
    "__version__",
    "FontSettings",
    "ThemeConfig",
    "is_cpp_code",
    "CellHighlighter",
    "calculate_required_height",
]
```

**Step 2: Create __main__.py**

```python
# cpp_highlight/__main__.py
"""Entry point for python -m cpp_highlight."""
from cpp_highlight.cli import main

if __name__ == "__main__":
    main()
```

**Step 3: Create CLI module**

```python
# cpp_highlight/cli.py
"""Command-line interface for ExcelCppSyntaxHighlight."""
import argparse
import sys
from pathlib import Path

from cpp_highlight.processor import process_excel


def main():
    """Main entry point for CLI."""
    parser = argparse.ArgumentParser(
        description="Apply C++ syntax highlighting to Excel cells",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m cpp_highlight input.xlsx -o output.xlsx
  python -m cpp_highlight code.xlsx --output highlighted.xlsx --verbose
        """,
    )

    parser.add_argument("input", help="Input Excel file path")
    parser.add_argument("-o", "--output", required=True, help="Output Excel file path")
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable verbose output"
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if not input_path.suffix.lower() in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
        print(
            f"Warning: Input file may not be a valid Excel file: {args.input}",
            file=sys.stderr,
        )

    count = process_excel(args.input, args.output, args.verbose)

    print(f"Processed {count} cells with C++ code")
    print(f"  Input:  {args.input}")
    print(f"  Output: {args.output}")


if __name__ == "__main__":
    main()
```

**Step 4: Create processor module**

```python
# cpp_highlight/processor.py
"""Excel file processing logic."""
import sys

import openpyxl

from cpp_highlight.config import FontSettings, ThemeConfig
from cpp_highlight.core import CellHighlighter, is_cpp_code


def process_excel(input_path: str, output_path: str, verbose: bool = False) -> int:
    """Process an Excel file and apply C++ syntax highlighting.

    Args:
        input_path: Path to input Excel file
        output_path: Path to output Excel file
        verbose: Enable verbose output

    Returns:
        Number of cells highlighted
    """
    if verbose:
        print(f"Loading: {input_path}")

    try:
        wb = openpyxl.load_workbook(input_path)
    except Exception as e:
        print(f"Error: Failed to load workbook: {e}", file=sys.stderr)
        sys.exit(1)

    highlighter = CellHighlighter()
    highlighted_count = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        if verbose:
            print(f"\nProcessing sheet: {sheet_name}")

        # Track row height requirements: {row_number: max_required_height}
        row_height_requirements = {}

        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and is_cpp_code(cell.value):
                    if verbose:
                        print(f"  {cell.coordinate}: Detected C++ code")

                    if highlighter.apply_to_cell(cell):
                        highlighted_count += 1
                        if verbose:
                            print(f"    -> Highlighted")

                        # Calculate and track required height
                        rich_text, required_height = highlighter.highlight(cell.value.value if hasattr(cell.value, '__iter__') else str(cell.value))
                        if required_height is not None:
                            row_num = cell.row
                            current_max = row_height_requirements.get(row_num, 0)
                            row_height_requirements[row_num] = max(
                                current_max, required_height
                            )

        # Apply row heights
        for row_num, required_height in row_height_requirements.items():
            original_height = ws.row_dimensions[row_num].height

            if original_height is None:
                ws.row_dimensions[row_num].height = required_height
            else:
                ws.row_dimensions[row_num].height = max(original_height, required_height)

    if verbose:
        print(f"\nSaving: {output_path}")

    try:
        wb.save(output_path)
    except Exception as e:
        print(f"Error: Failed to save workbook: {e}", file=sys.stderr)
        sys.exit(1)

    return highlighted_count
```

**Step 5: Run tests to verify**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/ -v`
Expected: All tests PASS

**Step 6: Test new package entry point**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m cpp_highlight --help`
Expected: Help message displayed

**Step 7: Commit**

```bash
git add cpp_highlight/__init__.py cpp_highlight/__main__.py cpp_highlight/cli.py cpp_highlight/processor.py
git commit -m "refactor: create modular package structure with CLI and processor modules"
```

---

## Task 10: Update Backward Compatibility and Finalize

**Files:**
- Modify: `cpp_highlight.py` (keep as backward-compatible entry point)

**Step 1: Update cpp_highlight.py as backward-compatible wrapper**

```python
#!/usr/bin/env python3
"""
C++ Excel Highlighter

A command-line tool that automatically detects C++ code in Excel cells
and applies syntax highlighting using the Atom One Light color theme.

This file provides backward compatibility for direct script execution.
For package usage, use: python -m cpp_highlight
"""

# Re-export all public API for backward compatibility
from cpp_highlight import (
    __version__,
    FontSettings,
    ThemeConfig,
    is_cpp_code,
    CellHighlighter,
    calculate_required_height,
)
from cpp_highlight.config import TOKEN_TYPE_NAMES, DEFAULT_THEME_COLORS
from cpp_highlight.core import C_DETECTORS_HIGH, C_DETECTORS_MEDIUM
from cpp_highlight.models import TextBlock
from cpp_highlight.cli import main

# Legacy function aliases for backward compatibility
get_color = None  # Will be set below

def _init_legacy():
    """Initialize legacy compatibility layer."""
    global get_color
    _theme = ThemeConfig.from_json()
    get_color = _theme.get_color

_init_legacy()

if __name__ == "__main__":
    main()
```

**Step 2: Run full test suite**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/ -v --tb=short`
Expected: All tests PASS

**Step 3: Test backward compatibility**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python cpp_highlight.py --help`
Expected: Help message displayed

**Step 4: Test new package entry point**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m cpp_highlight --help`
Expected: Help message displayed

**Step 5: Commit**

```bash
git add cpp_highlight.py
git commit -m "refactor: update main script as backward-compatible wrapper"
```

---

## Task 11: Add Integration Tests

**Files:**
- Create: `tests/test_integration.py`

**Step 1: Create integration test file**

```python
"""Integration tests for ExcelCppSyntaxHighlight."""
import pytest
import tempfile
from pathlib import Path

import openpyxl
from openpyxl import Workbook

from cpp_highlight import (
    is_cpp_code,
    CellHighlighter,
    ThemeConfig,
    FontSettings,
)
from cpp_highlight.processor import process_excel


class TestIntegration:
    """End-to-end integration tests."""

    @pytest.fixture
    def sample_workbook(self, tmp_path):
        """Create a sample workbook with C++ code."""
        wb = Workbook()
        ws = wb.active
        
        # Cell with C++ code
        ws["A1"] = """#include <iostream>
int main() {
    std::cout << "Hello" << std::endl;
    return 0;
}"""
        
        # Cell with plain text
        ws["A2"] = "This is just plain text"
        
        # Cell with another C++ snippet
        ws["A3"] = "std::vector<int> numbers = {1, 2, 3};"
        
        path = tmp_path / "input.xlsx"
        wb.save(path)
        return path

    def test_process_excel_highlights_cpp_cells(self, sample_workbook, tmp_path):
        """Processing Excel file highlights C++ cells."""
        output_path = tmp_path / "output.xlsx"
        
        count = process_excel(str(sample_workbook), str(output_path), verbose=False)
        
        assert count == 2  # A1 and A3 are C++ code
        
        # Verify output file exists and is valid
        assert output_path.exists()
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active
        
        # A1 should have rich text (highlighted)
        from openpyxl.cell.rich_text import CellRichText
        assert isinstance(ws["A1"].value, CellRichText)
        
        # A2 should still be plain text
        assert ws["A2"].value == "This is just plain text"

    def test_cell_highlighter_with_custom_theme(self):
        """CellHighlighter works with custom theme."""
        custom_colors = ThemeConfig(colors={"Keyword": "FF0000"})
        highlighter = CellHighlighter(theme=custom_colors)
        
        rich_text, height = highlighter.highlight("int x = 42;")
        
        assert rich_text is not None
        assert height is not None
        assert height > 0

    def test_cell_highlighter_with_custom_font(self):
        """CellHighlighter works with custom font settings."""
        custom_font = FontSettings(name="Courier New", size=14)
        highlighter = CellHighlighter(font=custom_font)
        
        rich_text, height = highlighter.highlight("int x = 42;")
        
        assert rich_text is not None

    def test_empty_workbook(self, tmp_path):
        """Processing empty workbook produces valid output."""
        wb = Workbook()
        ws = wb.active
        ws["A1"] = ""
        
        input_path = tmp_path / "empty.xlsx"
        output_path = tmp_path / "output.xlsx"
        wb.save(input_path)
        
        count = process_excel(str(input_path), str(output_path))
        
        assert count == 0
        assert output_path.exists()

    def test_multiline_code_height_calculation(self):
        """Multiline code gets correct height."""
        highlighter = CellHighlighter()
        
        code = "line1\nline2\nline3\nline4\nline5"
        _, height = highlighter.highlight(code)
        
        # 5 lines = base_height + 4 * line_height
        expected = 16.0 + 4 * 16.0
        assert height == expected
```

**Step 2: Run integration tests**

Run: `cd /mnt/d/GitRepos/ExcelCppSyntaxHighlight && python -m pytest tests/test_integration.py -v`
Expected: All tests PASS

**Step 3: Commit**

```bash
git add tests/test_integration.py
git commit -m "test: add integration tests for end-to-end functionality"
```

---

## Summary

After completing all tasks, the project structure will be:

```
ExcelCppSyntaxHighlight/
├── cpp_highlight.py              # Backward-compatible entry point
├── cpp_highlight/                # Package directory
│   ├── __init__.py              # Package exports
│   ├── __main__.py              # python -m entry point
│   ├── cli.py                   # CLI handling
│   ├── processor.py             # Excel processing
│   ├── config/
│   │   ├── __init__.py
│   │   ├── settings.py          # FontSettings dataclass
│   │   └── theme.py             # ThemeConfig class
│   ├── core/
│   │   ├── __init__.py
│   │   ├── detection.py         # C++ detection
│   │   └── highlighter.py       # CellHighlighter
│   └── models/
│       ├── __init__.py
│       └── text_block.py        # Custom TextBlock
├── tests/
│   ├── __init__.py
│   ├── conftest.py              # Shared fixtures
│   ├── test_detection.py        # Detection unit tests
│   ├── test_theme.py            # Theme tests
│   ├── test_height.py           # Height calculation tests
│   └── test_integration.py      # End-to-end tests
├── theme.json
├── requirements.txt
└── README.md
```

**Benefits achieved:**
- ✅ Comprehensive test coverage
- ✅ Separation of concerns
- ✅ No global state (dependency injection)
- ✅ Configurable font and theme settings
- ✅ Backward compatible
- ✅ Extensible for future multi-language support
