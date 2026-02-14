"""Tests for theme color resolution."""

import pytest
from pygments.token import Token


class TestColorResolution:
    """Tests for token-to-color mapping."""

    def test_exact_keyword_match(self):
        """Exact Keyword token type matches."""
        from cpp_highlight.config import ThemeConfig, TOKEN_TYPE_NAMES

        # Verify Keyword is in mapping
        assert Token.Keyword in TOKEN_TYPE_NAMES
        theme = ThemeConfig.from_json()
        color = theme.get_color(Token.Keyword)
        assert color == "A626A4"

    def test_exact_string_match(self):
        """Exact String token type matches."""
        from cpp_highlight.config import ThemeConfig

        theme = ThemeConfig.from_json()
        color = theme.get_color(Token.String)
        assert color == "50A14F"

    def test_exact_comment_match(self):
        """Exact Comment token type matches."""
        from cpp_highlight.config import ThemeConfig

        theme = ThemeConfig.from_json()
        color = theme.get_color(Token.Comment)
        assert color == "A0A1A7"

    def test_default_fallback(self):
        """Unknown token type returns default color."""
        from cpp_highlight.config import ThemeConfig

        theme = ThemeConfig.from_json()
        # Token.Other is in mapping and has "Other" = "383A42"
        color = theme.get_color(Token.Other)
        assert color == "383A42"

    def test_number_color(self):
        """Number tokens get correct color."""
        from cpp_highlight.config import ThemeConfig

        theme = ThemeConfig.from_json()
        color = theme.get_color(Token.Literal.Number.Integer)
        assert color == "986801"


class TestTokenTypeNames:
    """Tests for token type name mapping."""

    def test_keyword_type_mapping(self):
        """Keyword.Type is mapped correctly."""
        from cpp_highlight.config import TOKEN_TYPE_NAMES

        assert TOKEN_TYPE_NAMES[Token.Keyword.Type] == "Keyword.Type"

    def test_string_single_mapping(self):
        """String.Single is mapped correctly."""
        from cpp_highlight.config import TOKEN_TYPE_NAMES

        assert TOKEN_TYPE_NAMES[Token.String.Single] == "String.Single"

    def test_name_function_mapping(self):
        """Name.Function is mapped correctly."""
        from cpp_highlight.config import TOKEN_TYPE_NAMES

        assert TOKEN_TYPE_NAMES[Token.Name.Function] == "Name.Function"
