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
