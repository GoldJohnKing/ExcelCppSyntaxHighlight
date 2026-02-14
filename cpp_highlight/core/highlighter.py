"""Cell highlighting logic."""

import sys
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
    """Calculate the required row height for given tokens."""
    if font_settings is None:
        font_settings = FontSettings.default()

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
        """Initialize the highlighter."""
        self.theme = theme or ThemeConfig.from_json()
        self.font = font or FontSettings.default()
        self.lexer = (lexer or CppLexer)()

    def highlight(self, text: str) -> Tuple[Optional[CellRichText], Optional[float]]:
        """Apply syntax highlighting to C++ code text."""
        try:
            tokens = list(lex(text, self.lexer))

            # Remove trailing newline token
            if (
                tokens
                and tokens[-1][0] == Token.Text.Whitespace
                and tokens[-1][1] == "\n"
            ):
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
            print(f"Warning: Failed to highlight text: {e}", file=sys.stderr)
            return None, None

    def apply_to_cell(self, cell) -> bool:
        """Apply syntax highlighting to an Excel cell."""
        text = cell.value
        if not isinstance(text, str):
            return False

        rich_text, required_height = self.highlight(text)

        if rich_text is None:
            return False

        cell.value = rich_text
        cell.alignment = Alignment(wrap_text=True, vertical="top")

        return True
