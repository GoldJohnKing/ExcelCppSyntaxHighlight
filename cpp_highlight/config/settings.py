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
