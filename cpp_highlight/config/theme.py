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
        base_path = Path(sys.executable).parent
    else:
        base_path = Path(__file__).parent.parent.parent
    return base_path / "theme.json"


@dataclass
class ThemeConfig:
    """Theme configuration for syntax highlighting."""

    colors: Dict[str, str] = field(default_factory=lambda: DEFAULT_THEME_COLORS.copy())
    token_names: Dict[Token, str] = field(
        default_factory=lambda: TOKEN_TYPE_NAMES.copy()
    )
    default_color: str = "383A42"

    @classmethod
    def from_json(cls, path: Optional[Path] = None) -> "ThemeConfig":
        """Load theme configuration from JSON file."""
        if path is None:
            path = get_config_path()

        colors = DEFAULT_THEME_COLORS.copy()

        if not path.exists():
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
        """Get color for a token type."""
        type_name = self.token_names.get(token_type)
        if type_name and type_name in self.colors:
            return self.colors[type_name]

        for ttype, name in self.token_names.items():
            if token_type in ttype and name in self.colors:
                return self.colors[name]

        return self.default_color
