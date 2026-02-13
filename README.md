# C++ Excel Highlighter

A command-line tool that automatically detects C++ code in Excel cells and applies syntax highlighting using the Atom One Light color theme.

## Features

- **Automatic Detection**: Intelligently identifies cells containing C++ code using pattern matching
- **Syntax Highlighting**: Applies Atom One Light theme colors to C++ code
- **Comment Support**: Correctly detects and highlights C++ comments (`//` and `/* */`)
- **Format Preservation**: Maintains existing cell formatting (alignment, borders, etc.)
- **Error Tolerance**: Works even with syntax errors in the code
- **Pure Python**: No external dependencies other than openpyxl and Pygments

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python cpp_highlight.py input.xlsx -o output.xlsx
```

### Verbose Mode

```bash
python cpp_highlight.py input.xlsx -o output.xlsx --verbose
```

### Examples

```bash
# Process a single file
python cpp_highlight.py code_examples.xlsx -o highlighted.xlsx

# With verbose output to see which cells were processed
python cpp_highlight.py input.xlsx -o output.xlsx -v

# Process and verify
python cpp_highlight.py my_code.xlsx -o my_code_highlighted.xlsx --verbose
```

## How It Works

1. **Detection**: The tool scans all cells and uses pattern matching to identify C++ code:
   - High-confidence patterns: `#include`, `int main`, `std::`, `template<`, etc.
   - Medium-confidence patterns: keywords, types, control flow statements, **comments**
   - A cell is considered C++ code if it has 1+ high-confidence patterns OR 3+ medium-confidence matches (counting all occurrences)

2. **Tokenization**: Detected C++ code is tokenized using Pygments' C++ lexer

3. **Highlighting**: Each token is assigned a color based on the Atom One Light theme

4. **Output**: The highlighted text is saved as Rich Text in Excel, preserving all original formatting

## Color Theme

The tool uses the Atom One Light color scheme by default. Colors are defined in a JSON configuration file (`theme.json`) located in the same directory as the script or executable.

### Default Colors

| Token Type | Color | Example |
|------------|-------|---------|
| Keywords | Purple (#A626A4) | `int`, `class`, `return` |
| Strings | Green (#50A14F) | `"hello"` |
| Comments | Gray (#A0A1A7) | `// comment`, `/* block */` |
| Numbers | Gold (#986801) | `42`, `3.14` |
| Functions | Blue (#4078F2) | `main()`, `printf()` |
| Classes/Types | Gold (#C18401) | `std::string`, `vector` |
| Variables | Red (#E45649) | `x`, `count` |
| Operators | Dark Gray (#383A42) | `+`, `-`, `=` |

### Customizing Colors

You can customize the color theme by editing the `theme.json` file. The file will be automatically created on first run with the default Atom One Light theme.

**Example theme.json:**
```json
{
  "Keyword": "A626A4",
  "String": "50A14F",
  "Comment": "A0A1A7",
  "Number": "986801",
  "Name.Function": "4078F2",
  "Name.Class": "C18401",
  "Name": "E45649",
  "Operator": "383A42",
  "Punctuation": "383A42",
  "Text": "383A42"
}
```

Colors should be specified as 6-digit hex codes (without the `#` prefix).

## Detection Algorithm

The detection algorithm uses a confidence-based approach with occurrence counting:

### High Confidence Patterns (1+ match = detected)
- `#include <...>` or `#include "..."`
- `using namespace ...`
- `int main(`
- `std::`
- `template<`
- Function calls like `Class::method()`

### Medium Confidence Patterns (3+ total matches = detected)
- **Comments**: `//` single-line, `/* */` multi-line
- Type keywords: `int`, `char`, `float`, `double`, `void`, `bool`, `auto`, `const`
- Control flow: `for`, `while`, `if`, `else`, `switch`, `return`
- Class/struct definitions: `class`, `struct`, `enum`
- Access specifiers: `public:`, `private:`, `protected:`
- Preprocessor: `#define`, `#ifdef`, `#endif`
- Common I/O: `cout`, `cin`, `endl`, `printf`
- Stream operators: `<<`, `>>`
- STL containers: `string`, `vector`, `map`, `set`, `array`

**Note**: The algorithm counts total occurrences, not unique pattern types. For example, code with `int x = 10; int y = 20;` counts as 2 type keyword matches.

## Limitations

- **Mixed Content**: Cells containing both code and regular text are treated as a whole. If the cell is detected as code, the entire content will be highlighted.
- **Syntax Errors**: While the tool handles most syntax errors gracefully, unclosed strings or comments may cause incorrect highlighting of subsequent content.
- **Single Language**: Only C++ code is supported. Other languages will not be highlighted.

## Requirements

- Python 3.6+
- openpyxl >= 3.1.0
- Pygments >= 2.16.0

## License

MIT License
