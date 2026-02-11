## ADDED Requirements

### Requirement: Automatic C++ code detection
The system SHALL automatically detect cells containing C++ code based on predefined patterns using occurrence counting.

#### Scenario: High confidence detection
- **WHEN** a cell contains `#include <header>`
- **THEN** the system SHALL identify it as C++ code

#### Scenario: Multiple medium confidence patterns
- **WHEN** a cell contains at least 3 total matches from medium-confidence patterns
- **THEN** the system SHALL identify it as C++ code

#### Scenario: Comment-based detection
- **WHEN** a cell contains `// single-line comment`
- **THEN** each `//` occurrence SHALL count as 1 medium-confidence match

- **WHEN** a cell contains `/* multi-line comment */`
- **THEN** each comment block SHALL count as 1 medium-confidence match

#### Scenario: Occurrence counting
- **GIVEN** code contains `int x = 10; int y = 20;`
- **WHEN** counting medium-confidence matches
- **THEN** the system SHALL count 2 type keyword matches (both `int` occurrences)

### Requirement: C++ syntax highlighting
The system SHALL apply syntax highlighting to detected C++ code using the Atom One Light color theme.

#### Scenario: Keyword highlighting
- **WHEN** a cell contains the keyword `int`
- **THEN** the keyword SHALL be displayed in purple

#### Scenario: String highlighting
- **WHEN** a cell contains a string literal `"hello"`
- **THEN** the string SHALL be displayed in green

#### Scenario: Comment highlighting
- **WHEN** a cell contains a single-line comment `// comment`
- **THEN** the comment SHALL be displayed in gray

- **WHEN** a cell contains a multi-line comment `/* block */`
- **THEN** the comment SHALL be displayed in gray

### Requirement: Preserve existing formatting
The system SHALL preserve existing cell formatting (alignment, borders, fill) while only modifying font colors.

#### Scenario: Preserving cell alignment
- **WHEN** a cell has center alignment before highlighting
- **THEN** the cell SHALL maintain center alignment after highlighting

### Requirement: Handle syntax errors gracefully
The system SHALL apply highlighting even to code with syntax errors.

#### Scenario: Unclosed string
- **WHEN** a cell contains code with an unclosed string `"hello`
- **THEN** the system SHALL apply highlighting based on lexical analysis without failing

### Requirement: Command line interface
The system SHALL provide a command line interface for processing Excel files.

#### Scenario: Basic usage
- **WHEN** user runs `python cpp_highlight.py input.xlsx -o output.xlsx`
- **THEN** the system SHALL process the input file and save highlighted output

#### Scenario: Verbose mode
- **WHEN** user runs with `-v` or `--verbose` flag
- **THEN** the system SHALL display detailed processing information

### Requirement: Detection patterns
The system SHALL use the following detection patterns:

#### High Confidence Patterns (1+ match triggers detection)
- `#include <...>` or `#include "..."`
- `using namespace ...`
- `int main(`
- `std::`
- `template<`
- `Class::method()` (function calls via scope resolution)

#### Medium Confidence Patterns (3+ total matches triggers detection)
- **Comments**: `//` (single-line), `/* ... */` (multi-line)
- **Type keywords**: `int`, `char`, `float`, `double`, `void`, `bool`, `auto`, `const`, `constexpr`, `mutable`
- **Class definitions**: `class`, `struct`, `enum` followed by identifier
- **Control flow**: `for`, `while`, `if`, `else`, `switch`, `case`, `break`, `continue`, `return`
- **Access specifiers**: `public:`, `private:`, `protected:` (at start of line)
- **Preprocessor**: `#define`, `#ifdef`, `#ifndef`, `#endif`, `#pragma` (at start of line)
- **I/O operations**: `cout`, `cin`, `endl`, `printf`, `scanf`
- **Stream operators**: `<<`, `>>`
- **STL containers**: `string`, `vector`, `map`, `set`, `array`

**Note**: All matches are counted by total occurrences, not by unique pattern types.
