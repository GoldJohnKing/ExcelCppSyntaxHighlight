## Purpose

Define the Atom One Light color theme for C++ syntax highlighting in Excel exports.

## Requirements

### Requirement: Theme uses Atom One Light color palette
The system SHALL use the Atom One Light theme colors for all C++ syntax highlighting.

#### Scenario: Keywords are purple
- **WHEN** the highlighter processes a C++ keyword (e.g., `class`, `public`, `void`, `return`)
- **THEN** the token SHALL be rendered with color #A626A4 (purple)

#### Scenario: Operator words are purple
- **WHEN** the highlighter processes an operator word (e.g., `and`, `or`, `not`)
- **THEN** the token SHALL be rendered with color #A626A4 (purple)

#### Scenario: Strings are green
- **WHEN** the highlighter processes a string literal
- **THEN** the token SHALL be rendered with color #50A14F (green)

#### Scenario: Documentation strings are gray
- **WHEN** the highlighter processes a documentation string
- **THEN** the token SHALL be rendered with color #A0A1A7 (gray)

#### Scenario: Comments are gray
- **WHEN** the highlighter processes a single-line comment (`//`)
- **THEN** the token SHALL be rendered with color #A0A1A7 (gray)

- **WHEN** the highlighter processes a multi-line comment (`/* */`)
- **THEN** the token SHALL be rendered with color #A0A1A7 (gray)

#### Scenario: Preprocessor comments are purple
- **WHEN** the highlighter processes a preprocessor comment (e.g., `#if 0` blocks)
- **THEN** the token SHALL be rendered with color #A626A4 (purple)

#### Scenario: Numbers are gold
- **WHEN** the highlighter processes a numeric literal
- **THEN** the token SHALL be rendered with color #986801 (gold/brown)

#### Scenario: Functions are blue
- **WHEN** the highlighter processes a function name
- **THEN** the token SHALL be rendered with color #4078F2 (blue)

#### Scenario: Types are gold
- **WHEN** the highlighter processes a type name (e.g., `Rectangle`, `int`, `string`)
- **THEN** the token SHALL be rendered with color #C18401 (gold)

#### Scenario: Variables are red
- **WHEN** the highlighter processes a variable name
- **THEN** the token SHALL be rendered with color #E45649 (red)

#### Scenario: Operators are dark gray
- **WHEN** the highlighter processes an operator (e.g., `=`, `+`, `::`)
- **THEN** the token SHALL be rendered with color #383A42 (dark gray)

#### Scenario: General text is dark gray
- **WHEN** the highlighter processes general text or punctuation
- **THEN** the token SHALL be rendered with color #383A42 (dark gray)

#### Scenario: Errors are red
- **WHEN** the highlighter encounters a lexical error
- **THEN** the token SHALL be rendered with color #FF0000 (red)

### Requirement: Theme constant is properly named
The system SHALL define the theme with a descriptive constant name.

#### Scenario: Theme constant reflects actual theme
- **WHEN** examining the code
- **THEN** the theme dictionary SHALL be named `ATOM_ONE_LIGHT_THEME`

## Color Reference

| Color Name | Hex Code | Usage |
|------------|----------|-------|
| Purple | #A626A4 | Keywords, operator words, preprocessor comments |
| Green | #50A14F | String literals |
| Gray | #A0A1A7 | Comments, documentation strings |
| Gold | #986801 | Numeric literals |
| Gold/Brown | #C18401 | Type names, classes, built-ins |
| Blue | #4078F2 | Function names |
| Red | #E45649 | Variable names |
| Dark Gray | #383A42 | Operators, punctuation, general text |
| Red | #FF0000 | Lexical errors |
