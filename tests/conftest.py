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
