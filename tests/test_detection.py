"""Tests for C++ code detection logic."""

import pytest


class TestCppDetectionHighConfidence:
    """Tests for high-confidence C++ detection patterns."""

    def test_include_directive(self):
        """#include is high-confidence C++ indicator."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code("#include <vector>") is True
        assert is_cpp_code('#include "myfile.h"') is True

    def test_std_namespace(self):
        """std:: is high-confidence C++ indicator."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code("std::cout << hello;") is True
        assert is_cpp_code("auto x = std::make_unique<int>();") is True

    def test_int_main(self):
        """int main() is high-confidence C++ indicator."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code("int main() { return 0; }") is True
        assert is_cpp_code("int main(int argc, char* argv[])") is True

    def test_template_declaration(self):
        """template< is high-confidence C++ indicator."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code("template<typename T>") is True
        assert is_cpp_code("template<class T, class U>") is True

    def test_using_namespace(self):
        """using namespace is high-confidence C++ indicator."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code("using namespace std;") is True


class TestCppDetectionMediumConfidence:
    """Tests for medium-confidence C++ detection patterns."""

    def test_single_medium_pattern_insufficient(self):
        """Single medium pattern match is insufficient."""
        from cpp_highlight.core.detection import is_cpp_code

        # Only 1 match for 'int'
        assert is_cpp_code("int x") is False

    def test_two_medium_patterns_short_text_insufficient(self):
        """Two medium patterns in short text are insufficient."""
        from cpp_highlight.core.detection import is_cpp_code

        # 2 matches but short text
        assert is_cpp_code("int x; int y") is False

    def test_two_medium_patterns_long_text_sufficient(self):
        """Two medium patterns in long text (>100 chars) are sufficient."""
        from cpp_highlight.core.detection import is_cpp_code

        # 2 matches + long text
        long_code = "int x; int y; " + "x" * 100
        assert is_cpp_code(long_code) is True

    def test_three_medium_patterns_sufficient(self):
        """Three medium pattern matches are sufficient."""
        from cpp_highlight.core.detection import is_cpp_code

        # 3 occurrences of 'int'
        assert is_cpp_code("int x; int y; int z") is True

    def test_class_declaration(self):
        """class declaration counts as medium pattern."""
        from cpp_highlight.core.detection import is_cpp_code

        # 3 patterns: class, int, int
        code = "class MyClass { int x; int y; };"
        assert is_cpp_code(code) is True

    def test_for_loop(self):
        """for loop counts as medium pattern."""
        from cpp_highlight.core.detection import is_cpp_code

        # 3 patterns: for, int, <<
        code = "for (int i = 0; i < 10; ++i) { cout << i; }"
        assert is_cpp_code(code) is True


class TestCppDetectionEdgeCases:
    """Tests for edge cases in C++ detection."""

    def test_empty_string(self):
        """Empty string is not C++ code."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code("") is False

    def test_none_input(self):
        """None is not C++ code."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code(None) is False

    def test_non_string_input(self):
        """Non-string input is not C++ code."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code(123) is False
        assert is_cpp_code(["code"]) is False

    def test_plain_text(self):
        """Plain English text is not detected as C++."""
        from cpp_highlight.core.detection import is_cpp_code

        assert is_cpp_code("Hello, this is just plain text.") is False
        assert is_cpp_code("The quick brown fox jumps.") is False

    def test_code_with_comments(self):
        """C++ with comments is detected."""
        from cpp_highlight.core.detection import is_cpp_code

        # High confidence from #include
        code = "#include <vector>\n// This is a comment"
        assert is_cpp_code(code) is True

    def test_multiline_code(self):
        """Multiline C++ code is detected."""
        from cpp_highlight.core.detection import is_cpp_code

        code = """#include <iostream>
int main() {
    std::cout << "Hello";
    return 0;
}"""
        assert is_cpp_code(code) is True
