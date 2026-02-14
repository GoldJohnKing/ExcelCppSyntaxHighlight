# cpp_highlight/cli.py
"""Command-line interface for ExcelCppSyntaxHighlight."""

import argparse
import sys
from pathlib import Path

from cpp_highlight.processor import process_excel


def main():
    """Main entry point for CLI."""
    parser = argparse.ArgumentParser(
        description="Apply C++ syntax highlighting to Excel cells",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m cpp_highlight input.xlsx -o output.xlsx
  python -m cpp_highlight code.xlsx --output highlighted.xlsx --verbose
        """,
    )

    parser.add_argument("input", help="Input Excel file path")
    parser.add_argument("-o", "--output", required=True, help="Output Excel file path")
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable verbose output"
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if not input_path.suffix.lower() in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
        print(
            f"Warning: Input file may not be a valid Excel file: {args.input}",
            file=sys.stderr,
        )

    count = process_excel(args.input, args.output, args.verbose)

    print(f"Processed {count} cells with C++ code")
    print(f"  Input:  {args.input}")
    print(f"  Output: {args.output}")


if __name__ == "__main__":
    main()
