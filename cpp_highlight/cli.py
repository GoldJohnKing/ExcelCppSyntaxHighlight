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
  cpp_highlight.exe input.xlsx                    # Output: input_output.xlsx
  cpp_highlight.exe input.xlsx -o custom.xlsx     # Output: custom.xlsx
  cpp_highlight.exe code.xlsx -v                  # Verbose mode

Drag & Drop:
  Simply drag an Excel file onto cpp_highlight.exe to process it.
  Output will be saved as <filename>_output.xlsx in the same directory.
        """,
    )

    parser.add_argument("input", help="Input Excel file path (or drag & drop)")
    parser.add_argument(
        "-o", "--output", help="Output Excel file path (default: <input>_output.xlsx)"
    )
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

    # Generate default output path if not specified
    if args.output:
        output_path = args.output
    else:
        output_path = str(
            input_path.parent / f"{input_path.stem}_output{input_path.suffix}"
        )

    count = process_excel(str(input_path), output_path, args.verbose)

    print(f"Processed {count} cells with C++ code")
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_path}")


if __name__ == "__main__":
    main()
