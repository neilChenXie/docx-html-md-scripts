#!/usr/bin/env python3
#
# Convert Word document (.docx) to HTML format
#
# Uses pandoc Python package (pypandoc) instead of command-line tool.
#
# Usage:
#   python docx_to_html_py.py <input.docx> [output.html]

import os
import sys

import pypandoc


def main():
    if len(sys.argv) < 2:
        print("Usage: python docx_to_html_py.py <input.docx> [output.html]")
        print("Example: python docx_to_html_py.py document.docx")
        print("         python docx_to_html_py.py document.docx output.html")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) >= 3 else ""

    if not os.path.isfile(input_file):
        print(f"Error: File not found: {input_file}")
        sys.exit(1)

    input_file = os.path.realpath(input_file)

    if not input_file.lower().endswith(".docx"):
        print(f"Error: Input file must have .docx extension: {input_file}")
        sys.exit(1)

    input_dir = os.path.dirname(input_file)
    input_basename = os.path.splitext(os.path.basename(input_file))[0]

    if not output_file:
        output_file = os.path.join(input_dir, f"{input_basename}.html")
    else:
        output_file = os.path.expanduser(output_file)
        if not os.path.isabs(output_file):
            output_file = os.path.join(os.getcwd(), output_file)

    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)

    print(f"Converting: {input_file}")
    print(f"Output:     {output_file}")

    pandoc_version = pypandoc.get_pandoc_version()
    print(f"Using Pandoc ({pandoc_version}) via pypandoc...")

    media_dir = os.path.splitext(output_file)[0] + ".files"

    output = pypandoc.convert_file(
        input_file,
        "html",
        format="docx",
        extra_args=[
            "--standalone",
            "--extract-media=" + media_dir,
        ],
        outputfile=output_file,
    )

    print("[OK] Pandoc conversion successful")

    if os.path.isfile(output_file):
        file_size = os.path.getsize(output_file)
        print()
        print("=" * 40)
        print("[OK] Conversion successful!")
        print(f"[OK] Output: {output_file}")
        print(f"[OK] Size: {file_size} bytes")
        print("=" * 40)


if __name__ == "__main__":
    main()