import os
import sys
import argparse

# Add project root to sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, project_root)

from markitdown.backend.word_converter import WordConverter

def main():
    parser = argparse.ArgumentParser(description="Convert Word files to Markdown or HTML.")
    parser.add_argument("--input", required=True, help="Path to the input Word (.docx) file")
    parser.add_argument("--output-dir", required=True, help="Directory to save output files")
    parser.add_argument("--format", required=True, choices=[
        "md", "md_with_images", "md_with_base64",
        "html", "html_with_images", "html_with_base64"
    ], help="Output format")

    args = parser.parse_args()

    input_file = args.input
    output_dir = args.output_dir
    fmt = args.format

    if not os.path.exists(input_file):
        print(f"Input file not found: {input_file}")
        return

    converter = WordConverter()

    try:
        if fmt == "md":
            output_path = converter.convert_to_md(input_file, output_dir)
        elif fmt == "md_with_images":
            output_path = converter.convert_to_md_with_images(input_file, output_dir)
        elif fmt == "md_with_base64":
            output_path = converter.convert_to_md_with_base64_images(input_file, output_dir)
        elif fmt == "html":
            output_path = converter.convert_to_html(input_file, output_dir)
        elif fmt == "html_with_images":
            output_path = converter.convert_to_html_with_images(input_file, output_dir)
        elif fmt == "html_with_base64":
            output_path = converter.convert_to_html_with_base64_images(input_file, output_dir)
        else:
            print("Unsupported format.")
            return

        print(f"Conversion successful: {output_path}")

    except Exception as e:
        print(f"Conversion failed: {e}")

if __name__ == "__main__":
    main()
