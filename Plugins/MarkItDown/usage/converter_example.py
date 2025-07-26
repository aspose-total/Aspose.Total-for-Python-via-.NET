import os
import sys
from pathlib import Path

import logging

logging.basicConfig(level=logging.INFO)

# Add project root to sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, project_root)

from markitdown.backend.word_converter import WordConverter
from markitdown.backend.excel_converter import ExcelConverter
from markitdown.backend.powerpoint_convertor import PptConverter

def convert_document(input_file: str, output_dir: str = "output", output_filename: str = None):
    if not os.path.exists(input_file):
        print(f"Input file not found: {input_file}")
        return

    ext = Path(input_file).suffix.lower()
    converter = None

    if ext in [".docx", ".doc",".pdf"]:
        converter = WordConverter()
    elif ext in [".xlsx", ".xls"]:
        converter = ExcelConverter()
    elif ext in [".pptx", ".ppt"]:
        converter = PptConverter()    
    else:
        print(f"Unsupported file extension: {ext}")
        return

    # Default output filename if not given
    #if not output_filename:
     #   output_filename = f"{Path(input_file).stem}.md"

    if not output_filename:
        output_filename = f"{Path(input_file).stem}.md"

    #if not output_filename:
     #   output_filename = f"{Path(input_file).stem}.html"

    markdown_path = converter.convert_to_md(input_file, output_dir, output_filename)
    #markdown_path = converter.convert_to_md_with_images(input_file, output_dir, output_filename)
    print(f"Markdown written to: {markdown_path}")
    #json_path = converter.convert_to_json(input_file, output_dir, output_filename)
    #print(f"JSON written to: {json_path}")
    #html_path = converter.convert_to_html(input_file, output_dir, output_filename)
    #print(f"HTML written to: {html_path}")


def main():
    # Choose the file you want to test
    word_file = os.path.join(os.path.dirname(__file__), "templates/word/WordTablesImages.docx")
    pdf_file = os.path.join(os.path.dirname(__file__), "templates/pdf/WordTablesImages.pdf")
    excel_file = os.path.join(os.path.dirname(__file__), "templates/excel/test-01.xlsx")
    powerpoint_file = os.path.join(os.path.dirname(__file__), "templates/powerpoint/powerpoint_with_image.pptx")
    #Uncomment one of the following lines to test:
    convert_document(word_file, output_dir="output", output_filename="from_word.md")
    #convert_document(pdf_file, output_dir="output", output_filename="from_word_for_pdf.md")
    #convert_document(excel_file, output_dir="output", output_filename="from_excel.md")
    #convert_document(excel_file, output_dir="output", output_filename="from_excel.html")
    #convert_document(excel_file, output_dir="output", output_filename="from_excel.json")
    #convert_document(powerpoint_file, output_dir="output", output_filename="from_powerpoint.html")

if __name__ == "__main__":
    main()
