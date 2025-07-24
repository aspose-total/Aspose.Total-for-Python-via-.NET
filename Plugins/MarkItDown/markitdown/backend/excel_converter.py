import os
import tempfile
from aspose.cells import Workbook, MarkdownSaveOptions, JsonSaveOptions, HtmlSaveOptions
from .interfaces import IDocumentConverter

class ExcelConverter(IDocumentConverter):
    def _convert(self, file_path: str, options, suffix: str) -> str:
        workbook = Workbook(file_path)
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            temp_path = tmp.name
        workbook.save(temp_path, options)
        try:
            with open(temp_path, "r", encoding="utf-8") as f:
                return f.read()
        finally:
            os.remove(temp_path)

    def convert_to_md(self, file_path: str) -> str:
        return self._convert(file_path, MarkdownSaveOptions(), ".md")

    def convert_to_json(self, file_path: str) -> str:
        return self._convert(file_path, JsonSaveOptions(), ".json")

    def convert_to_html(self, file_path: str) -> str:
        return self._convert(file_path, HtmlSaveOptions(), ".html")
