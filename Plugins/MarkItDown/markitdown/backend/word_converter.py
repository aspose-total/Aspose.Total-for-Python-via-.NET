import os
import tempfile
import aspose.words as aw
from .interfaces import IDocumentConverter

class WordConverter(IDocumentConverter):
    def _convert(self, file_path: str, save_format, suffix: str) -> str:
        doc = aw.Document(file_path)
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            temp_path = tmp.name
        doc.save(temp_path, save_format)
        try:
            with open(temp_path, "r", encoding="utf-8") as f:
                return f.read()
        finally:
            os.remove(temp_path)

    def convert_to_md(self, file_path: str) -> str:
        return self._convert(file_path, aw.SaveFormat.MARKDOWN, ".md")

    def convert_to_html(self, file_path: str) -> str:
        return self._convert(file_path, aw.SaveFormat.HTML, ".html")

    def convert_to_json(self, file_path: str) -> str:
        raise NotImplementedError("Word to JSON conversion is not supported yet. Coming soon!")
    