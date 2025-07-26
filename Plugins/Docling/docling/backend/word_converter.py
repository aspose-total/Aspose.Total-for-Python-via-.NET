import os
import logging
from pathlib import Path
import aspose.words as aw
from .interfaces import IDocumentConverter

logger = logging.getLogger(__name__)

class WordConverter(IDocumentConverter):
    def _ensure_output_path(self, output_dir: str, output_filename: str, default_ext: str) -> str:
        try:
            os.makedirs(output_dir, exist_ok=True)
            if output_filename is None:
                output_filename = f"output{default_ext}"
            return os.path.join(output_dir, output_filename)
        except Exception as e:
            logger.exception("Failed to build output path.")
            raise

    def _remove_shapes(self, doc: aw.Document):
        builder = aw.DocumentBuilder(doc)
        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        for i, shape in enumerate(list(shapes), start=1):
            builder.move_to(shape)
            builder.write(f"<!-- Shape {i} Removed -->")
            shape.remove()

    def _convert(self, file_path: str, save_format, output_dir: str, output_filename: str, ext: str) -> str:
        try:
            logger.info(f"Converting {file_path} to format {save_format}...")
            doc = aw.Document(file_path)
            self._remove_shapes(doc)
            output_path = self._ensure_output_path(output_dir, output_filename, ext)
            doc.save(output_path, save_format)
            logger.info(f"Saved to {output_path}")
            return output_path
        except Exception as e:
            logger.exception("Failed during conversion.")
            raise

    def _convert_to_markdown_with_options(self, file_path: str, output_dir: str, output_filename: str, export_base64: bool) -> str:
        try:
            logger.info(f"Converting {file_path} to Markdown with base64={export_base64}...")
            doc = aw.Document(file_path)
            os.makedirs(output_dir, exist_ok=True)

            if output_filename is None:
                output_filename = f"{Path(file_path).stem}.md"

            md_path = os.path.join(output_dir, output_filename)

            options = aw.saving.MarkdownSaveOptions()
            options.export_images_as_base64 = export_base64

            if not export_base64:
                image_folder_name = f"{Path(output_filename).stem}_images"
                image_folder_path = os.path.join(output_dir, image_folder_name)
                options.images_folder = image_folder_path
                options.images_folder_alias = image_folder_name

            doc.save(md_path, options)
            logger.info(f"Markdown saved to {md_path}")
            return md_path
        except Exception as e:
            logger.exception("Failed during Markdown conversion.")
            raise

    def _convert_to_html_with_options(self, file_path: str, output_dir: str, output_filename: str, export_base64: bool) -> str:
        try:
            logger.info(f"Converting {file_path} to HTML with base64={export_base64}...")
            doc = aw.Document(file_path)
            os.makedirs(output_dir, exist_ok=True)

            if output_filename is None:
                output_filename = f"{Path(file_path).stem}.html"

            html_path = os.path.join(output_dir, output_filename)

            options = aw.saving.HtmlSaveOptions()
            options.export_images_as_base64 = export_base64
            options.export_original_url_for_linked_images = False

            if not export_base64:
                image_folder_name = f"{Path(output_filename).stem}_images"
                image_folder_path = os.path.join(output_dir, image_folder_name)
                options.images_folder = image_folder_path
                options.images_folder_alias = image_folder_name

            doc.save(html_path, options)
            logger.info(f"HTML saved to {html_path}")
            return html_path
        except Exception as e:
            logger.exception("Failed during HTML conversion.")
            raise
    
    # === Interface Methods ===

    # === Markdown Methods ===
    def convert_to_md(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        return self._convert(file_path, aw.SaveFormat.MARKDOWN, output_dir, output_filename, ".md")

    def convert_to_md_with_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        return self._convert_to_markdown_with_options(file_path, output_dir, output_filename, export_base64=False)

    def convert_to_md_with_base64_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        return self._convert_to_markdown_with_options(file_path, output_dir, output_filename, export_base64=True)

    # === HTML Methods ===
    def convert_to_html(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        return self._convert(file_path, aw.SaveFormat.HTML, output_dir, output_filename, ".html")

    def convert_to_html_with_images(self, file_path: str , output_dir: str, output_filename: str = None) -> str:
        return self._convert_to_html_with_options(file_path, output_dir, output_filename, export_base64=False)

    def convert_to_html_with_base64_images(self, file_path: str, output_filename: str, output_dir: str = None) -> str:
        return self._convert_to_html_with_options(file_path, output_filename, output_dir, export_base64=True)
    
    # === JSON Methods ===
    def convert_to_json(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        logger.warning("convert_to_json is not implemented.")
        raise NotImplementedError("Word to JSON conversion is not supported yet. Coming soon!")

    def convert_to_json_with_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        logger.warning("convert_to_json_with_images is not implemented.")
        raise NotImplementedError("Word to JSON conversion is not supported yet. Coming soon!")

    def convert_to_json_with_base64_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        logger.warning("convert_to_json_with_base64_images is not implemented.")
        raise NotImplementedError("Word to JSON conversion is not supported yet. Coming soon!")
