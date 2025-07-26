import os
import logging
import aspose.cells as ac
from .interfaces import IDocumentConverter

logger = logging.getLogger(__name__)

class ExcelConverter(IDocumentConverter):
    def _get_output_path(self, file_path: str, suffix: str, output_dir: str, output_file: str = None) -> str:
        os.makedirs(output_dir, exist_ok=True)
        if output_file:
            return os.path.join(output_dir, output_file)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        return os.path.join(output_dir, f"{base_name}{suffix}")

    def _convert(self, file_path: str, save_format, suffix: str, output_dir: str, output_file: str = None) -> str:
        try:
            workbook = ac.Workbook(file_path)
            output_path = self._get_output_path(file_path, suffix, output_dir, output_file)
            workbook.save(output_path, save_format)
            logger.info(f"Converted: {file_path} -> {output_path}")
            return output_path
        except Exception as e:
            logger.error(f"Failed to convert {file_path}: {e}")
            raise

    def _convert_with_options(self, file_path: str, output_dir: str, export_base64: bool,
                              save_format, suffix: str, output_file: str = None,
                              options_factory=None) -> str:
        try:
            workbook = ac.Workbook(file_path)
            options = options_factory() if options_factory else None
            if options:
                options.export_images_as_base64 = export_base64

            output_path = self._get_output_path(file_path, suffix, output_dir, output_file)
            os.makedirs(output_dir, exist_ok=True)

            if options and not export_base64:
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                image_folder_name = f"{base_name}_images"
                image_folder_path = os.path.join(output_dir, image_folder_name)
                #options.images_folder = image_folder_path
                #options.images_folder_alias = image_folder_name

            workbook.save(output_path, options or save_format)
            logger.info(f"Converted with options: {file_path} -> {output_path}")
            return output_path
        except Exception as e:
            logger.error(f"Conversion with options failed for {file_path}: {e}")
            raise

    # === Markdown Methods ===
    def convert_to_md(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        return self._convert(file_path, ac.SaveFormat.MARKDOWN, ".md", output_dir, output_file)

    def convert_to_md_with_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        return self._convert_with_options(
            file_path, output_dir, export_base64=False,
            save_format=ac.SaveFormat.MARKDOWN,
            suffix=".md", output_file=output_file,
            options_factory=ac.MarkdownSaveOptions
        )

    def convert_to_md_with_base64_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        return self._convert_with_options(
            file_path, output_dir, export_base64=True,
            save_format=ac.SaveFormat.MARKDOWN,
            suffix=".md", output_file=output_file,
            options_factory=ac.MarkdownSaveOptions
        )

    # === HTML Methods ===
    def convert_to_html(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        return self._convert(file_path, ac.SaveFormat.HTML, ".html", output_dir, output_file)

    def convert_to_html_with_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        return self._convert_with_options(
            file_path, output_dir, export_base64=False,
            save_format=ac.SaveFormat.HTML,
            suffix=".html", output_file=output_file,
            options_factory=ac.HtmlSaveOptions
        )

    def convert_to_html_with_base64_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        return self._convert_with_options(
            file_path, output_dir, export_base64=True,
            save_format=ac.SaveFormat.HTML,
            suffix=".html", output_file=output_file,
            options_factory=ac.HtmlSaveOptions
        )

    # === JSON Methods ===
    def convert_to_json(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        return self._convert(file_path, ac.SaveFormat.JSON, ".json", output_dir, output_file)

    def convert_to_json_with_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        """
        return self._convert_with_options(
            file_path, output_dir, export_base64=False,
            save_format=ac.SaveFormat.JSON,
            suffix=".json", output_file=output_file,
            options_factory=ac.JsonSaveOptions
        )
        """
        logger.warning("convert_to_json_with_images is not implemented.")
        raise NotImplementedError("Excel to JSON conversion with images is not supported yet. Coming soon!")

    def convert_to_json_with_base64_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        """
        return self._convert_with_options(
            file_path, output_dir, export_base64=True,
            save_format=ac.SaveFormat.JSON,
            suffix=".json", output_file=output_file,
            options_factory=ac.JsonSaveOptions
        )
        """
        logger.warning("convert_to_json_with_base64_images is not implemented.")
        raise NotImplementedError("Excel to JSON conversion with base64 images is not supported yet. Coming soon!")
        
