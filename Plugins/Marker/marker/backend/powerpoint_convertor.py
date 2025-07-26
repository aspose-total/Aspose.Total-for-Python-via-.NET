# powerpoint_converter.py

import os
import logging
import aspose.slides as slides
from .interfaces import IDocumentConverter

logger = logging.getLogger(__name__)

class PptConverter(IDocumentConverter):
    def convert_to_md(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        os.makedirs(output_dir, exist_ok=True)
        markdown_file_path = os.path.join(output_dir, output_file)
        try:
            with slides.Presentation(file_path) as pres:
                save_options = slides.export.MarkdownSaveOptions()
                save_options.export_type = slides.export.MarkdownExportType.VISUAL

                # ✅ This ensures images are saved alongside .md file
                save_options.images_save_folder_name = ""
                save_options.base_path = output_dir
                pres.save(markdown_file_path, slides.export.SaveFormat.MD, save_options)
            logger.info(f"PowerPoint converted to Markdown: {markdown_file_path}")
            return markdown_file_path
        except Exception as e:
            logger.error(f"Failed to convert PowerPoint to Markdown: {e}")
            raise

    def convert_to_md_with_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        logger.warning("convert_to_md_with_images is not implemented for PptConverter")
        raise NotImplementedError("convert_to_md_with_images is not implemented for PptConverter")

    def convert_to_md_with_base64_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        logger.warning("convert_to_md_with_base64_images is not implemented for PptConverter")
        raise NotImplementedError("convert_to_md_with_base64_images is not implemented for PptConverter")

    def convert_to_html(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        os.makedirs(output_dir, exist_ok=True)
        output_file = output_file or "converted_ppt.html"
        html_file_path = os.path.join(output_dir, output_file)

        try:
            with slides.Presentation(file_path) as pres:
                save_options = slides.export.MarkdownSaveOptions()
                save_options.export_type = slides.export.HTMLSaveOptions
                pres.save(html_file_path,save_options)
                #pres.save(html_file_path, slides.export.SaveFormat.HTML)

            logger.info(f"PowerPoint converted to HTML: {html_file_path}")
            return html_file_path
        except Exception as e:
            logger.error(f"Failed to convert PowerPoint to HTML: {e}")
            raise

    def convert_to_html_with_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        logger.warning("convert_to_html_with_images is not implemented for PptConverter")
        raise NotImplementedError("convert_to_html_with_images is not implemented for PptConverter")

    def convert_to_html_with_base64_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        logger.warning("convert_to_html_with_base64_images is not implemented for PptConverter")
        raise NotImplementedError("convert_to_html_with_base64_images is not implemented for PptConverter")

    def convert_to_json(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        logger.warning("convert_to_json is not implemented for PptConverter")
        raise NotImplementedError("convert_to_json is not implemented for PptConverter")

    def convert_to_json_with_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        logger.warning("convert_to_json_with_images is not implemented for PptConverter")
        raise NotImplementedError("convert_to_json_with_images is not implemented for PptConverter")

    def convert_to_json_with_base64_images(self, file_path: str, output_dir: str, output_file: str = None) -> str:
        logger.warning("convert_to_json_with_base64_images is not implemented for PptConverter")
        raise NotImplementedError("convert_to_json_with_base64_images is not implemented for PptConverter")
