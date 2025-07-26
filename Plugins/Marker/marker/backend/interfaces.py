# interfaces.py

from abc import ABC, abstractmethod

from abc import ABC, abstractmethod

class IDocumentConverter(ABC):
    @abstractmethod
    def convert_to_md(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the input document to Markdown format.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated Markdown file.
        """
        pass

    def convert_to_md_with_images(self, file_path: str , output_dir: str , output_filename: str = None) -> str:
        """
        Convert the input document to Markdown with image links (not base64).

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated Markdown file.
        """
        pass

    def convert_to_md_with_base64_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the input document to Markdown with base64-encoded images.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated Markdown file.
        """
        pass

    @abstractmethod
    def convert_to_json(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the input document to JSON format.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated JSON file.
        """
        pass

    @abstractmethod
    def convert_to_json_with_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the document to JSON and export any images as separate files.

        Intended for structured content pipelines where both text and media are extracted.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated JSON file.
        """
        pass

    @abstractmethod
    def convert_to_json_with_base64_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the document to JSON with images embedded as base64.

        Useful for API or data exchange scenarios where everything must be self-contained.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated JSON file.
        """
        pass


    @abstractmethod
    def convert_to_html(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the input document to HTML format.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated HTML file.
        """
        pass

    @abstractmethod
    def convert_to_html_with_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the document to HTML with images saved as separate files.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated HTML file.
        """
        pass

    @abstractmethod    
    def convert_to_html_with_base64_images(self, file_path: str, output_dir: str, output_filename: str = None) -> str:
        """
        Convert the document to HTML with images embedded as base64.

        Args:
            file_path (str): Path to the input file.
            output_dir (str): Directory where output should be saved.
            output_filename (str, optional): Optional name for the output file.

        Returns:
            str: Path to the generated HTML file.
        """
        pass

