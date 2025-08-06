import os
import tempfile
import logging

try:
    import aspose.words as aw
except ImportError:
    aw = None
try:
    import aspose.cells as ac
except ImportError:
    ac = None

def convert_word_to_excel(input_path: str, output_path: str) -> None:
    """
    Convert a Word (.doc/.docx) document to Excel (.xlsx) using Aspose.Words and Aspose.Cells.
    Args:
        input_path (str): Path to the input Word document.
        output_path (str): Path to save the output Excel file.
    Raises:
        FileNotFoundError: If the input file does not exist.
        ImportError: If required Aspose packages are missing.
        Exception: For other conversion errors.
    """
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    if not aw or not ac:
        logging.error("Required Aspose packages are not installed.")
        raise ImportError("aspose-words and aspose-cells must be installed.")
    if not os.path.isfile(input_path):
        logging.error(f"Input file not found: {input_path}")
        raise FileNotFoundError(f"Input file not found: {input_path}")
    try:
        logging.info(f"Loading Word document: {input_path}")
        doc = aw.Document(input_path)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.html') as tmp_html:
            html_path = tmp_html.name
        logging.info(f"Saving as temporary HTML: {html_path}")
        doc.save(html_path, aw.SaveFormat.HTML)
        logging.info(f"Loading HTML into Aspose.Cells")
        workbook = ac.Workbook(html_path)
        logging.info(f"Saving as Excel file: {output_path}")
        workbook.save(output_path)
        logging.info(f"Conversion successful: {output_path}")
    except Exception as e:
        logging.error(f"Conversion failed: {e}")
        raise
    finally:
        if 'html_path' in locals() and os.path.exists(html_path):
            os.remove(html_path)
            logging.info(f"Temporary HTML file deleted: {html_path}")
