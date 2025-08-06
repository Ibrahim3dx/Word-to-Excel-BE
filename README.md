# Word to Excel Converter (Python Backend)

This project provides a Python function to convert Microsoft Word documents (`.doc`/`.docx`) to Excel spreadsheets (`.xlsx`) using Aspose.Words and Aspose.Cells for Python via .NET.

## Features
- Converts Word documents to Excel format with high fidelity
- Uses Aspose.Words for Word-to-HTML conversion
- Uses Aspose.Cells for HTML-to-Excel conversion
- Handles temporary files safely
- Includes robust error handling and logging

## Requirements
- Python 3.10 (recommended, due to Aspose.Cells compatibility)
- aspose-words
- aspose-cells

> **Note:** Aspose.Cells for Python via .NET does **not** support Python 3.11+ (including 3.13). Use Python 3.10 for full compatibility.

## Installation
1. Create and activate a Python 3.10 virtual environment:
   ```powershell
   # Windows (PowerShell)
   py -3.10 -m venv .venv
   .venv\Scripts\Activate.ps1
   ```
2. Install dependencies:
   ```sh
   pip install aspose-words aspose-cells
   ```

## Usage

### Function
Use the provided function in `convert_word_to_excel.py`:

```python
from convert_word_to_excel import convert_word_to_excel
convert_word_to_excel('input.docx', 'output.xlsx')
```

### CLI Test Script
A test script is included:
```sh
python test_convert_word_to_excel.py test.docx output.xlsx
```

## Error Handling
- Raises `FileNotFoundError` if the input file does not exist
- Raises `ImportError` if required Aspose packages are missing
- Logs and raises exceptions for conversion errors

## Troubleshooting
- If you see `ModuleNotFoundError: No module named 'aspose.cells'`, ensure you are using Python 3.10 and have installed the package in the correct environment.
- Check the logs for detailed error messages.

## License
This project is free to use. See Aspose license terms for library usage.
