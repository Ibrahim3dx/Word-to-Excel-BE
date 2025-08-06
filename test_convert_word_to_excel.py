import sys
from convert_word_to_excel import convert_word_to_excel

if __name__ == "__main__":
    # Example usage: python test_convert_word_to_excel.py input.docx output.xlsx
    if len(sys.argv) < 3:
        print("Usage: python test_convert_word_to_excel.py <input.docx> <output.xlsx>")
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    try:
        convert_word_to_excel(input_path, output_path)
        print(f"Conversion completed: {output_path}")
    except Exception as e:
        print(f"Error: {e}")
