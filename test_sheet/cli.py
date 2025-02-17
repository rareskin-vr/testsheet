# testsheet/cli.py
import sys
import os
from test_sheet import test_sheet

def main():
    if len(sys.argv) != 2:
        print("Usage: testsheet <input_test_script_file_or_directory>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_file = f"{os.path.splitext(os.path.basename(input_path))[0]}_test_documentation.xlsx"

    extractor = test_sheet.TestSheet(input_path, output_file)
    extractor.run()

if __name__ == "__main__":
    main()