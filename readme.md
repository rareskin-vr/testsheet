# TestSheet

**TestSheet** is a Python tool designed to extract test cases from test scripts and export them into an Excel file. It automatically parses test functions, extracts descriptions, preconditions, test steps, and expected outputs from comments, and organizes them into a structured Excel sheet.

## Features

- **Extract Test Cases**: Automatically extracts test cases from Python test scripts.
- **Export to Excel**: Generates an Excel file with test case details.
- **Supports Multiline Comments**: Handles both single-line and multiline comments for test case details.
- **CLI Support**: Easy-to-use command-line interface for quick integration into workflows.

## Installation
You can install `testsheet` via pip:

```bash
pip install testsheet
````
# Usage
Command-Line Interface (CLI)<br>
Run testsheet from the command line to process a single file or a directory of test scripts:
```bash
testsheet <input_test_script_file_or_directory>
```
This will generate an Excel file named <input_file_or_directory>_test_documentation.xlsx in the current working directory.

# Example
Process a single test script file:

```bash
testsheet test_example.py
```
This will generate test_example_test_documentation.xlsx

Process a directory of test scripts:
```bash
testsheet tests_dir/
```
This will generate tests_test_documentation.xlsx, containing all test cases from files in the tests_dir/ directory.

# Supported Comment Formats
TestSheet recognizes specific tags in comments to extract test case details. Use the following tags in your test script comments:
- **Description**: # Description: <description_text>
- **Precondition**: # Precondition: <precondition_text>
- **Test Step**: # Step: <step_text>
- **Expected Output**: # Expected Output: <expected_output_text>

Example Test Script
``` python
# Description: This is a sample test case for login functionality.
def test_login():
    # Precondition: User must be on the login page.
    ''' Step: Enter valid username and password.\n
    With pass length min 8 character
    '''
    # Expected Output: User should be logged in successfully.
    
    # Step: Click the login button.
    # Expected Output: Login button should be enabled.
    pass
```
# Contributing
Contributions are welcome! If you'd like to contribute, please follow these steps:
- Fork the repository.
- Create a new branch for your feature or bugfix.
- Submit a pull request.

# License
This project is licensed under the MIT License. See the LICENSE file for details.

# Support
If you encounter any issues or have questions, please open an issue on GitHub.
