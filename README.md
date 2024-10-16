# ExcelToDump

ExcelToDump is a command-line tool that reads an Excel file, extracts test information, and generates a ZIP archive containing XML files with the extracted content. The tool is useful for converting structured data from Excel into a format that can be used for testing automation purposes.

## Features

- Reads Excel files and extracts data from key columns (e.g., "Precondition", "Action", "Expected Result").
- Generates XML files representing the extracted data in a structured format.
- Packs the generated XML files into a ZIP archive, making it easy to distribute or integrate with other tools.
- Automatically categorizes extracted parameters as "Text", "Numeric", or "Comparison".
- Operates directly from the command line, providing an efficient workflow for automation.

## Installation

1. Clone the repository or download the source files:
   ```sh
   git clone https://github.com/redonhalimaj/exceltodump.git
   cd exceltodump
   ```

2. Install the package locally using `pip`:
   ```sh
   pip install -e .
   ```

## Requirements

- Python 3.6+
- pandas (for reading Excel files)
- openpyxl (for Excel file compatibility with pandas)

You can install the required dependencies by running:
```sh
pip install pandas openpyxl
```

## Usage

After installing the package, you can use the command-line tool to convert an Excel file into a ZIP file containing the generated XML:

```sh
exceltodump <path_to_excel_file>
```

For example:

```sh
exceltodump Syntax.xlsx
```

### Arguments
- `<path_to_excel_file>`: The path to the Excel file you want to convert.

### Example Output
Upon successful execution, the tool generates a ZIP file named `project-dump.zip`, which contains the following XML files:

- **output_test_elements.xml**: This XML contains the test elements extracted from the Excel file, including data types, equivalence classes, and other structured information.
- **output_testcase.xml**: This XML represents the generated test case, including interactions and parameters.
- **project-dump.xml**: A placeholder XML representing the project dump.

## Excel File Format
The tool expects the Excel file to have certain columns to extract relevant data, such as:

- **Precondition**: Describes the preconditions required for the test.
- **Action**: Specifies the steps to be taken during the test.
- **Expected Result**: Defines the expected outcome after executing the action.

The tool extracts operations, parameters, and other related data from these columns using regex patterns.

## How It Works
1. **Excel Reading**: The tool reads the provided Excel file using `pandas`.
2. **Data Extraction**: It extracts operations, parameters, and descriptions using predefined regex patterns.
3. **XML Generation**: The extracted information is used to create XML representations of the test data.
4. **ZIP Creation**: The generated XML files are packed into a ZIP archive named `project-dump.zip`.

## Error Handling
- If the specified Excel file is not found, the tool will display an error message.
- If `project-dump.xml` is missing, a placeholder will be created to ensure the ZIP file generation completes without issues.

## Development
If you want to make changes to the project, you can install it in editable mode:

```sh
pip install -e .
```

### Running Tests
To run the tests for this project, you can use the `pytest` framework (make sure you have installed the required test dependencies):

```sh
pytest
```

## Troubleshooting
- **ModuleNotFoundError**: If you see this error, ensure all Python modules (`converter.py`, `reader.py`, etc.) are in the correct directory and imported correctly.
- **File Not Found**: Ensure the Excel file is in the correct path and has the expected columns.
- **XML Errors**: If XML generation fails, verify the extracted data from the Excel file. Improper formatting can cause issues during XML generation.

## Contributing
Contributions are welcome! Feel free to open an issue or submit a pull request.

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -m 'Add new feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Open a pull request.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.

## Contact
If you have any questions, feel free to reach out at [redon_halimaj@hotmail.com].

