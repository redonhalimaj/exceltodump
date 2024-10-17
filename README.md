# exceltodump

**exceltodump** is a command-line tool designed to convert Excel-based test cases into a structured `project_dump.xml` file and package it into a `project-dump.zip` archive. This tool automates the transformation of test case data, ensuring consistency and efficiency in your testing workflows.

## Table of Contents
- [Introduction](#introduction)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Command Syntax](#command-syntax)
  - [Example](#example)
- [File Structure](#file-structure)
- [How It Works](#how-it-works)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Introduction

In software development and quality assurance, managing and importing test cases efficiently is crucial. **exceltodump** streamlines this process by:

- Parsing Excel files containing test case data.
- Categorizing parameters based on their types (e.g., Text, Numeric, Comparison).
- Generating a structured XML file (`project_dump.xml`) compatible with your project management tools.
- Packaging the updated XML into a ZIP archive for easy distribution and import.

## Features

- **Automated Conversion**: Seamlessly convert Excel test cases to XML format.
- **Parameter Categorization**: Automatically categorize parameters as Text, Numeric, or Comparison based on their values.
- **Representative Mapping**: Ensure all parameters reference valid representatives to avoid import warnings.
- **Error Handling**: Provides informative logs and error messages to assist in troubleshooting.
- **Flexible Integration**: Easily integrates into existing workflows and can be invoked from any directory.

## Prerequisites

Before installing and using **exceltodump**, ensure you have the following:

- Python 3.6 or higher installed on your system.
- `pip` (Python package installer) available in your system's PATH.
- Git (optional, for cloning the repository).

## Installation

### 1. Clone the Repository (Optional)
If you haven't already, clone the repository to your local machine:

```bash
git clone https://github.com/redonhalimaj/exceltodump.git
cd exceltodump
```

### 2. Install Dependencies
**exceltodump** relies on the following Python packages:

- `pandas`
- `openpyxl` (for Excel file handling)

Install the dependencies using `pip`:

```bash
pip install pandas openpyxl
```

### 3. Install the Package
Navigate to the root directory of the project (where `setup.py` is located) and install the package in editable mode:

```bash
pip install -e .
```

The `-e` flag installs the package in editable mode, allowing you to modify the code without reinstalling.

### 4. Verify Installation
After installation, verify that the **exceltodump** command is available:

```bash
exceltodump --help
```

You should see the help message detailing usage instructions.

## Usage

**exceltodump** is executed via the command line and requires two positional arguments, one for the Excel file path and the other the corresponding projec-dump.xml.

### Command Syntax

```bash
exceltodump <excel_file> <project-dump.xml>
```
- `<excel_file>`: Path to the Excel file (e.g., `test.xlsx`).

### Example

Assuming you have:

- An Excel file named `test.xlsx` located in `C:\AnyPath`.

Navigate to the directory:

```bash
cd C:\\AnyPath\\testtesttest
```

Run the **exceltodump** command:

```bash
exceltodump test.xlsx
```

**Expected Output**:

```vbnet
INFO: Running reader.py with Excel file 'test.xlsx'
INFO: Excel file 'test.xlsx' loaded successfully.
INFO: Generating test elements XML.
INFO: Generating test case XML.
INFO: Writing XML files.
XML file 'output_test_elements.xml' generated successfully.
XML file 'output_testcase.xml' generated successfully.
INFO: Updating project dump 'project_dump.xml'.
INFO: 'project_dump.xml' successfully updated with new test-elements and testcase.
INFO: Zipping the updated project dump.
INFO: project-dump.zip successfully created.
INFO: Process completed successfully. 'project-dump.zip' has been created.
```

This process will generate:

- `output_test_elements.xml`: Contains the structured test elements.
- `output_testcase.xml`: Contains the structured test case data.
- `project_dump.xml`: Updated with new test elements and test cases.
- `project-dump.zip`: A ZIP archive of the updated `project_dump.xml`.

## File Structure

Here's an overview of the key files and directories in the **exceltodump** project:

```arduino
exceltodump/
├── exceltodump/
│   ├── __init__.py
│   ├── reader.py
│   ├── converter.py
│   └── main.py
├── tests/
│   └── test_exceltodump.py
├── setup.py
├── requirements.txt
└── README.md
```
- **exceltodump/**: Contains the main Python modules.
  - **reader.py**: Handles reading and parsing the Excel file.
  - **converter.py**: Handles XML generation and project dump updates.
  - **main.py**: Entry point for the command-line interface.
- **tests/**: Contains unit tests.
- **setup.py**: Configuration for package installation.
- **requirements.txt**: Lists Python dependencies.
- **README.md**: Documentation file (this file).

## How It Works

### 1. Reading the Excel File
The tool reads the specified Excel file using `pandas`, extracting data from the Precondition, Action, and Expected Result columns. It uses regular expressions to identify operations (`#op[Op Name](Op Param)`) and parameters (`#p[Param Name]`).

### 2. Categorizing Parameters
Parameters are categorized based on their values:

- **Numeric**: Values matching numeric patterns (e.g., `123`, `45.67`).
- **Text**: Alphanumeric identifiers or specific string patterns (e.g., `"A2"`, `"TV_Signal1"`).
- **Comparison**: Common comparison operators enclosed in quotes (e.g., `"=="`, `"!="`).
- **Empty**: Parameters with no value.

### 3. Generating XML Elements
Using the categorized parameters, the tool generates XML elements for data types (`datatype`) and interactions (`interaction`). Each data type has representatives to ensure valid references within the XML.

### 4. Updating `project_dump.xml`
The existing `project_dump.xml` is updated by replacing the `<test-elements>` section with the newly generated XML and appending the new test cases.

### 5. Packaging into ZIP
Finally, the updated `project_dump.xml` is zipped into `project-dump.zip` for easy distribution and import.

## Troubleshooting

During the import process, you might encounter warnings like:

```vbnet
import warning: affected object: interaction-call-parameter-datatype (PK: 35514498747060388)  - potential problem: unable to find referred representative in imported project
```

This indicates that the XML references a representative that hasn't been defined. To resolve this:

### 1. Ensure All Data Types Have Representatives
Modify the converter to define representatives for all data types used in your test cases. For example:

- **Text**: Add a default representative like `"Auto_Param_Text"`.
- **Numeric**: Add a default representative like `"Auto_Param_Numeric"`.
- **Comparison**: Add all necessary comparison operators as representatives (`"=="`, `"!="`, etc.).

### 2. Verify Representative Mapping
Ensure that the `representative_mapping` correctly maps each category and value to their corresponding PKs. This mapping is crucial for the call-parameter elements to reference existing representatives.

### 3. Consistent Naming Conventions
Maintain consistent naming conventions across your Excel data, scripts, and XML references. Avoid discrepancies like using hyphens (`-`) vs. underscores (`_`) in filenames and identifiers.

### 4. Increase Logging for Debugging
To gain more insight into the process, adjust the logging level in `main.py` to `DEBUG`:

```python
import logging

def setup_logging():
    """Configures the logging settings."""
    logging.basicConfig(
        level=logging.DEBUG,  # Set to DEBUG for detailed logs
        format='%(levelname)s: %(message)s'
    )
```

This will provide more detailed logs, helping you identify where the mapping might be failing.

### 5. Reinstall the Package After Changes
After making changes to the scripts, reinstall the package to ensure the latest code is in effect:

```bash
pip install -e .
```

### 6. Validate `project_dump.xml` Structure
Ensure that the existing `project_dump.xml` has the expected structure, especially the `<test-elements>` and `<children>` nodes. Any deviation can cause the script to malfunction.

## Contributing

Contributions are welcome! If you encounter bugs, have feature requests, or wish to improve the documentation, please follow these steps:

### Fork the Repository
Click the "Fork" button on the repository page to create a personal copy.

### Create a Branch

```bash
git checkout -b feature/YourFeatureName
```

### Make Changes
Implement your changes or fixes.

### Commit Your Changes

```bash
git commit -m "Add Your Feature Description"
```


### Create a Pull Request
Navigate to the original repository and submit a pull request detailing your changes.

## License

This project is licensed under the MIT License.

## Contact

For any questions, issues, or suggestions, please open an issue in the repository or contact the maintainer at redon_halimaj@hotmail.com.

**Happy Testing!**

