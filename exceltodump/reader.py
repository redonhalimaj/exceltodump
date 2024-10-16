# exceltodump/reader.py

import pandas as pd
import re
import logging
from collections import defaultdict

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Define regex patterns
OP_PATTERN = re.compile(r"#op\[(.*?)\]\((.*?)\)")
PARAM_PATTERN = re.compile(r"#p\[(.*?)\]")
INTERACTION_PATTERN = re.compile(r"\[(.*?)\]")

def categorize_param_value(param_value):
    """Categorize parameter value based on its type."""
    param_value = param_value.strip()
    if re.match(r"^[0-9]+(\.[0-9]+)?$", param_value):  # Numeric
        return 'Numeric'
    elif re.match(r'^"A[0-9]+"$', param_value):  # Text (e.g., "A2")
        return 'Text'
    elif re.match(r'^"TV_[A-Za-z0-9_]+"$', param_value):  # Text (e.g., "TV_Signal1")
        return 'Text'
    elif re.match(r'^(\d+ sec|[0-9]+ min)$', param_value):  # Text (e.g., "10 sec")
        return 'Text'
    elif re.match(r'^"(==|!=|>=|<=|>|<)"$', param_value):  # Comparison
        return 'Comparison'
    else:
        logging.warning(f"Uncategorized parameter value: {param_value}. Defaulting to 'Text'.")
        return 'Text'

def extract_operations_params(text, categorized_params):
    """Extract operations and parameters from the given text and categorize them."""
    if not isinstance(text, str):
        return [], [], categorized_params  # Return empty lists if text is not valid

    # Extract operations and parameters
    operations = OP_PATTERN.findall(text)
    parameters = PARAM_PATTERN.findall(text)

    # Extract all items inside square brackets
    all_descriptions = INTERACTION_PATTERN.findall(text)

    # Remove descriptions that are operations (#op[...]) or parameters (#p[...])
    descriptions = [
        interact for interact in all_descriptions 
        if f"#op[{interact}]" not in text and f"#p[{interact}]" not in text and len(interact) > 1
    ]

    # Clean up operations by removing the #p part inside them
    cleaned_operations = []

    for op in operations:
        operation_name, param_value = op
        param_value_cleaned = re.sub(r"#p\[(.*?)\]", r"\1", param_value)  # Clean the parameter

        method_params = re.split(r',\s*', param_value_cleaned)  # Split method params (like "A2", 0, ...)

        param_placeholders = []
        param_details = []

        # Categorize each parameter and add to global categorized_params
        for param in method_params:
            param = param.strip()
            if param == '':
                param_placeholders.append("")  # Handle empty parameters
                param_details.append({"category": "Empty", "value": ""})
                continue
            category = categorize_param_value(param)
            if category == 'Text':
                categorized_params["Text"].append(param)
                param_placeholders.append("Auto_Param_Text")
            elif category == 'Numeric':
                categorized_params["Numeric"].append(param)
                param_placeholders.append("Auto_Param_Numeric")
            elif category == 'Comparison':
                categorized_params["Comparison"].append(param)
                param_placeholders.append("Auto_Param_Comparison")
            else:
                # Fallback to 'Text' for any unforeseen categories
                categorized_params["Text"].append(param)
                param_placeholders.append("Auto_Param_Text")
            param_details.append({"category": category, "value": param})

        # Pad with empty strings to ensure 5 parameters
        while len(param_placeholders) < 5:
            param_placeholders.append("")
            param_details.append({"category": "Empty", "value": ""})

        cleaned_operations.append({
            "operation": operation_name,
            "parameters": param_placeholders,  # List of parameter placeholders
            "param_details": param_details     # Detailed parameter info
        })

    return descriptions, cleaned_operations, categorized_params

def read_excel(file_path):
    """Read and process the Excel file to extract test elements and test cases."""
    try:
        excel_data = pd.read_excel(file_path)
        logging.info(f"Excel file '{file_path}' loaded successfully.")
    except FileNotFoundError:
        logging.error(f"Error: Excel file '{file_path}' not found.")
        raise
    except Exception as e:
        logging.error(f"Error reading Excel file '{file_path}': {e}")
        raise

    # Prepare dictionary to store the data
    output_data = {}

    # Global categorized parameters using sets for uniqueness
    categorized_params = {
        "Text": set(),
        "Numeric": set(),
        "Comparison": set()
    }

    # Columns to process
    columns_to_process = ['Precondition', 'Action', 'Expected Result']

    # Apply extraction to the specified columns
    for index, row in excel_data.iterrows():
        row_data = {}
        test_elements = {}
        test_case_operations = []

        for column in columns_to_process:
            cell_content = row.get(column)
            if isinstance(cell_content, str) and cell_content.strip():
                descriptions, ops, categorized_params = extract_operations_params(
                    cell_content, categorized_params)
                test_elements[column.replace(' ', '_')] = {
                    "Descriptions": descriptions,
                    "Operations": ops  # Keep operations as is for detailed info
                }
                # Collect operations for testcase
                test_case_operations.extend(ops)

        # Add test-elements and testcase to row_data
        row_data['test-elements'] = test_elements
        row_data['testcase'] = test_case_operations

        # Add row data to output if it contains any extracted information
        if test_elements or test_case_operations:
            output_data[f"Row_{index}"] = row_data

    # Convert sets to sorted lists
    output_data["Generated_Parameters"] = {
        key: sorted(list(value)) for key, value in categorized_params.items() if value
    }

    return output_data
