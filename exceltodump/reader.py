# exceltodump/reader.py

import json
import pandas as pd
import re
import logging

# Define regex patterns
op_pattern = r"#op\[(.*?)\]\((.*?)\)"       # Match operations like #op[Op Name](Op Param)
param_pattern = r"#p\[(.*?)\]"            # Match parameters like #p[Param Name]
interaction_pattern = r"\[(.*?)\]"       # Match anything inside square brackets []

def categorize_param_value(param_value):
    """Categorize parameter value based on its type."""
    param_value = param_value.strip()
    if re.match(r"^[0-9]+(\.[0-9]+)?$", param_value):  # Match numbers (integers/floats)
        return 'Numeric'
    elif re.match(r'^"A[0-9]+"$', param_value):        # Match signal identifiers (e.g., "A2", "A3")
        return 'Text'
    elif re.match(r'^"TV_[A-Za-z0-9_]+"$', param_value):  # Match signal/variable names like "TV_Signal1"
        return 'Text'
    elif re.match(r'^(\d+ sec|[0-9]+ min)$', param_value):  # Match time or duration (e.g., "10 sec")
        return 'Text'
    elif re.match(r'^"(==|!=|>=|<=|>|<)"$', param_value):  # Match comparison operators enclosed in quotes
        return 'Comparison'
    else:
        return 'Text'

def extract_operations_params(text, categorized_params):
    if not isinstance(text, str):
        return [], [], categorized_params  # Return empty lists if the text is not a valid string

    # Extract operations and parameters
    operations = re.findall(op_pattern, text)
    parameters = re.findall(param_pattern, text)

    # Extract all items inside square brackets
    all_descriptions = re.findall(interaction_pattern, text)

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
                if param not in categorized_params["Text"]:
                    categorized_params["Text"].append(param)
                param_placeholders.append("Auto_Param_Text")
            elif category == 'Numeric':
                if param not in categorized_params["Numeric"]:
                    categorized_params["Numeric"].append(param)
                param_placeholders.append("Auto_Param_Numeric")
            elif category == 'Comparison':
                if param not in categorized_params["Comparison"]:
                    categorized_params["Comparison"].append(param)
                param_placeholders.append("Auto_Param_Comparison")
            else:
                # Fallback to 'Text' for any unforeseen categories
                if param not in categorized_params["Text"]:
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
    """
    Reads the Excel file and processes it to generate output_data.

    Args:
        file_path (str): Path to the Excel file.

    Returns:
        dict: Processed data.
    """
    try:
        excel_data = pd.read_excel(file_path)
        logging.info(f"Excel file '{file_path}' loaded successfully.")
    except FileNotFoundError:
        logging.error(f"Excel file '{file_path}' not found.")
        raise
    except Exception as e:
        logging.error(f"Error loading Excel file '{file_path}': {e}")
        raise

    # Prepare dictionary to store the data
    output_data = {}

    # Global categorized parameters
    categorized_params = {
        "Text": [],
        "Numeric": [],
        "Comparison": []
    }

    # Apply extraction to the 'Precondition', 'Action', and 'Expected Result' columns
    for index, row in excel_data.iterrows():
        row_data = {}
        test_elements = {}
        test_case_operations = []

        # Process Precondition
        if isinstance(row.get('Precondition'), str) and row['Precondition'].strip():
            descriptions, ops, categorized_params = extract_operations_params(
                row['Precondition'], categorized_params)
            test_elements['Precondition'] = {
                "Descriptions": descriptions,
                "Operations": ops  # Keep operations as is for detailed info
            }
            # Collect operations for testcase
            test_case_operations.extend(ops)

        # Process Action
        if isinstance(row.get('Action'), str) and row['Action'].strip():
            descriptions, ops, categorized_params = extract_operations_params(
                row['Action'], categorized_params)
            test_elements['Action'] = {
                "Descriptions": descriptions,
                "Operations": ops  # Keep operations as is for detailed info
            }
            # Collect operations for testcase
            test_case_operations.extend(ops)

        # Process Expected Result
        if isinstance(row.get('Expected Result'), str) and row['Expected Result'].strip():
            descriptions, ops, categorized_params = extract_operations_params(
                row['Expected Result'], categorized_params)
            test_elements['Expected_Result'] = {
                "Descriptions": descriptions,
                "Operations": ops  # Keep operations as is for detailed info
            }
            # Collect operations for testcase
            test_case_operations.extend(ops)

        # Add test-elements and testcase to row_data
        row_data['test-elements'] = test_elements
        row_data['testcase'] = test_case_operations

        # Add row data to output if it contains any extracted information
        if row_data:
            output_data[f"Row_{index}"] = row_data

    # Add the grouped global parameters to the output data
    output_data["Generated_Parameters"] = {
        key: value for key, value in categorized_params.items() if value  # Include only non-empty categories
    }

    return output_data
