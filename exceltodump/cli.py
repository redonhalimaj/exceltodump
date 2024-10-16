# exceltodump/cli.py

import argparse
import logging
from .reader import read_excel
from .json_generator import generate_json
from .xml_generator import generate_test_elements_xml, prettify_xml
from .project_dump_updater import update_project_dump
from .zipper import zip_project_dump
from .validator import validate_xml

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(description="Convert Excel test cases to a zipped project-dump.xml.")
    parser.add_argument('excel_file', type=str, help='Path to the Excel file (e.g., test.xlsx)')
    parser.add_argument('--project-dump', type=str, default='project_dump.xml', help='Path to the existing project_dump.xml file')
    parser.add_argument('--output-json', type=str, default='output.json', help='Path to the intermediate JSON file')
    parser.add_argument('--output-xml', type=str, default='output_test_elements.xml', help='Path to the intermediate test elements XML file')
    parser.add_argument('--output-testcase', type=str, default='output_testcase.xml', help='Path to the intermediate testcase XML file')
    parser.add_argument('--zip-path', type=str, default='project_dump.zip', help='Path for the output zip file')
    parser.add_argument('--xsd-schema', type=str, default='project_dump_schema.xsd', help='Path to the XSD schema file for validation')
    return parser.parse_args()

def main():
    """Main function to execute the conversion process."""
    # Configure logging
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    # Parse arguments
    args = parse_arguments()

    # Step 1: Read and parse Excel
    try:
        output_data = read_excel(args.excel_file)
    except Exception:
        logging.error("Failed to read and parse the Excel file.")
        return

    # Step 2: Generate JSON
    try:
        generate_json(output_data, args.output_json)
    except Exception:
        logging.error("Failed to generate JSON data.")
        return

    # Step 3: Generate XML
    try:
        test_elements_xml, interactions, parameter_mapping, representative_mapping = generate_test_elements_xml(output_data)
        test_elements_pretty = prettify_xml(test_elements_xml)
    except Exception:
        logging.error("Failed to generate test elements XML.")
        return

    # Save test elements XML to file (optional)
    try:
        with open(args.output_xml, 'w', encoding='utf-8') as f:
            f.write(test_elements_pretty)
        logging.info(f"Test elements XML saved at '{args.output_xml}'.")
    except Exception as e:
        logging.error(f"Error saving test elements XML: {e}")
        return

    # Step 4: Generate Testcase XML (Assuming you have a separate function for this)
    # For simplicity, let's assume the testcase XML is part of test_elements.xml
    # If separate, implement similar to test_elements_xml

    # Step 5: Update project_dump.xml
    try:
        update_project_dump(test_elements_xml, None, args.project_dump)  # Replace 'None' with testcase_xml if separate
    except Exception:
        logging.error("Failed to update project_dump.xml.")
        return

    # Step 6: Validate the updated project_dump.xml
    is_valid = validate_xml(args.project_dump, args.xsd_schema)
    if not is_valid:
        logging.error("Validation failed. Please check the XML structure.")
        return

    # Step 7: Zip the updated project_dump.xml
    try:
        zip_project_dump(args.project_dump, args.zip_path)
    except Exception:
        logging.error("Failed to zip the project_dump.xml.")
        return

    logging.info("Conversion process completed successfully.")

if __name__ == "__main__":
    main()
