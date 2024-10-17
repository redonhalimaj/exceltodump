# exceltodump/main.py

import argparse
import logging
import sys
import os
from .reader import read_excel
from .converter import (
    generate_test_elements_xml, 
    generate_test_case_xml, 
    update_project_dump, 
    zip_project_dump
)
from xml.etree.ElementTree import tostring
import xml.dom.minidom

def setup_logging():
    """Configures the logging settings."""
    logging.basicConfig(
        level=logging.INFO,  # Change to DEBUG for more detailed logs
        format='%(levelname)s: %(message)s'
    )

def pretty_xml(element):
    """Returns a pretty-printed XML string for the Element."""
    rough_string = tostring(element, 'utf-8')
    reparsed = xml.dom.minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def main():
    setup_logging()

    parser = argparse.ArgumentParser(
        description='Convert Excel test cases to project_dump.xml and generate project-dump.zip.'
    )
    parser.add_argument('excel_file', help='Path to the Excel file containing test cases.')
    parser.add_argument('project_dump', help='Path to the existing project_dump.xml file.')

    args = parser.parse_args()

    excel_file = args.excel_file
    project_dump = args.project_dump

    # Check if Excel file exists
    if not os.path.isfile(excel_file):
        logging.error(f"Excel file '{excel_file}' does not exist.")
        sys.exit(1)

    # Check if project_dump.xml exists
    if not os.path.isfile(project_dump):
        logging.error(f"Project dump file '{project_dump}' does not exist.")
        sys.exit(1)

    try:
        # Step 1: Read and process the Excel file
        logging.info(f"Running reader.py with Excel file '{excel_file}'")
        data = read_excel(excel_file)
        logging.debug(f"Data extracted from Excel: {data}")
    except Exception as e:
        logging.error(f"Failed to read and process Excel file: {e}")
        sys.exit(1)

    try:
        # Step 2: Generate test_elements XML
        logging.info("Generating test elements XML.")
        test_elements_xml, interactions, parameter_mapping, representative_mapping = generate_test_elements_xml(data)
        logging.debug("Test elements XML generated.")
    except Exception as e:
        logging.error(f"Failed to generate test elements XML: {e}")
        sys.exit(1)

    try:
        # Step 3: Serialize and write test_elements XML to file
        logging.info("Writing 'output_test_elements.xml'.")
        with open("output_test_elements.xml", "w", encoding="utf-8") as f:
            f.write(pretty_xml(test_elements_xml))
        logging.info("XML file 'output_test_elements.xml' generated successfully.")
    except Exception as e:
        logging.error(f"Failed to write 'output_test_elements.xml': {e}")
        sys.exit(1)

    try:
        # Step 4: Generate test case XML
        logging.info("Generating test case XML.")
        testcase_xml = generate_test_case_xml(data, interactions, parameter_mapping, representative_mapping)
        logging.debug("Test case XML generated.")
    except Exception as e:
        logging.error(f"Failed to generate test case XML: {e}")
        sys.exit(1)

    try:
        # Step 5: Serialize and write test case XML to file
        logging.info("Writing 'output_testcase.xml'.")
        with open("output_testcase.xml", "w", encoding="utf-8") as f:
            f.write(pretty_xml(testcase_xml))
        logging.info("XML file 'output_testcase.xml' generated successfully.")
    except Exception as e:
        logging.error(f"Failed to write 'output_testcase.xml': {e}")
        sys.exit(1)

    try:
        # Step 6: Update project_dump.xml
        logging.info(f"Updating project dump '{project_dump}'.")
        update_project_dump(test_elements_xml, testcase_xml, project_dump)
    except Exception as e:
        logging.error(f"Failed to update project dump: {e}")
        sys.exit(1)

    try:
        # Step 7: Zip the updated project_dump.xml
        logging.info("Zipping the updated project dump.")
        zip_project_dump(project_dump, 'project-dump.zip')
    except Exception as e:
        logging.error(f"Failed to create zip file: {e}")
        sys.exit(1)

    logging.info("Process completed successfully. 'project-dump.zip' has been created.")

if __name__ == "__main__":
    main()
