# exceltodump/project_dump_updater.py

import xml.etree.ElementTree as ET
import logging
import shutil
import os

def update_project_dump(test_elements_xml, testcase_xml, project_dump_path='project_dump.xml'):
    """
    Update the existing project_dump.xml by replacing <test-elements> and <testcase> sections.

    Args:
        test_elements_xml (Element): The new <test-elements> XML element.
        testcase_xml (Element): The new <testcase> XML element.
        project_dump_path (str): Path to the existing project_dump.xml file.

    Returns:
        None
    """
    try:
        # Backup the existing project_dump.xml
        backup_path = f"{project_dump_path}.bak"
        shutil.copy(project_dump_path, backup_path)
        logging.info(f"Backup created at {backup_path}")

        # Load the existing project_dump.xml
        tree = ET.parse(project_dump_path)
        root = tree.getroot()

        # Replace the <test-elements> section
        for elem in root.findall(".//test-elements"):
            root.remove(elem)
            root.append(test_elements_xml)
            logging.info("<test-elements> section replaced.")

        # Find the <children> node and append the new <testcase>
        children_node = root.find(".//children")
        if children_node is not None:
            children_node.clear()  # Clear existing children
            children_node.append(testcase_xml)  # Append new testcase
            logging.info("<children> node updated with new <testcase>.")
        else:
            logging.warning("No <children> node found in project_dump.xml.")

        # Save the updated project_dump.xml
        tree.write(project_dump_path, encoding="utf-8", xml_declaration=True)
        logging.info(f"project_dump.xml successfully updated.")
    except FileNotFoundError:
        logging.error(f"Error: '{project_dump_path}' file not found.")
        raise
    except Exception as e:
        logging.error(f"Error updating project_dump.xml: {e}")
        raise
