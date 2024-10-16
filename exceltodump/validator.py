# exceltodump/validator.py
import xmlschema
import logging

def validate_xml(xml_path, xsd_path='project_dump_schema.xsd'):
    """
    Validate an XML file against an XSD schema.

    Args:
        xml_path (str): Path to the XML file.
        xsd_path (str): Path to the XSD schema file.

    Returns:
        bool: True if valid, False otherwise.
    """
    try:
        schema = xmlschema.XMLSchema(xsd_path)
        schema.validate(xml_path)
        logging.info(f"XML file '{xml_path}' is valid against the schema.")
        return True
    except xmlschema.exceptions.XMLSchemaValidationError as e:
        logging.error(f"XML validation error: {e}")
        return False
    except FileNotFoundError:
        logging.error(f"XSD schema file '{xsd_path}' not found.")
        return False
    except Exception as e:
        logging.error(f"Error during XML validation: {e}")
        return False
