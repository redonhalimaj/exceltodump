# exceltodump/xml_generator.py

import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import logging
from .utils import generate_unique_pk, generate_uuid

def prettify_xml(elem):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def create_datatype_element(name, representatives, datatype_mapping, representative_mapping):
    """
    Create a <datatype> XML element.

    Args:
        name (str): Name of the datatype.
        representatives (list): List of representative values.
        datatype_mapping (dict): Mapping from datatype names to PKs.
        representative_mapping (dict): Mapping from datatype to representative name to PK.

    Returns:
        Element: The <element type="datatype"> XML element.
    """
    datatype = ET.Element("element", type="datatype")
    pk = generate_unique_pk()
    uid = f"iTB-DT-{pk[:6]}"
    ET.SubElement(datatype, "pk").text = pk
    ET.SubElement(datatype, "name").text = name
    ET.SubElement(datatype, "uid").text = uid
    ET.SubElement(datatype, "locker")
    ET.SubElement(datatype, "status").text = "3"
    ET.SubElement(datatype, "description").text = f"Datatype for {name}"
    ET.SubElement(datatype, "html-description").text = "<html><body></body></html>"
    ET.SubElement(datatype, "historyPK").text = pk
    ET.SubElement(datatype, "identicalVersionPK").text = "-1"
    ET.SubElement(datatype, "references")
    ET.SubElement(datatype, "kind").text = "regular"
    ET.SubElement(datatype, "fields")
    ET.SubElement(datatype, "instances-arrays")

    equivalence_classes = ET.SubElement(datatype, "equivalence-classes")
    eq_class_elem = ET.SubElement(equivalence_classes, "equivalence-class")
    eq_class_pk = generate_unique_pk()
    ET.SubElement(eq_class_elem, "pk").text = eq_class_pk
    ET.SubElement(eq_class_elem, "name").text = name
    ET.SubElement(eq_class_elem, "description").text = name
    ET.SubElement(eq_class_elem, "ordering").text = "1024"

    default_rep_pk = None

    representatives_elem = ET.SubElement(eq_class_elem, "representatives")
    for i, representative in enumerate(representatives, start=1):
        representative_elem = ET.SubElement(representatives_elem, "representative")
        rep_pk = generate_unique_pk()
        ET.SubElement(representative_elem, "pk").text = rep_pk
        rep_name = representative.strip('"').strip()
        if rep_name:
            ET.SubElement(representative_elem, "name").text = rep_name
        else:
            ET.SubElement(representative_elem, "name")
        ET.SubElement(representative_elem, "ordering").text = str(i * 1024)
        ET.SubElement(representative_elem, "type").text = "text"
        ET.SubElement(representative_elem, "values")
        if i == 1:
            default_rep_pk = rep_pk

        # Store representative mapping
        if name not in representative_mapping:
            representative_mapping[name] = {}
        representative_mapping[name][rep_name] = rep_pk

    if default_rep_pk:
        default_rep_ref = ET.SubElement(eq_class_elem, "default-representative-ref")
        default_rep_ref.set("pk", default_rep_pk)

    ET.SubElement(datatype, "old-versions")

    # Store the mapping of datatype name to its PK
    datatype_mapping[name] = pk

    return datatype

def generate_test_elements_xml(data):
    """
    Generate the <test-elements> XML structure.

    Args:
        data (dict): The input data containing test elements and parameters.

    Returns:
        Element: The <test-elements> XML element.
        list: List of interactions for the testcase.
        dict: Parameter mapping.
        dict: Representative mapping.
    """
    test_elements = ET.Element("test-elements")
    datatype_mapping = {}  # Mapping from datatype names to PKs
    representative_mapping = {}  # Mapping from datatype to representative name to PK
    parameter_mapping = {}    # Mapping from parameter PK to parameter details
    interactions_list = []

    # Create Datatypes Subdivision
    datatype_subdivision = ET.SubElement(test_elements, "element", type="subdivision")
    ET.SubElement(datatype_subdivision, "pk").text = generate_unique_pk()
    ET.SubElement(datatype_subdivision, "name").text = "Datatypes"
    ET.SubElement(datatype_subdivision, "uid").text = f"iTB-SD-{generate_unique_pk()[:6]}"
    ET.SubElement(datatype_subdivision, "locker")
    ET.SubElement(datatype_subdivision, "description")
    ET.SubElement(datatype_subdivision, "html-description").text = "<html><body></body></html>"
    ET.SubElement(datatype_subdivision, "historyPK").text = generate_unique_pk()
    ET.SubElement(datatype_subdivision, "identicalVersionPK").text = "-1"
    ET.SubElement(datatype_subdivision, "references")
    ET.SubElement(datatype_subdivision, "old-versions")

    # Generate Datatypes from Generated_Parameters
    generated_params = data.get("Generated_Parameters", {})
    for param_category, values in generated_params.items():
        # Ensure 'Empty' and 'Text' datatypes are present
        if param_category not in ["Empty", "Text"]:
            continue
        representatives = values.copy()
        if param_category != "Empty" and "" not in representatives:
            representatives.append("")  # Include empty string for non-empty datatypes

        datatype_elem = create_datatype_element(
            name=param_category,
            representatives=representatives,
            datatype_mapping=datatype_mapping,
            representative_mapping=representative_mapping
        )
        datatype_subdivision.append(datatype_elem)

    # Create Subdivisions for 'Precondition', 'Action', 'Expected_Result'
    sections = ['Precondition', 'Action', 'Expected_Result']
    subdivisions = {}
    for section in sections:
        subdivision = ET.SubElement(test_elements, "element", type="subdivision")
        ET.SubElement(subdivision, "pk").text = generate_unique_pk()
        ET.SubElement(subdivision, "name").text = section
        ET.SubElement(subdivision, "uid").text = f"iTB-SD-{generate_unique_pk()[:6]}"
        ET.SubElement(subdivision, "locker")
        ET.SubElement(subdivision, "description")
        ET.SubElement(subdivision, "html-description").text = "<html><body></body></html>"
        ET.SubElement(subdivision, "historyPK").text = generate_unique_pk()
        ET.SubElement(subdivision, "identicalVersionPK").text = "-1"
        ET.SubElement(subdivision, "references")
        ET.SubElement(subdivision, "old-versions")
        subdivisions[section] = subdivision

    # Iterate through each row and generate interactions
    for row_key, row_data in data.items():
        if row_key == "Generated_Parameters":
            continue  # Skip parameter definitions
        test_elements_data = row_data.get("test-elements", {})
        testcase_ops = row_data.get("testcase", [])
        for section in sections:
            section_data = test_elements_data.get(section.replace(' ', '_'), {})
            operations = section_data.get("Operations", [])
            for operation in operations:
                operation_name = operation.get("operation")
                parameters = operation.get("parameters", [])
                param_details = operation.get("param_details", [])

                # Create Interaction Element
                interaction_elem = ET.Element("element", type="interaction")
                interaction_pk = generate_unique_pk()
                interaction_uid = f"iTB-IA-{generate_unique_pk()[:6]}"
                ET.SubElement(interaction_elem, "pk").text = interaction_pk
                ET.SubElement(interaction_elem, "name").text = operation_name
                ET.SubElement(interaction_elem, "uid").text = interaction_uid
                ET.SubElement(interaction_elem, "locker")
                ET.SubElement(interaction_elem, "status").text = "3"
                default_call_type = ET.SubElement(interaction_elem, "default-call-type")
                default_call_type.set("name", "Flow")
                default_call_type.set("value", "0")
                ET.SubElement(interaction_elem, "description")
                ET.SubElement(interaction_elem, "html-description").text = "<html><body></body></html>"
                ET.SubElement(interaction_elem, "historyPK").text = interaction_pk
                ET.SubElement(interaction_elem, "identicalVersionPK").text = "-1"
                ET.SubElement(interaction_elem, "references")
                ET.SubElement(interaction_elem, "preconditions")
                ET.SubElement(interaction_elem, "postconditions")

                # Parameters
                parameters_elem = ET.SubElement(interaction_elem, "parameters")
                param_pks = []
                for idx, param_placeholder in enumerate(parameters):
                    param_elem = ET.SubElement(parameters_elem, "parameter")
                    param_pk = generate_unique_pk()
                    ET.SubElement(param_elem, "pk").text = param_pk
                    param_name = f"Param{idx+1}"
                    ET.SubElement(param_elem, "name").text = param_name
                    # Determine datatype-ref based on placeholder
                    datatype_ref = ET.SubElement(param_elem, "datatype-ref")
                    if param_placeholder.startswith("Auto_Param_Text"):
                        datatype_ref.set("pk", datatype_mapping.get("Text", ""))
                    elif param_placeholder.startswith("Auto_Param_Numeric"):
                        datatype_ref.set("pk", datatype_mapping.get("Numeric", ""))
                    elif param_placeholder.startswith("Auto_Param_Comparison"):
                        datatype_ref.set("pk", datatype_mapping.get("Comparison", ""))
                    else:
                        datatype_ref.set("pk", datatype_mapping.get("Empty", ""))
                    ET.SubElement(param_elem, "definition-type").text = "0"
                    ET.SubElement(param_elem, "use-type").text = "1"
                    signature_uid = generate_uuid()
                    ET.SubElement(param_elem, "signature-uid").text = signature_uid

                    # Store parameter PKs and details
                    param_pks.append({
                        'pk': param_pk,
                        'name': param_name,
                        'datatype_pk': datatype_ref.get("pk"),
                        'signature_uid': signature_uid
                    })

                ET.SubElement(interaction_elem, "call-sequence")
                ET.SubElement(interaction_elem, "old-versions")

                # Append interaction to the appropriate subdivision
                subdivision = subdivisions.get(section)
                if subdivision is not None:
                    subdivision.append(interaction_elem)
                else:
                    logging.warning(f"Subdivision '{section}' not found.")

                # Collect interaction data for testcase
                interactions_list.append({
                    'name': operation_name,
                    'pk': interaction_pk,
                    'parameters': param_pks,
                    'phase': 'Setup' if section == 'Precondition' else 'TestStep' if section == 'Action' else 'Teardown',
                    'param_details': param_details
                })

                # Map parameter PKs to their details
                for param_pk_info, param_detail in zip(param_pks, param_details):
                    param_pk = param_pk_info['pk']
                    parameter_mapping[param_pk] = {
                        'pk': param_pk,
                        'name': param_pk_info['name'],
                        'datatype_category': param_detail.get('category', 'Empty'),
                        'value': param_detail.get('value', ''),
                        'signature_uid': param_pk_info['signature_uid']
                    }

    return test_elements, interactions_list, parameter_mapping, representative_mapping
