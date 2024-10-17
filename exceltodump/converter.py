# exceltodump/converter.py

import json
import random
import zipfile
import os 
import uuid
from collections import defaultdict
from xml.etree.ElementTree import Element, SubElement, tostring, parse
import xml.dom.minidom
import logging

def generate_unique_pk():
    """Generates a unique primary key."""
    return str(random.randint(10**16, 10**17 - 1))

def generate_uuid():
    """Generates a UUID."""
    return str(uuid.uuid4())

def create_datatype_element(name, pk, uid, representatives, datatype_mapping, representative_mapping):
    """Creates a datatype XML element."""
    datatype = Element("element", type="datatype")
    SubElement(datatype, "pk").text = pk
    SubElement(datatype, "name").text = name
    SubElement(datatype, "uid").text = uid
    SubElement(datatype, "locker")
    SubElement(datatype, "status").text = "3"
    SubElement(datatype, "description").text = f"Datatype for {name}"
    SubElement(datatype, "html-description").text = "<html><body></body></html>"
    SubElement(datatype, "historyPK").text = pk
    SubElement(datatype, "identicalVersionPK").text = "-1"
    SubElement(datatype, "references")
    SubElement(datatype, "kind").text = "regular"
    SubElement(datatype, "fields")
    SubElement(datatype, "instances-arrays")

    equivalence_classes = SubElement(datatype, "equivalence-classes")
    eq_class_elem = SubElement(equivalence_classes, "equivalence-class")
    eq_class_pk = generate_unique_pk()
    SubElement(eq_class_elem, "pk").text = eq_class_pk
    SubElement(eq_class_elem, "name").text = name
    SubElement(eq_class_elem, "description").text = name
    SubElement(eq_class_elem, "ordering").text = "1024"
    # Prepare to set the default representative PK
    default_rep_pk = None

    representatives_elem = SubElement(eq_class_elem, "representatives")
    for i, representative in enumerate(representatives, start=1):
        representative_elem = SubElement(representatives_elem, "representative")
        rep_pk = generate_unique_pk()
        SubElement(representative_elem, "pk").text = rep_pk
        rep_name = representative.strip('"').strip()  # Remove quotes and trim whitespace
        # Handle empty representative names
        if rep_name == "":
            SubElement(representative_elem, "name")
        else:
            SubElement(representative_elem, "name").text = rep_name
        SubElement(representative_elem, "ordering").text = str(i * 1024)
        SubElement(representative_elem, "type").text = "text"
        SubElement(representative_elem, "values")
        if i == 1:
            default_rep_pk = rep_pk  # Set the first representative as default

        # Store representative mapping
        if name not in representative_mapping:
            representative_mapping[name] = {}
        representative_mapping[name][rep_name] = rep_pk

    if default_rep_pk:
        default_rep_ref = SubElement(eq_class_elem, "default-representative-ref")
        default_rep_ref.set("pk", default_rep_pk)

    SubElement(datatype, "old-versions")

    # Store the mapping of datatype name to its PK
    datatype_mapping[name] = pk

    return datatype

# exceltodump/converter.py

# ... [Previous imports and functions] ...

# Main function to generate the test elements XML structure
def generate_test_elements_xml(data):
    test_elements = Element("test-elements")
    datatype_mapping = {}  # Mapping from datatype names to PKs
    representative_mapping = {}  # Mapping from datatype to representative name to PK
    parameter_mapping = {}    # Mapping from parameter PK to parameter details
    operation_parameters = {}  # Mapping from operation name to parameters and their categories

    # Process Data Types first to build datatype_mapping
    datatype_subdivision = SubElement(test_elements, "element", type="subdivision")
    SubElement(datatype_subdivision, "pk").text = "307965"
    SubElement(datatype_subdivision, "name").text = "Datatypes"
    SubElement(datatype_subdivision, "uid").text = "iTB-SD-307965"
    SubElement(datatype_subdivision, "locker")
    SubElement(datatype_subdivision, "description")
    SubElement(datatype_subdivision, "html-description").text = "<html><body></body></html>"
    SubElement(datatype_subdivision, "historyPK").text = "307965"
    SubElement(datatype_subdivision, "identicalVersionPK").text = "-1"
    SubElement(datatype_subdivision, "references")
    SubElement(datatype_subdivision, "old-versions")

    generated_params = data.get('Generated_Parameters', {})
    for param_name, param_values in generated_params.items():
        # Ensure empty strings are included in representatives if needed
        if "" not in param_values:
            param_values.append("")  # Include empty string
        datatype_pk = generate_unique_pk()
        datatype_elem = create_datatype_element(
            param_name,
            datatype_pk,
            "iTB-DT-" + datatype_pk[-6:],
            param_values,
            datatype_mapping,
            representative_mapping
        )
        datatype_subdivision.append(datatype_elem)

    # Ensure "Empty" datatype is always created
    if "Empty" not in datatype_mapping:
        empty_datatype_pk = generate_unique_pk()
        empty_datatype_elem = create_datatype_element(
            "Empty",
            empty_datatype_pk,
            "iTB-DT-" + empty_datatype_pk[-6:],
            [""],  # Representative is empty string
            datatype_mapping,
            representative_mapping
        )
        datatype_subdivision.append(empty_datatype_elem)

    # Create a generic "Text" datatype with a default representative
    if "Text" not in datatype_mapping:
        text_datatype_pk = generate_unique_pk()
        text_datatype_elem = create_datatype_element(
            "Text",
            text_datatype_pk,
            "iTB-DT-" + text_datatype_pk[-6:],
            ["Auto_Param_Text"],  # Add a default representative
            datatype_mapping,
            representative_mapping
        )
        datatype_subdivision.append(text_datatype_elem)

    # Create a "Numeric" datatype with a default representative
    if "Numeric" not in datatype_mapping:
        numeric_datatype_pk = generate_unique_pk()
        numeric_datatype_elem = create_datatype_element(
            "Numeric",
            numeric_datatype_pk,
            "iTB-DT-" + numeric_datatype_pk[-6:],
            ["Auto_Param_Numeric"],  # Add a default representative
            datatype_mapping,
            representative_mapping
        )
        datatype_subdivision.append(numeric_datatype_elem)

    # Create a "Comparison" datatype with all comparison operators as representatives
    if "Comparison" not in datatype_mapping:
        comparison_datatype_pk = generate_unique_pk()
        comparison_representatives = ['==', '!=', '>=', '<=', '>', '<']
        comparison_datatype_elem = create_datatype_element(
            "Comparison",
            comparison_datatype_pk,
            "iTB-DT-" + comparison_datatype_pk[-6:],
            comparison_representatives,  # Add all comparison operators
            datatype_mapping,
            representative_mapping
        )
        datatype_subdivision.append(comparison_datatype_elem)

    # Now process the interactions and collect interaction PKs and parameter PKs for test cases
    interactions = []
    interaction_mapping = {}  # Mapping from operation name to interaction data

    # First pass: Collect all parameters for each operation
    for row_key in data:
        if row_key == 'Generated_Parameters':
            continue  # Skip Generated_Parameters, already processed
        row_data = data[row_key]

        test_elements_data = row_data.get('test-elements', {})
        sections = ['Precondition', 'Action', 'Expected_Result']
        for section in sections:
            operations = test_elements_data.get(section, {}).get('Operations', [])
            for operation in operations:
                operation_name = operation['operation']
                parameters = operation.get('parameters', [])
                param_details_list = operation.get('param_details', [])
                if operation_name not in operation_parameters:
                    operation_parameters[operation_name] = {
                        'param_positions': defaultdict(set),
                        'calls': [],
                        'sections': set()
                    }
                # Record parameter categories at each position
                for idx, param_detail in enumerate(param_details_list):
                    category = param_detail.get('category', 'Empty')
                    operation_parameters[operation_name]['param_positions'][idx].add(category)
                # Store the operation call for later
                operation_parameters[operation_name]['calls'].append({
                    'section': section,
                    'operation': operation_name,
                    'parameters': parameters,
                    'param_details': param_details_list
                })
                # Record the section
                operation_parameters[operation_name]['sections'].add(section)

    # Create subdivisions for 'Precondition', 'Action', 'Expected_Result'
    subdivisions = {}
    for section in ['Precondition', 'Action', 'Expected_Result']:
        subdivision = SubElement(test_elements, "element", type="subdivision")
        SubElement(subdivision, "pk").text = generate_unique_pk()
        SubElement(subdivision, "name").text = section
        SubElement(subdivision, "uid").text = "iTB-SD-" + generate_unique_pk()[-6:]
        SubElement(subdivision, "locker")
        SubElement(subdivision, "description")
        SubElement(subdivision, "html-description").text = "<html><body></body></html>"
        SubElement(subdivision, "historyPK").text = generate_unique_pk()
        SubElement(subdivision, "identicalVersionPK").text = "-1"
        SubElement(subdivision, "references")
        SubElement(subdivision, "old-versions")
        subdivisions[section] = subdivision

    # Second pass: Create interaction elements
    for operation_name, op_data in operation_parameters.items():
        interaction_pk = generate_unique_pk()
        interaction_uid = "iTB-IA-" + interaction_pk[-6:]

        # Create interaction element
        interaction_elem = Element("element", type="interaction")
        SubElement(interaction_elem, "pk").text = interaction_pk
        SubElement(interaction_elem, "name").text = operation_name
        SubElement(interaction_elem, "uid").text = interaction_uid
        SubElement(interaction_elem, "locker").text = ""
        SubElement(interaction_elem, "status").text = "3"
        default_call_type = SubElement(interaction_elem, "default-call-type")
        default_call_type.set("name", "Flow")
        default_call_type.set("value", "0")
        SubElement(interaction_elem, "description").text = ""
        SubElement(interaction_elem, "html-description").text = "<html><body></body></html>"
        SubElement(interaction_elem, "historyPK").text = interaction_pk
        SubElement(interaction_elem, "identicalVersionPK").text = "-1"
        SubElement(interaction_elem, "references")
        SubElement(interaction_elem, "preconditions")
        SubElement(interaction_elem, "postconditions")

        # Parameters
        parameters_elem = SubElement(interaction_elem, "parameters")
        param_pks = []
        max_params = max(op_data['param_positions'].keys()) + 1  # +1 because index starts at 0
        for idx in range(max_params):
            param_pk = generate_unique_pk()
            param_elem = SubElement(parameters_elem, "parameter")
            SubElement(param_elem, "pk").text = param_pk
            param_elem_name = f"Param{idx+1}"
            SubElement(param_elem, "name").text = param_elem_name
            # Set datatype-ref
            datatype_ref = SubElement(param_elem, "datatype-ref")
            categories = op_data['param_positions'][idx]
            if len(categories) == 1:
                category = next(iter(categories))
                datatype_pk = datatype_mapping.get(category, datatype_mapping['Empty'])
            else:
                # Multiple categories, use 'Text' datatype
                datatype_pk = datatype_mapping['Text']
            datatype_ref.set("pk", datatype_pk)
            SubElement(param_elem, "definition-type").text = "0"
            SubElement(param_elem, "use-type").text = "1"
            signature_uid = generate_uuid()
            SubElement(param_elem, "signature-uid").text = signature_uid

            # Store parameter PKs and signature UIDs for mapping
            param_pks.append({
                'pk': param_pk,
                'name': param_elem_name,
                'datatype_pk': datatype_pk,
                'signature_uid': signature_uid
            })

        SubElement(interaction_elem, "call-sequence")
        SubElement(interaction_elem, "old-versions")

        # Determine primary section
        primary_section = next(iter(op_data['sections']))

        # Store interaction data
        interaction_mapping[operation_name] = {
            'pk': interaction_pk,
            'uid': interaction_uid,
            'element': interaction_elem,
            'param_pks': param_pks,
            'primary_section': primary_section
        }

    # Append interaction elements to the appropriate subdivisions
    for operation_name, interaction_data in interaction_mapping.items():
        primary_section = interaction_data['primary_section']
        subdivision = subdivisions[primary_section]
        subdivision.append(interaction_data['element'])

    # Build interactions list for test case
    interactions = []
    for operation_name, op_data in operation_parameters.items():
        interaction_info = interaction_mapping[operation_name]
        for call in op_data['calls']:
            interactions.append({
                'name': operation_name,
                'pk': interaction_info['pk'],
                'parameters': interaction_info['param_pks'],
                'phase': 'Setup' if call['section'] == 'Precondition' else 'TestStep' if call['section'] == 'Action' else 'Teardown',
                'param_details': call['param_details']
            })
            # Map parameter PKs to their details
            for param_pk_info, param_detail in zip(interaction_info['param_pks'], call['param_details']):
                param_pk = param_pk_info['pk']
                parameter_mapping[param_pk] = {
                    'pk': param_pk,
                    'name': param_pk_info['name'],
                    'datatype_category': param_detail.get('category', 'Empty'),
                    'value': param_detail.get('value', ''),
                    'signature_uid': param_pk_info['signature_uid']
                }

    return test_elements, interactions, parameter_mapping, representative_mapping


# Function to generate the test case XML
def generate_test_case_xml(data, interactions, parameter_mapping, representative_mapping):
    testcase_pk, testcase_uid = generate_unique_pk(), generate_unique_pk()[-6:]

    # Create the root element
    testcase = Element('testcase')

    # Add the test case pk
    SubElement(testcase, 'pk').text = str(testcase_pk)

    # Add the test case name
    SubElement(testcase, 'name').text = "Generated Test Case"

    # Add order-pos
    SubElement(testcase, 'order-pos').text = '1024'

    # Add uid
    SubElement(testcase, 'uid').text = f'iTB-TC-{testcase_uid}'

    # Specification element
    specification = SubElement(testcase, 'specification')

    # Details element
    details = SubElement(specification, 'details')
    SubElement(details, 'version')
    details_pk = generate_unique_pk()
    SubElement(details, 'pk').text = details_pk
    SubElement(details, 'identicalVersionPK').text = '-1'
    SubElement(details, 'locker')
    SubElement(details, 'responsible')
    SubElement(details, 'reviewer')
    SubElement(details, 'priority').text = '0'
    SubElement(details, 'status').text = '2'
    SubElement(details, 'target-date')
    SubElement(details, 'references')

    # Add description elements
    SubElement(specification, 'description')
    html_desc = SubElement(specification, 'html-description')
    html_desc.text = '<html><body></body></html>'
    SubElement(specification, 'review-comments')
    html_review_comments = SubElement(specification, 'html-review-comments')
    html_review_comments.text = '<html><body></body></html>'
    SubElement(specification, 'keywords')
    SubElement(specification, 'edited-requirements')
    SubElement(specification, 'non-edited-requirements')
    SubElement(specification, 'userDefinedFields')

    # Interaction element within specification
    interaction = SubElement(specification, 'interaction')

    # Generate a unique pk for the interaction
    interaction_pk, interaction_uid = generate_unique_pk(), generate_unique_pk()[-6:]

    SubElement(interaction, 'pk').text = str(interaction_pk)
    SubElement(interaction, 'name').text = "Generated Test Case Interaction"
    SubElement(interaction, 'uid').text = f'iTB-IA-{interaction_uid}'
    SubElement(interaction, 'locker')
    SubElement(interaction, 'status').text = '3'
    default_call_type = SubElement(interaction, 'default-call-type')
    default_call_type.set("name", "Flow")
    default_call_type.set("value", "0")
    SubElement(interaction, 'description')
    html_interaction_desc = SubElement(interaction, 'html-description')
    html_interaction_desc.text = '<html><body></body></html>'
    SubElement(interaction, 'historyPK').text = str(interaction_pk)
    SubElement(interaction, 'identicalVersionPK').text = "-1"
    SubElement(interaction, 'references')
    SubElement(interaction, 'preconditions')
    SubElement(interaction, 'postconditions')
    parameters_elem = SubElement(interaction, 'parameters')

    # Call sequence
    call_sequence = SubElement(interaction, 'call-sequence')

    # Sort interactions by phase and order (if needed)
    phase_order = {'Setup': 0, 'TestStep': 1, 'Teardown': 2}
    interactions_sorted = sorted(interactions, key=lambda x: (phase_order.get(x['phase'], 99), interactions.index(x)))

    # Create interaction-call elements
    for interaction_info in interactions_sorted:
        interaction_call = SubElement(call_sequence, 'interaction-call')
        SubElement(interaction_call, 'interaction-ref', pk=str(interaction_info['pk']))
        SubElement(interaction_call, 'description')
        html_call_desc = SubElement(interaction_call, 'html-description')
        html_call_desc.text = '<html><body></body></html>'
        SubElement(interaction_call, 'comment')
        html_call_comment = SubElement(interaction_call, 'html-comment')
        html_call_comment.text = '<html><body></body></html>'
        SubElement(interaction_call, 'type').text = '0'
        SubElement(interaction_call, 'phase').text = interaction_info['phase']
        parameter_values = SubElement(interaction_call, 'parameter-values')

        # Add parameter values
        param_details_list = interaction_info['param_details']
        for param_info, param_detail in zip(interaction_info['parameters'], param_details_list):
            param_pk = param_info['pk']
            call_parameter = SubElement(parameter_values, 'call-parameter', default_value="false", type="representative")
            SubElement(call_parameter, 'pk').text = generate_unique_pk()
            SubElement(call_parameter, 'parameter-datatype-ref', pk=param_pk)
            category = param_detail.get('category', 'Empty')
            value = param_detail.get('value', '')
            # Handle empty parameters
            if value is None:
                value = ''
            value = value.strip().strip('"')  # Trim whitespace and remove surrounding quotes
            # Get the representative PK
            rep_pk = representative_mapping.get(category, {}).get(value)
            if not rep_pk:
                # If representative not found, use Empty datatype's representative
                rep_pk = representative_mapping.get('Empty', {}).get('', '')
                if not rep_pk:
                    rep_pk = generate_unique_pk()  # Generate a new PK if Empty datatype not found
            SubElement(call_parameter, 'representative-ref', pk=rep_pk)
        SubElement(interaction_call, 'marker')

    # Add parameter-combinations
    parameter_combinations = SubElement(specification, 'parameter-combinations')
    parameter_combination = SubElement(parameter_combinations, 'parameter-combination')
    param_comb_pk = generate_unique_pk()
    SubElement(parameter_combination, 'pk').text = param_comb_pk
    SubElement(parameter_combination, 'comment')
    html_param_comb_comment = SubElement(parameter_combination, 'html-comment')
    html_param_comb_comment.text = '<html><body></body></html>'
    SubElement(parameter_combination, 'ordering').text = '1024'
    SubElement(parameter_combination, 'uid').text = f'iTB-TC-{testcase_uid}-PC-{param_comb_pk[-6:]}'
    SubElement(parameter_combination, 'keywords')
    SubElement(parameter_combination, 'edited-requirements')
    SubElement(parameter_combination, 'non-edited-requirements')
    SubElement(parameter_combination, 'userDefinedFields')
    values_elem = SubElement(parameter_combination, 'values')

    # Add values for the test case parameters (if any)
    # For simplicity, we can leave this empty or define as needed

    # Add old-versions
    SubElement(specification, 'old-versions')

    # Automation element
    automation = SubElement(testcase, 'automation')
    auto_details = SubElement(automation, 'details')
    SubElement(auto_details, 'version')
    auto_details_pk = generate_unique_pk()
    SubElement(auto_details, 'pk').text = auto_details_pk
    SubElement(auto_details, 'identicalVersionPK').text = "-1"
    SubElement(auto_details, 'locker')
    SubElement(auto_details, 'responsible')
    SubElement(auto_details, 'reviewer')
    SubElement(auto_details, 'priority').text = '0'
    SubElement(auto_details, 'status').text = '6'
    SubElement(auto_details, 'target-date')
    SubElement(auto_details, 'references')
    SubElement(automation, 'script-editor')
    SubElement(automation, 'script-template')
    SubElement(automation, 'old-versions')

    # Add execution-cycles
    SubElement(testcase, 'execution-cycles')

    return testcase


def update_project_dump(test_elements_xml, testcase_xml, project_dump_path='project-dump.xml'):
    """Replaces test-elements and testcase in project_dump.xml."""
    try:
        # Load the existing project_dump.xml
        tree = parse(project_dump_path)
        root = tree.getroot()

        # Replace the <test-elements> section
        test_elements_found = False
        for parent in root.findall(".//test-elements/.."):
            for elem in parent.findall("test-elements"):
                parent.remove(elem)
                test_elements_found = True

        if not test_elements_found:
            logging.warning("No 'test-elements' found in the project_dump.xml")

        # Append new <test-elements>
        test_elements_node = parse("output_test_elements.xml").getroot()
        root.append(test_elements_node)

        # Find the children node where test cases should go
        children_found = False
        for children_node in root.findall(".//children"):
            # Replace or add <testcase> within the <children> node
            children_node.clear()  # Clear existing children
            testcase_node = parse("output_testcase.xml").getroot()
            children_node.append(testcase_node)  # Append new test cases
            children_found = True

        if not children_found:
            logging.warning("No 'children' node found in the project_dump.xml")

        # Save the updated project_dump.xml
        tree.write(project_dump_path, encoding="utf-8", xml_declaration=True)
        logging.info(f"'{project_dump_path}' successfully updated with new test-elements and testcase.")

    except FileNotFoundError:
        logging.error(f"File '{project_dump_path}' not found.")
    except Exception as e:
        logging.error(f"Error updating '{project_dump_path}': {e}")

def zip_project_dump(project_dump_path='project-dump.xml', zip_path='project-dump.zip'):
    """Zips the project_dump.xml into a zip file."""
    try:
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(project_dump_path, os.path.basename(project_dump_path))
        logging.info(f"{zip_path} successfully created.")
    except Exception as e:
        logging.error(f"Error creating zip file: {e}")
