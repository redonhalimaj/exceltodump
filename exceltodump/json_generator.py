# exceltodump/json_generator.py

import json
import logging

def generate_json(output_data, json_file_path):
    """Generate and save JSON data to a file."""
    try:
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(output_data, json_file, indent=4, ensure_ascii=False)
        logging.info(f"JSON file saved at: {json_file_path}")
    except Exception as e:
        logging.error(f"Error writing JSON file '{json_file_path}': {e}")
        raise
