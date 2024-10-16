# exceltodump/zipper.py

import zipfile
import os
import logging

def zip_project_dump(project_dump_path='project_dump.xml', zip_path='project_dump.zip'):
    """
    Compress the project_dump.xml into a zip file.

    Args:
        project_dump_path (str): Path to the project_dump.xml file.
        zip_path (str): Desired path for the output zip file.

    Returns:
        None
    """
    try:
        if os.path.exists(zip_path):
            logging.warning(f"{zip_path} already exists and will be overwritten.")

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(project_dump_path, os.path.basename(project_dump_path))
        logging.info(f"{zip_path} successfully created.")
    except Exception as e:
        logging.error(f"Error creating zip file '{zip_path}': {e}")
        raise
