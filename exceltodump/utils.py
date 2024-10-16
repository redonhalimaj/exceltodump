# exceltodump/utils.py

import uuid
import random

def generate_unique_pk():
    """
    Generate a unique primary key using UUID4.

    Returns:
        str: A unique primary key as a hexadecimal string.
    """
    return uuid.uuid4().hex

def generate_uuid():
    """
    Generate a UUID4 string.

    Returns:
        str: A UUID4 string.
    """
    return str(uuid.uuid4())
