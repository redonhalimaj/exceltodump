# setup.py

from setuptools import setup, find_packages

setup(
    name='exceltodump',
    version='1.0.0',
    description='Convert Excel test cases to a zipped project-dump.xml for import into the testing tool.',
    author='Redon Halimaj',
    author_email='redon_halimaj@hotmail.com',
    packages=find_packages(),
    install_requires=[
        'pandas',
        'xmlschema',
    ],
    entry_points={
        'console_scripts': [
            'exceltodump=exceltodump.main:main',
        ],
    },
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
    ],
)
