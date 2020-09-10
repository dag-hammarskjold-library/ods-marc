import sys
from setuptools import setup, find_packages

with open("README.md") as f:
    long_description = f.read()
    
with open("requirements.txt") as f:
    requirements = list(filter(None,f.read().split('\n')))

setup(
    name = 'ods-marc',
    version = '0.1',
    url = 'http://github.com/dag-hammarskjold-library/ods-marc',
    author = 'United Nations Dag Hammarskjöld Library',
    author_email = 'library-ny@un.org',
    license = 'http://www.opensource.org/licenses/bsd-license.php',
    packages = find_packages(exclude=['test']),
    #test_suite = 'tests',
    #install_requires = requirements,
    description = 'Convert ODS data to MARC records',
    long_description = long_description,
    long_description_content_type = "text/markdown",
    python_requires = '>=3.6',
    entry_points = {
        'console_scripts': [
            'ods-marc=ods_marc:run'
        ]
    }
)

