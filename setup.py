from setuptools import setup, find_packages
# To use a consistent encoding
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='SpreadsheetTo',
    version='0.1.0',

    description='Spreadsheet to [CSV | JSON | ... ] Extractor',
    long_description=long_description,
    url='https://github.com/timfreund/SpreadsheetTo',
    author='Tim Freund',
    author_email='tim@freunds.net',
    license='MIT',

    # See https://pypi.python.org/pypi?%3Aaction=list_classifiers
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
    ],

    # What does your project relate to?
    keywords='opendata',

    packages=find_packages(exclude=['contrib', 'docs', 'tests']),

    install_requires=[],

    entry_points={
        'console_scripts': [
            'spreadsheet-to=spreadsheetto:cli',
        ],
    },
)
