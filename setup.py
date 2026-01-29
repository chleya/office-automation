#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Setup script for Office Automation library.
"""

from setuptools import setup, find_packages
import os

# Read the README file
with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

# Read requirements
with open('requirements.txt', 'r', encoding='utf-8') as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

# Get version from the package
def get_version():
    """Get version from office_automation module."""
    import re
    with open('office_automation.py', 'r', encoding='utf-8') as f:
        content = f.read()
        match = re.search(r"__version__\s*=\s*['\"]([^'\"]+)['\"]", content)
        if match:
            return match.group(1)
    return '1.0.0'

setup(
    name='office-automation',
    version=get_version(),
    author='Office Automation Team',
    author_email='office-automation@example.com',
    description='A comprehensive Python library for automating Microsoft Office and WPS Office tasks',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/yourusername/office-automation',
    packages=find_packages(exclude=['tests', 'tests.*', 'examples', 'examples.*']),
    include_package_data=True,
    package_data={
        'office_automation': ['py.typed'],
    },
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Topic :: Office/Business :: Office Suites',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Operating System :: OS Independent',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: MacOS',
        'Operating System :: POSIX :: Linux',
    ],
    python_requires='>=3.8',
    install_requires=requirements,
    extras_require={
        'dev': [
            'pytest>=7.0.0',
            'pytest-cov>=4.0.0',
            'black>=23.0.0',
            'isort>=5.12.0',
            'flake8>=6.0.0',
            'mypy>=1.0.0',
            'twine>=4.0.0',
            'wheel>=0.40.0',
        ],
        'docs': [
            'sphinx>=7.0.0',
            'sphinx-rtd-theme>=1.3.0',
            'sphinx-autodoc-typehints>=1.25.0',
        ],
        'full': [
            'python-docx>=1.1.0',
            'openpyxl>=3.1.0',
            'python-pptx>=0.6.23',
            'pandas>=2.0.0',
            'numpy>=1.24.0',
            'reportlab>=4.0.0',
        ],
    },
    entry_points={
        'console_scripts': [
            'office-automation=office_automation.cli:main',
        ],
    },
    project_urls={
        'Bug Reports': 'https://github.com/yourusername/office-automation/issues',
        'Source': 'https://github.com/yourusername/office-automation',
        'Documentation': 'https://office-automation.readthedocs.io/',
    },
    keywords=[
        'office',
        'automation',
        'microsoft-office',
        'wps-office',
        'word',
        'excel',
        'powerpoint',
        'document',
        'spreadsheet',
        'presentation',
        'python',
    ],
)