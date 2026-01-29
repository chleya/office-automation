"""
Pytest configuration for Office Automation tests
"""

import os
import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest


@pytest.fixture
def test_data_dir():
    """Fixture for test data directory."""
    data_dir = project_root / "tests" / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


@pytest.fixture
def test_output_dir():
    """Fixture for test output directory."""
    output_dir = project_root / "tests" / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


@pytest.fixture
def cleanup_test_files():
    """Fixture to clean up test files after tests."""
    files_to_clean = []
    
    def add_file(file_path):
        files_to_clean.append(file_path)
    
    yield add_file
    
    # Cleanup after test
    for file_path in files_to_clean:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except:
                pass