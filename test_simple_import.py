#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simple import test for Office Automation
"""

import os
import sys

# Add project root to path
project_root = r'F:\skill\office-automation'
sys.path.insert(0, project_root)

print("Testing Office Automation import...")
print("=" * 50)

try:
    from office_automation import OfficeAutomation
    print("Import successful!")
    
    # Create instance
    office = OfficeAutomation()
    print("Instance created!")
    
    # Get info
    info = office.get_info()
    print(f"System info retrieved")
    print(f"   Version: {info['version']}")
    print(f"   Python: {info['python_version'].split()[0]}")
    
    # Test modules
    print("\nModule status:")
    for module_name, module_info in info['modules'].items():
        status = "[OK]" if module_info.get('available', False) else "[NO]"
        print(f"   {status} {module_name}: {module_info}")
    
    # Test quick functions
    print("\nTesting quick functions...")
    
    # Test document creation
    try:
        from office_automation import quick_create_document
        print("quick_create_document imported")
    except Exception as e:
        print(f"quick_create_document: {e}")
    
    # Test spreadsheet creation
    try:
        from office_automation import quick_create_spreadsheet
        print("quick_create_spreadsheet imported")
    except Exception as e:
        print(f"quick_create_spreadsheet: {e}")
    
    # Test conversion
    try:
        from office_automation import quick_convert
        print("quick_convert imported")
    except Exception as e:
        print(f"quick_convert: {e}")
    
    print("\nAll tests completed!")
    
except Exception as e:
    print(f"Import failed: {e}")
    import traceback
    traceback.print_exc()