#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test running an example file
"""

import os
import sys

# Add project root to path
project_root = r'F:\skill\office-automation'
sys.path.insert(0, project_root)

print("Testing example file execution...")
print("=" * 50)

# Test create_report.py
print("\n1. Testing create_report.py...")
try:
    # Import the example module
    import examples.create_report as report_example
    
    # Run the main function
    success = report_example.main()
    
    if success:
        print("create_report.py executed successfully!")
    else:
        print("create_report.py failed")
        
except Exception as e:
    print(f"Error running create_report.py: {e}")

# Test process_spreadsheet.py
print("\n2. Testing process_spreadsheet.py...")
try:
    import examples.process_spreadsheet as spreadsheet_example
    
    success = spreadsheet_example.main()
    
    if success:
        print("process_spreadsheet.py executed successfully!")
    else:
        print("process_spreadsheet.py failed")
        
except Exception as e:
    print(f"Error running process_spreadsheet.py: {e}")

print("\n" + "=" * 50)
print("Test completed!")