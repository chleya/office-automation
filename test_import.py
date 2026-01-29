#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test import for Office Automation
"""

import os
import sys

# Add current directory to path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

print(f"Current directory: {current_dir}")
print(f"Python path: {sys.path[:3]}")

# Test importing core modules
modules_to_test = [
    'core.word_processor',
    'core.excel_processor', 
    'core.powerpoint_processor',
    'core.format_converter',
    'core.wps_integration',
    'utils.error_handler',
    'utils.templates',
    'utils.batch_processor'
]

for module_path in modules_to_test:
    try:
        # Convert path to module import
        if module_path.startswith('core.'):
            module_name = module_path.replace('core.', '')
            __import__(f'core.{module_name}')
            print(f"✅ {module_path}: Import successful")
        elif module_path.startswith('utils.'):
            module_name = module_path.replace('utils.', '')
            __import__(f'utils.{module_name}')
            print(f"✅ {module_path}: Import successful")
    except Exception as e:
        print(f"❌ {module_path}: {e}")

# Test main import
print("\n" + "="*50)
print("Testing main office_automation import...")

try:
    from office_automation import OfficeAutomation
    print("✅ office_automation import successful!")
    
    office = OfficeAutomation()
    print("✅ OfficeAutomation instance created!")
    
    info = office.get_info()
    print(f"✅ System info retrieved: version {info['version']}")
    
except Exception as e:
    print(f"❌ office_automation import failed: {e}")
    import traceback
    traceback.print_exc()