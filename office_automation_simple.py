#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office Automation - Simplified Main Entry Point

A robust implementation that handles import issues gracefully.
"""

import os
import sys
from pathlib import Path
from typing import Dict, Any, Optional, List
import json


# Add project root to Python path for imports
_project_root = Path(__file__).parent
sys.path.insert(0, str(_project_root))


def import_core_module(module_name, class_name):
    """Robust module import with fallbacks."""
    try:
        # Try absolute import
        module = __import__(f'core.{module_name}', fromlist=[class_name])
        return getattr(module, class_name)
    except ImportError:
        try:
            # Try direct file import
            import importlib.util
            module_path = _project_root / 'core' / f'{module_name}.py'
            spec = importlib.util.spec_from_file_location(
                f'core.{module_name}', 
                str(module_path)
            )
            if spec and spec.loader:
                module = importlib.util.module_from_spec(spec)
                sys.modules[f'core.{module_name}'] = module
                spec.loader.exec_module(module)
                return getattr(module, class_name)
        except Exception:
            pass
    
    # Fallback to dummy class
    print(f"⚠️ Warning: Using dummy {class_name} (core.{module_name} not found)")
    
    class DummyClass:
        def __init__(self, config=None):
            self.config = config or {}
            self.available = False
        
        def get_capabilities(self):
            return {'available': False, 'dummy': True}
        
        def __getattr__(self, name):
            # Return a dummy method for any attribute access
            def dummy_method(*args, **kwargs):
                print(f"⚠️ {self.__class__.__name__}.{name}() is not available")
                return None
            return dummy_method
    
    return DummyClass


def import_utils_module(module_name, class_name=None):
    """Robust utility module import."""
    try:
        module = __import__(f'utils.{module_name}', fromlist=[class_name] if class_name else [])
        if class_name:
            return getattr(module, class_name)
        return module
    except ImportError:
        print(f"⚠️ Warning: utils.{module_name} not found")
        
        if class_name == 'OfficeAutomationError':
            return Exception
        elif class_name == 'handle_office_error':
            def dummy_decorator(func):
                return func
            return dummy_decorator
        elif class_name == 'TemplateManager':
            class DummyTemplateManager:
                def __init__(self, config=None): pass
            return DummyTemplateManager
        elif class_name == 'BatchProcessor':
            class DummyBatchProcessor:
                def __init__(self, **kwargs): pass
            return DummyBatchProcessor
        else:
            return None


# Import core modules
WordProcessor = import_core_module('word_processor', 'WordProcessor')
ExcelProcessor = import_core_module('excel_processor', 'ExcelProcessor')
PowerPointProcessor = import_core_module('powerpoint_processor', 'PowerPointProcessor')
FormatConverter = import_core_module('format_converter', 'FormatConverter')
WPSIntegration = import_core_module('wps_integration', 'WPSIntegration')

# Import utility modules
OfficeAutomationError = import_utils_module('error_handler', 'OfficeAutomationError')
handle_office_error = import_utils_module('error_handler', 'handle_office_error')
TemplateManager = import_utils_module('templates', 'TemplateManager')
BatchProcessor = import_utils_module('batch_processor', 'BatchProcessor')


class OfficeAutomation:
    """Main Office Automation class with robust error handling."""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or {}
        self._setup_config()
        
        # Initialize modules
        self.word = WordProcessor(config=self.config.get('word', {}))
        self.excel = ExcelProcessor(config=self.config.get('excel', {}))
        self.powerpoint = PowerPointProcessor(config=self.config.get('powerpoint', {}))
        self.converter = FormatConverter(config=self.config.get('converter', {}))
        self.wps = WPSIntegration(config=self.config.get('wps', {}))
        
        # Initialize utilities
        self.templates = TemplateManager(config=self.config.get('templates', {}))
        
        # Batch processor needs the other processors
        batch_config = self.config.get('batch', {})
        self.batch = BatchProcessor(
            word_processor=self.word,
            excel_processor=self.excel,
            powerpoint_processor=self.powerpoint,
            converter=self.converter,
            config=batch_config
        )
        
        print(f"Office Automation initialized")
        if hasattr(self.wps, 'available') and self.wps.available:
            print(f"  WPS Office: Available ({getattr(self.wps, 'version', 'unknown')})")
        else:
            print(f"  WPS Office: Not available")
    
    def _setup_config(self):
        """Set up default configuration."""
        defaults = {
            'default_templates': {
                'report': 'templates/report.docx',
                'invoice': 'templates/invoice.xlsx',
                'presentation': 'templates/presentation.pptx'
            },
            'temp_dir': 'temp',
            'encoding': 'utf-8'
        }
        
        # Merge with user config
        for key, value in defaults.items():
            if key not in self.config:
                self.config[key] = value
            elif isinstance(value, dict) and isinstance(self.config[key], dict):
                self.config[key] = {**value, **self.config[key]}
    
    def get_info(self) -> Dict[str, Any]:
        """Get system information."""
        info = {
            'version': '1.0.0',
            'project_root': str(_project_root),
            'python_version': sys.version,
            'modules': {}
        }
        
        # Check each module
        for name, module in [
            ('word', self.word),
            ('excel', self.excel),
            ('powerpoint', self.powerpoint),
            ('converter', self.converter),
            ('wps', self.wps)
        ]:
            if hasattr(module, 'get_capabilities'):
                info['modules'][name] = module.get_capabilities()
            else:
                info['modules'][name] = {'available': False}
        
        return info
    
    def test_functionality(self) -> Dict[str, bool]:
        """Test basic functionality of all modules."""
        tests = {}
        
        # Test Word
        try:
            if hasattr(self.word, 'test'):
                tests['word'] = self.word.test()
            else:
                tests['word'] = False
        except:
            tests['word'] = False
        
        # Test Excel
        try:
            if hasattr(self.excel, 'test'):
                tests['excel'] = self.excel.test()
            else:
                tests['excel'] = False
        except:
            tests['excel'] = False
        
        # Test PowerPoint
        try:
            if hasattr(self.powerpoint, 'test'):
                tests['powerpoint'] = self.powerpoint.test()
            else:
                tests['powerpoint'] = False
        except:
            tests['powerpoint'] = False
        
        # Test WPS
        tests['wps'] = getattr(self.wps, 'available', False)
        
        return tests


# Quick access functions
def quick_create_document(output_path: str, content: str = '') -> bool:
    """Quick document creation."""
    try:
        office = OfficeAutomation()
        return office.word.create_document(output_path, content)
    except:
        return False


def quick_create_spreadsheet(output_path: str, data: List[List] = None) -> bool:
    """Quick spreadsheet creation."""
    try:
        office = OfficeAutomation()
        return office.excel.create_workbook(output_path, data)
    except:
        return False


def quick_convert(input_path: str, output_format: str) -> bool:
    """Quick format conversion."""
    try:
        office = OfficeAutomation()
        return office.converter.convert(input_path, output_format)
    except:
        return False


if __name__ == "__main__":
    # Simple CLI
    import argparse
    
    parser = argparse.ArgumentParser(description='Office Automation CLI')
    parser.add_argument('--info', action='store_true', help='Show system info')
    parser.add_argument('--test', action='store_true', help='Test functionality')
    
    args = parser.parse_args()
    
    office = OfficeAutomation()
    
    if args.info:
        info = office.get_info()
        print(json.dumps(info, indent=2, ensure_ascii=False))
    
    elif args.test:
        tests = office.test_functionality()
        print("Functionality Tests:")
        for module, result in tests.items():
            status = "✅" if result else "❌"
            print(f"  {status} {module}: {'Available' if result else 'Not available'}")
    
    else:
        print("Office Automation - Ready")
        print("Usage:")
        print("  python office_automation.py --info   # Show system info")
        print("  python office_automation.py --test   # Test functionality")
        print("\nQuick functions:")
        print("  from office_automation import quick_create_document")
        print("  from office_automation import quick_create_spreadsheet")
        print("  from office_automation import quick_convert")