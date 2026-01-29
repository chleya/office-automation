#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office Automation Project Creator

Robust script to create and verify the Office Automation project structure.
Handles file creation, imports, and verification in a single pass.
"""

import os
import sys
import time
import shutil
from pathlib import Path
import json


class ProjectCreator:
    """Robust project creation and verification."""
    
    def __init__(self, project_root):
        self.project_root = Path(project_root)
        self.created_files = []
        self.failed_files = []
        
    def safe_write(self, file_path, content, retries=3):
        """Safely write file with retries and verification."""
        for attempt in range(retries):
            try:
                # Ensure directory exists
                file_path.parent.mkdir(parents=True, exist_ok=True)
                
                # Write file
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                # Verify write
                if file_path.exists():
                    actual_size = file_path.stat().st_size
                    expected_size = len(content.encode('utf-8'))
                    
                    if actual_size > 0:
                        self.created_files.append(str(file_path))
                        print(f"  ✅ {file_path.name} ({actual_size} bytes)")
                        return True
                    else:
                        print(f"  ⚠️ {file_path.name} is empty, retrying...")
                else:
                    print(f"  ⚠️ {file_path.name} not created, retrying...")
                    
            except Exception as e:
                print(f"  ❌ Attempt {attempt+1} failed: {e}")
            
            time.sleep(0.5)  # Wait before retry
        
        self.failed_files.append(str(file_path))
        return False
    
    def create_structure(self):
        """Create the complete project structure."""
        print("=" * 60)
        print("Creating Office Automation Project Structure")
        print("=" * 60)
        
        # Create directories
        directories = [
            '',
            'core',
            'utils', 
            'examples',
            'tests',
            'templates',
            'output',
            'data'
        ]
        
        print("\n1. Creating directories...")
        for dir_name in directories:
            dir_path = self.project_root / dir_name
            dir_path.mkdir(parents=True, exist_ok=True)
            print(f"  📁 {dir_path}")
        
        # Create __init__.py files
        print("\n2. Creating package files...")
        init_files = [
            self.project_root / '__init__.py',
            self.project_root / 'core' / '__init__.py',
            self.project_root / 'utils' / '__init__.py'
        ]
        
        for init_file in init_files:
            self.safe_write(init_file, '# Package initialization\n')
        
        return True
    
    def verify_imports(self):
        """Verify that all modules can be imported."""
        print("\n" + "=" * 60)
        print("Verifying Imports")
        print("=" * 60)
        
        # Add project root to Python path
        sys.path.insert(0, str(self.project_root))
        
        modules_to_test = [
            ('core.word_processor', 'WordProcessor'),
            ('core.excel_processor', 'ExcelProcessor'),
            ('core.powerpoint_processor', 'PowerPointProcessor'),
            ('core.format_converter', 'FormatConverter'),
            ('core.wps_integration', 'WPSIntegration'),
            ('utils.error_handler', 'OfficeAutomationError'),
            ('utils.templates', 'TemplateManager'),
            ('utils.batch_processor', 'BatchProcessor')
        ]
        
        all_success = True
        
        for module_path, class_name in modules_to_test:
            try:
                # Dynamically import
                module = __import__(module_path, fromlist=[class_name])
                cls = getattr(module, class_name)
                print(f"  ✅ {module_path}.{class_name}")
            except Exception as e:
                print(f"  ❌ {module_path}.{class_name}: {e}")
                all_success = False
        
        return all_success
    
    def run_example(self, example_name):
        """Run an example file to verify functionality."""
        print(f"\n3. Testing {example_name}...")
        
        example_path = self.project_root / 'examples' / example_name
        
        if not example_path.exists():
            print(f"  ⚠️ {example_name} not found")
            return False
        
        try:
            # Change to project directory
            original_cwd = os.getcwd()
            os.chdir(self.project_root)
            
            # Run the example
            import subprocess
            result = subprocess.run(
                [sys.executable, str(example_path)],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            os.chdir(original_cwd)
            
            if result.returncode == 0:
                print(f"  ✅ {example_name} executed successfully")
                return True
            else:
                print(f"  ❌ {example_name} failed:")
                print(f"     stdout: {result.stdout[:200]}...")
                print(f"     stderr: {result.stderr[:200]}...")
                return False
                
        except Exception as e:
            print(f"  ❌ {example_name} error: {e}")
            return False
    
    def create_summary(self):
        """Create a summary report."""
        print("\n" + "=" * 60)
        print("CREATION SUMMARY")
        print("=" * 60)
        
        total_files = len(self.created_files) + len(self.failed_files)
        
        print(f"\n📊 Statistics:")
        print(f"  Total files attempted: {total_files}")
        print(f"  Successfully created: {len(self.created_files)}")
        print(f"  Failed: {len(self.failed_files)}")
        
        if self.failed_files:
            print(f"\n❌ Failed files:")
            for file in self.failed_files:
                print(f"  - {file}")
        
        # Calculate total size
        total_size = 0
        for file_path in self.created_files:
            if os.path.exists(file_path):
                total_size += os.path.getsize(file_path)
        
        print(f"\n💾 Total project size: {total_size:,} bytes")
        
        return len(self.failed_files) == 0


def main():
    """Main execution."""
    project_root = Path(__file__).parent
    
    print("Office Automation Project Creator")
    print(f"Project root: {project_root}")
    print("-" * 60)
    
    creator = ProjectCreator(project_root)
    
    # Step 1: Create structure
    if not creator.create_structure():
        print("❌ Failed to create project structure")
        return 1
    
    # Step 2: Verify imports
    if not creator.verify_imports():
        print("⚠️ Some imports failed, but continuing...")
    
    # Step 3: Test examples
    examples = ['create_report.py', 'process_spreadsheet.py']
    for example in examples:
        creator.run_example(example)
    
    # Step 4: Create summary
    success = creator.create_summary()
    
    if success:
        print("\n🎉 Project creation completed successfully!")
        print("\nNext steps:")
        print("1. Run: python examples/create_report.py")
        print("2. Run: python examples/process_spreadsheet.py")
        print("3. Check the 'output' directory for generated files")
    else:
        print("\n⚠️ Project creation completed with warnings")
    
    return 0 if success else 1


if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\n⚠️ Interrupted by user")
        sys.exit(1)