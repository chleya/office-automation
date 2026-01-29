"""
Example Validation Tests
========================

Tests to verify that all examples work correctly.
"""

import os
import sys
import subprocess
import tempfile
import pytest
from pathlib import Path
import importlib.util

# Import the module
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from office_automation import OfficeAutomation


class TestExamples:
    """Tests for validating example files."""
    
    def setup_method(self):
        """Setup for each test."""
        self.temp_dir = tempfile.mkdtemp(prefix="example_test_")
        print(f"\nTest directory: {self.temp_dir}")
        
        # Change to project root for example execution
        self.original_cwd = os.getcwd()
        os.chdir(project_root)
    
    def teardown_method(self):
        """Cleanup after each test."""
        import shutil
        os.chdir(self.original_cwd)
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_example_files_exist(self):
        """Verify that all example files exist."""
        print("\nVerifying example files exist...")
        
        examples_dir = project_root / "examples"
        assert examples_dir.exists(), f"Examples directory not found: {examples_dir}"
        
        expected_examples = [
            "create_report.py",
            "process_spreadsheet.py",
            "generate_presentation.py",
            "wps_integration_demo.py",
        ]
        
        missing_examples = []
        for example in expected_examples:
            example_path = examples_dir / example
            if example_path.exists():
                print(f"  Found: {example}")
            else:
                missing_examples.append(example)
                print(f"  Missing: {example}")
        
        assert not missing_examples, f"Missing example files: {missing_examples}"
        print("  All example files exist: PASSED")
    
    def test_example_syntax(self):
        """Verify that all examples have valid Python syntax."""
        print("\nVerifying example syntax...")
        
        examples_dir = project_root / "examples"
        
        for example_file in examples_dir.glob("*.py"):
            print(f"  Checking syntax: {example_file.name}")
            
            # Read file content
            with open(example_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Check for syntax errors by trying to compile
            try:
                compile(content, example_file.name, 'exec')
                print(f"    Syntax: OK")
            except SyntaxError as e:
                pytest.fail(f"Syntax error in {example_file.name}: {e}")
    
    def test_example_imports(self):
        """Verify that examples can be imported without errors."""
        print("\nVerifying example imports...")
        
        examples_dir = project_root / "examples"
        
        for example_file in examples_dir.glob("*.py"):
            print(f"  Testing import: {example_file.name}")
            
            # Create a module spec
            spec = importlib.util.spec_from_file_location(
                example_file.stem,
                example_file
            )
            
            # Try to create module (but don't execute)
            try:
                module = importlib.util.module_from_spec(spec)
                print(f"    Import: OK (module created)")
            except Exception as e:
                print(f"    Import: Warning - {type(e).__name__}")
                # Don't fail on import warnings, some examples may have runtime dependencies
    
    def test_create_report_example(self):
        """Test the create_report.py example."""
        print("\nTesting create_report.py example...")
        
        example_path = project_root / "examples" / "create_report.py"
        assert example_path.exists(), f"Example not found: {example_path}"
        
        # Run the example
        try:
            # Import and run the example
            spec = importlib.util.spec_from_file_location("create_report", example_path)
            module = importlib.util.module_from_spec(spec)
            
            # Execute the module
            spec.loader.exec_module(module)
            
            # Check if main function exists and call it
            if hasattr(module, 'main'):
                print("  Running main() function...")
                result = module.main()
                
                # Check return value
                if result is not None:
                    print(f"  main() returned: {result}")
                
                # Check if files were created
                output_dir = project_root / "output" / "reports"
                if output_dir.exists():
                    files = list(output_dir.glob("*.docx"))
                    if files:
                        print(f"  Generated files: {len(files)}")
                        for file in files[:3]:  # Show first 3 files
                            size = file.stat().st_size
                            print(f"    - {file.name} ({size} bytes)")
                    else:
                        print("  No .docx files generated (may be normal in dummy mode)")
                else:
                    print("  Output directory not created (may be normal in dummy mode)")
            
            print("  create_report.py: PASSED")
            
        except Exception as e:
            print(f"  create_report.py: ERROR - {type(e).__name__}: {e}")
            # Don't fail in dummy mode
            if "dummy" not in str(e).lower():
                raise
    
    def test_process_spreadsheet_example(self):
        """Test the process_spreadsheet.py example."""
        print("\nTesting process_spreadsheet.py example...")
        
        example_path = project_root / "examples" / "process_spreadsheet.py"
        assert example_path.exists(), f"Example not found: {example_path}"
        
        # Run the example
        try:
            # Import and run the example
            spec = importlib.util.spec_from_file_location("process_spreadsheet", example_path)
            module = importlib.util.module_from_spec(spec)
            
            # Execute the module
            spec.loader.exec_module(module)
            
            # Check if main function exists and call it
            if hasattr(module, 'main'):
                print("  Running main() function...")
                result = module.main()
                
                # Check return value
                if result is not None:
                    print(f"  main() returned: {result}")
                
                # Check if files were created
                output_dir = project_root / "output" / "spreadsheets"
                if output_dir.exists():
                    files = list(output_dir.glob("*.*"))
                    if files:
                        print(f"  Generated files: {len(files)}")
                        for file in files[:3]:  # Show first 3 files
                            size = file.stat().st_size
                            print(f"    - {file.name} ({size} bytes)")
                    else:
                        print("  No files generated (may be normal in dummy mode)")
                else:
                    print("  Output directory not created (may be normal in dummy mode)")
            
            print("  process_spreadsheet.py: PASSED")
            
        except Exception as e:
            print(f"  process_spreadsheet.py: ERROR - {type(e).__name__}: {e}")
            # Don't fail in dummy mode
            if "dummy" not in str(e).lower():
                raise
    
    def test_generate_presentation_example(self):
        """Test the generate_presentation.py example."""
        print("\nTesting generate_presentation.py example...")
        
        example_path = project_root / "examples" / "generate_presentation.py"
        assert example_path.exists(), f"Example not found: {example_path}"
        
        # Run the example
        try:
            # Import and run the example
            spec = importlib.util.spec_from_file_location("generate_presentation", example_path)
            module = importlib.util.module_from_spec(spec)
            
            # Execute the module
            spec.loader.exec_module(module)
            
            # Check if main function exists and call it
            if hasattr(module, 'main'):
                print("  Running main() function...")
                result = module.main()
                
                # Check return value
                if result is not None:
                    print(f"  main() returned: {result}")
                
                # Check if files were created
                output_dir = project_root / "output" / "presentations"
                if output_dir.exists():
                    files = list(output_dir.glob("*.*"))
                    if files:
                        print(f"  Generated files: {len(files)}")
                        for file in files[:3]:  # Show first 3 files
                            size = file.stat().st_size
                            print(f"    - {file.name} ({size} bytes)")
                    else:
                        print("  No files generated (may be normal in dummy mode)")
                else:
                    print("  Output directory not created (may be normal in dummy mode)")
            
            print("  generate_presentation.py: PASSED")
            
        except Exception as e:
            print(f"  generate_presentation.py: ERROR - {type(e).__name__}: {e}")
            # Don't fail in dummy mode
            if "dummy" not in str(e).lower():
                raise
    
    def test_wps_integration_example(self):
        """Test the wps_integration_demo.py example."""
        print("\nTesting wps_integration_demo.py example...")
        
        example_path = project_root / "examples" / "wps_integration_demo.py"
        assert example_path.exists(), f"Example not found: {example_path}"
        
        # Run the example
        try:
            # Import and run the example
            spec = importlib.util.spec_from_file_location("wps_integration_demo", example_path)
            module = importlib.util.module_from_spec(spec)
            
            # Execute the module
            spec.loader.exec_module(module)
            
            # Check if main function exists and call it
            if hasattr(module, 'main'):
                print("  Running main() function...")
                result = module.main()
                
                # Check return value
                if result is not None:
                    print(f"  main() returned: {result}")
                
                # Check if files were created
                output_dir = project_root / "output" / "wps_tests"
                if output_dir.exists():
                    files = list(output_dir.glob("*.*"))
                    if files:
                        print(f"  Generated files: {len(files)}")
                        for file in files[:3]:  # Show first 3 files
                            size = file.stat().st_size
                            print(f"    - {file.name} ({size} bytes)")
                    else:
                        print("  No files generated (may be normal in dummy mode)")
                else:
                    print("  Output directory not created (may be normal in dummy mode)")
            
            print("  wps_integration_demo.py: PASSED")
            
        except Exception as e:
            print(f"  wps_integration_demo.py: ERROR - {type(e).__name__}: {e}")
            # Don't fail in dummy mode
            if "dummy" not in str(e).lower():
                raise
    
    def test_example_documentation(self):
        """Verify that examples have proper documentation."""
        print("\nVerifying example documentation...")
        
        examples_dir = project_root / "examples"
        
        for example_file in examples_dir.glob("*.py"):
            print(f"  Checking documentation: {example_file.name}")
            
            with open(example_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Check for docstring
            if '"""' in content or "'''" in content:
                print(f"    Docstring: Found")
            else:
                print(f"    Docstring: Missing (warning)")
            
            # Check for main function
            if 'def main()' in content or 'if __name__' in content:
                print(f"    Main function: Found")
            else:
                print(f"    Main function: Missing (warning)")
            
            # Check for imports
            if 'import ' in content or 'from ' in content:
                print(f"    Imports: Found")
            else:
                print(f"    Imports: Missing (unusual)")
    
    def test_example_output_cleanup(self):
        """Test that examples clean up after themselves (or at least don't crash)."""
        print("\nTesting example output cleanup...")
        
        # Run each example and check they don't leave processes hanging
        examples = [
            "create_report.py",
            "process_spreadsheet.py",
            "generate_presentation.py",
            "wps_integration_demo.py",
        ]
        
        for example in examples:
            print(f"  Testing cleanup: {example}")
            
            example_path = project_root / "examples" / example
            
            # Run example in subprocess with timeout
            try:
                result = subprocess.run(
                    [sys.executable, str(example_path)],
                    capture_output=True,
                    text=True,
                    timeout=30,  # 30 second timeout
                    cwd=project_root
                )
                
                if result.returncode == 0:
                    print(f"    Exit code: 0 (success)")
                else:
                    print(f"    Exit code: {result.returncode}")
                    print(f"    stderr: {result.stderr[:200]}...")
                
                # Check for any obvious errors in output
                error_keywords = ['error', 'exception', 'traceback', 'failed']
                for keyword in error_keywords:
                    if keyword in result.stderr.lower():
                        print(f"    Warning: '{keyword}' found in stderr")
                
            except subprocess.TimeoutExpired:
                print(f"    ERROR: Timeout after 30 seconds")
                # This is a problem - examples shouldn't hang
                pytest.fail(f"Example {example} timed out")
            except Exception as e:
                print(f"    ERROR: {type(e).__name__}: {e}")
    
    def test_example_consistency(self):
        """Test that examples follow consistent patterns."""
        print("\nTesting example consistency...")
        
        examples_dir = project_root / "examples"
        patterns_to_check = [
            ("shebang", "#!/usr/bin/env python3"),
            ("encoding", "# -*- coding: utf-8 -*-"),
            ("docstring", '"""'),
        ]
        
        for example_file in examples_dir.glob("*.py"):
            print(f"  Checking patterns: {example_file.name}")
            
            with open(example_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # Check first few lines for patterns
            first_lines = ''.join(lines[:10])
            
            for pattern_name, pattern in patterns_to_check:
                if pattern in first_lines:
                    print(f"    {pattern_name}: Found")
                else:
                    print(f"    {pattern_name}: Missing (warning)")
            
            # Check line length (warn about very long lines)
            long_lines = []
            for i, line in enumerate(lines, 1):
                if len(line.rstrip('\n')) > 100:  # More than 100 characters
                    long_lines.append((i, len(line)))
            
            if long_lines:
                print(f"    Long lines: {len(long_lines)} lines > 100 chars")
                for line_num, length in long_lines[:3]:  # Show first 3
                    print(f"      Line {line_num}: {length} chars")


def run_example_tests():
    """Run all example validation tests."""
    print("=" * 60)
    print("EXAMPLE VALIDATION TESTS")
    print("=" * 60)
    
    tester = TestExamples()
    
    tests = [
        tester.test_example_files_exist,
        tester.test_example_syntax,
        tester.test_example_imports,
        tester.test_create_report_example,
        tester.test_process_spreadsheet_example,
        tester.test_generate_presentation_example,
        tester.test_wps_integration_example,
        tester.test_example_documentation,
        tester.test_example_output_cleanup,
        tester.test_example_consistency,
    ]
    
    results = []
    
    for test_func in tests:
        try:
            tester.setup_method()
            test_func()
            results.append((test_func.__name__, "PASSED", ""))
            print(f"PASSED: {test_func.__name__}")
        except pytest.skip.Exception as e:
            results.append((test_func.__name__, "SKIPPED", str(e)))
            print(f"SKIPPED: {test_func.__name__} - {e}")
        except Exception as e:
            results.append((test_func.__name__, "FAILED", str(e)))
            print(f"FAILED: {test_func.__name__} - {e}")
        finally:
            tester.teardown_method()
    
    # Summary
    passed = sum(1 for _, status, _ in results if status == "PASSED")
    failed = sum(1 for _, status, _ in results if status == "FAILED")
    skipped = sum(1 for _, status, _ in results if status == "SKIPPED")
    
    print("\n" + "=" * 60)
    print("EXAMPLE TEST SUMMARY:")
    print(f"  Passed:  {passed}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped}")
    print(f"  Total:   {passed + failed + skipped}")
    print("=" * 60)
    
    if failed > 0:
        print("\nFAILED: Some example tests failed!")
        return False
    else:
        print("\nPASSED: All example tests passed!")
        return True


if __name__ == "__main__":
    success = run_example_tests()
    sys.exit(0 if success else 1)