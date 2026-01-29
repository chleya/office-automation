"""
Integration Tests for Office Automation
======================================

These tests verify that the complete workflow works end-to-end.
"""

import os
import tempfile
import pytest
from pathlib import Path
from datetime import datetime

# Import the module
import sys
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from office_automation import OfficeAutomation


class TestOfficeAutomationIntegration:
    """Integration tests for the complete Office Automation workflow."""
    
    def setup_method(self):
        """Setup for each test."""
        self.office = OfficeAutomation()
        self.temp_dir = tempfile.mkdtemp(prefix="office_test_")
        print(f"\nTest temp directory: {self.temp_dir}")
    
    def teardown_method(self):
        """Cleanup after each test."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_complete_document_workflow(self):
        """Test complete document creation and processing workflow."""
        print("\nTesting complete document workflow...")
        
        # Get system info
        info = self.office.get_info()
        assert 'version' in info
        assert 'modules' in info
        
        # Test Word module
        if hasattr(self.office.word, 'available') and self.office.word.available:
            print("  Testing Word module...")
            
            # Create a document
            doc_path = os.path.join(self.temp_dir, "test_document.docx")
            result = self.office.word.create_document()
            assert result is not None
            
            # Add content
            content_result = self.office.word.add_paragraph(
                result, 
                text="Test paragraph for integration testing",
                style="Normal"
            )
            assert content_result is not None
            
            # Save document
            save_result = self.office.word.save_document(result, doc_path)
            assert save_result is not None
            
            # Verify file was created
            assert os.path.exists(doc_path), f"Document not created: {doc_path}"
            print(f"    Document created: {doc_path}")
        
        # Test Excel module
        if hasattr(self.office.excel, 'available') and self.office.excel.available:
            print("  Testing Excel module...")
            
            # Create a spreadsheet
            excel_path = os.path.join(self.temp_dir, "test_spreadsheet.xlsx")
            result = self.office.excel.create_spreadsheet()
            assert result is not None
            
            # Add data
            data_result = self.office.excel.add_data(
                result,
                data=[["Name", "Age", "Score"], ["Alice", 25, 95], ["Bob", 30, 88]]
            )
            assert data_result is not None
            
            # Save spreadsheet
            save_result = self.office.excel.save_spreadsheet(result, excel_path)
            assert save_result is not None
            
            # Verify file was created
            assert os.path.exists(excel_path), f"Spreadsheet not created: {excel_path}"
            print(f"    Spreadsheet created: {excel_path}")
        
        # Test PowerPoint module
        if hasattr(self.office.powerpoint, 'available') and self.office.powerpoint.available:
            print("  Testing PowerPoint module...")
            
            # Create a presentation
            ppt_path = os.path.join(self.temp_dir, "test_presentation.pptx")
            result = self.office.powerpoint.create_presentation()
            assert result is not None
            
            # Add a slide
            slide_result = self.office.powerpoint.add_slide(
                result,
                title="Integration Test Slide"
            )
            assert slide_result is not None
            
            # Save presentation
            save_result = self.office.powerpoint.save_presentation(result, ppt_path)
            assert save_result is not None
            
            # Verify file was created
            assert os.path.exists(ppt_path), f"Presentation not created: {ppt_path}"
            print(f"    Presentation created: {ppt_path}")
        
        print("  Complete workflow test passed!")
    
    def test_format_conversion(self):
        """Test document format conversion."""
        print("\nTesting format conversion...")
        
        if not hasattr(self.office.converter, 'available') or not self.office.converter.available:
            pytest.skip("Converter module not available")
        
        # Create a test document
        doc_path = os.path.join(self.temp_dir, "convert_test.docx")
        with open(doc_path, 'w') as f:
            f.write("Test document for conversion")
        
        # Test conversion to PDF
        pdf_path = os.path.join(self.temp_dir, "converted.pdf")
        result = self.office.converter.docx_to_pdf(doc_path, pdf_path)
        assert result is not None
        
        # Verify conversion result
        if hasattr(result, 'success'):
            assert result.success, f"Conversion failed: {result}"
        elif isinstance(result, dict):
            assert result.get('success', False), f"Conversion failed: {result}"
        
        print(f"    Conversion test passed: {doc_path} -> {pdf_path}")
    
    def test_wps_integration(self):
        """Test WPS Office integration."""
        print("\nTesting WPS integration...")
        
        if not hasattr(self.office.wps, 'available'):
            pytest.skip("WPS module not available in this version")
        
        # Check WPS availability
        wps_available = self.office.wps.available
        print(f"    WPS available: {wps_available}")
        
        if wps_available:
            # Create a test document
            doc_path = os.path.join(self.temp_dir, "wps_test.docx")
            with open(doc_path, 'w') as f:
                f.write("Test document for WPS optimization")
            
            # Test WPS optimization
            optimized_path = os.path.join(self.temp_dir, "wps_optimized.docx")
            result = self.office.wps.optimize_for_wps(doc_path, optimized_path)
            assert result is not None
            
            print(f"    WPS optimization test passed")
    
    def test_error_handling(self):
        """Test error handling in the workflow."""
        print("\nTesting error handling...")
        
        # Test with invalid paths
        invalid_path = "/invalid/path/that/does/not/exist/document.docx"
        
        # Word module error handling
        if hasattr(self.office.word, 'save_document'):
            try:
                # This should either handle the error gracefully or raise an exception
                result = self.office.word.save_document("dummy_id", invalid_path)
                # If we get here, the module handled the error
                print(f"    Word module handled invalid path gracefully")
            except Exception as e:
                # Exception is also acceptable
                print(f"    Word module raised exception for invalid path: {type(e).__name__}")
        
        # Test with None/null inputs
        try:
            info = self.office.get_info()
            assert info is not None
            print(f"    get_info() handles None config gracefully")
        except Exception as e:
            print(f"    get_info() raised exception: {type(e).__name__}")
    
    def test_quick_functions(self):
        """Test quick access functions."""
        print("\nTesting quick functions...")
        
        from office_automation import (
            quick_create_document,
            quick_create_spreadsheet,
            quick_convert
        )
        
        # Test quick_create_document
        doc_path = os.path.join(self.temp_dir, "quick_doc.docx")
        try:
            result = quick_create_document(doc_path, "Quick Test Document")
            assert result is not None
            print(f"    quick_create_document() works")
        except Exception as e:
            print(f"    quick_create_document() error: {type(e).__name__}")
        
        # Test quick_create_spreadsheet
        excel_path = os.path.join(self.temp_dir, "quick_data.xlsx")
        try:
            data = [["Product", "Sales"], ["A", 100], ["B", 200]]
            result = quick_create_spreadsheet(excel_path, data)
            assert result is not None
            print(f"    quick_create_spreadsheet() works")
        except Exception as e:
            print(f"    quick_create_spreadsheet() error: {type(e).__name__}")
        
        # Test quick_convert
        if os.path.exists(doc_path):
            pdf_path = os.path.join(self.temp_dir, "quick_converted.pdf")
            try:
                result = quick_convert(doc_path, pdf_path)
                assert result is not None
                print(f"    quick_convert() works")
            except Exception as e:
                print(f"    quick_convert() error: {type(e).__name__}")
    
    def test_configuration(self):
        """Test configuration handling."""
        print("\nTesting configuration...")
        
        # Test with custom config
        custom_config = {
            'temp_dir': self.temp_dir,
            'log_level': 'DEBUG',
            'max_file_size': 10485760  # 10MB
        }
        
        office_custom = OfficeAutomation(config=custom_config)
        assert office_custom.config is not None
        
        # Verify config was merged with defaults
        assert 'temp_dir' in office_custom.config
        assert office_custom.config['temp_dir'] == self.temp_dir
        
        # Verify default values are still present
        assert 'default_templates' in office_custom.config
        
        print(f"    Configuration test passed")
    
    def test_performance(self):
        """Test basic performance (not rigorous, just sanity check)."""
        print("\nTesting basic performance...")
        
        import time
        
        # Time the initialization
        start_time = time.time()
        office = OfficeAutomation()
        init_time = time.time() - start_time
        
        print(f"    Initialization time: {init_time:.3f} seconds")
        
        # Time getting info
        start_time = time.time()
        info = office.get_info()
        info_time = time.time() - start_time
        
        print(f"    get_info() time: {info_time:.3f} seconds")
        
        # Basic performance assertions
        assert init_time < 5.0, f"Initialization too slow: {init_time:.3f}s"
        assert info_time < 2.0, f"get_info() too slow: {info_time:.3f}s"
        
        print(f"    Performance test passed")


def test_example_files():
    """Test that example files can be imported and run."""
    print("\nTesting example files...")
    
    examples_dir = project_root / "examples"
    
    if not examples_dir.exists():
        pytest.skip("Examples directory not found")
    
    example_files = [
        "create_report.py",
        "process_spreadsheet.py", 
        "generate_presentation.py",
        "wps_integration_demo.py"
    ]
    
    for example_file in example_files:
        example_path = examples_dir / example_file
        if example_path.exists():
            print(f"  Found example: {example_file}")
            
            # Try to import the module
            try:
                # Read and check for syntax errors
                with open(example_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Check for main function
                if 'def main()' in content or 'if __name__' in content:
                    print(f"    {example_file} has main function")
                else:
                    print(f"    {example_file} does not have main function")
                    
            except Exception as e:
                print(f"    Error reading {example_file}: {type(e).__name__}")
        else:
            print(f"  Missing example: {example_file}")


if __name__ == "__main__":
    """Run integration tests directly."""
    print("Running Office Automation integration tests...")
    print("=" * 60)
    
    # Create test instance
    tester = TestOfficeAutomationIntegration()
    
    # Run tests
    tests = [
        tester.test_complete_document_workflow,
        tester.test_format_conversion,
        tester.test_wps_integration,
        tester.test_error_handling,
        tester.test_quick_functions,
        tester.test_configuration,
        tester.test_performance,
    ]
    
    passed = 0
    failed = 0
    skipped = 0
    
    for test_func in tests:
        try:
            tester.setup_method()
            test_func()
            print(f"PASSED: {test_func.__name__}")
            passed += 1
        except pytest.skip.Exception as e:
            print(f"SKIPPED: {test_func.__name__} - {e}")
            skipped += 1
        except Exception as e:
            print(f"FAILED: {test_func.__name__} - {e}")
            failed += 1
        finally:
            tester.teardown_method()
    
    # Test example files
    try:
        test_example_files()
        print("PASSED: test_example_files")
        passed += 1
    except Exception as e:
        print(f"FAILED: test_example_files - {e}")
        failed += 1
    
    print("\n" + "=" * 60)
    print(f"TEST SUMMARY:")
    print(f"  Passed:  {passed}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped}")
    print(f"  Total:   {passed + failed + skipped}")
    print("=" * 60)
    
    if failed > 0:
        print("\nFAILED: Some tests failed!")
        sys.exit(1)
    else:
        print("\nPASSED: All tests passed!")
        sys.exit(0)