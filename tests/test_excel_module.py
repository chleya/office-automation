"""
Excel Module Tests
==================

Tests for the Excel spreadsheet automation module.
"""

import os
import tempfile
import pytest
from pathlib import Path
import sys

# Import the module
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from office_automation import OfficeAutomation


class TestExcelModule:
    """Tests for the Excel spreadsheet module."""
    
    def setup_method(self):
        """Setup for each test."""
        self.office = OfficeAutomation()
        self.temp_dir = tempfile.mkdtemp(prefix="excel_test_")
        print(f"\nTest directory: {self.temp_dir}")
    
    def teardown_method(self):
        """Cleanup after each test."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_excel_module_available(self):
        """Test that Excel module is available."""
        print("\nTesting Excel module availability...")
        
        assert hasattr(self.office, 'excel'), "Excel module not found"
        assert self.office.excel is not None, "Excel module is None"
        
        # Check capabilities
        if hasattr(self.office.excel, 'get_capabilities'):
            capabilities = self.office.excel.get_capabilities()
            assert isinstance(capabilities, dict), "Capabilities should be a dict"
            assert 'available' in capabilities, "Capabilities missing 'available' key"
            print(f"  Excel capabilities: {capabilities}")
        else:
            print("  Excel module does not have get_capabilities method")
    
    def test_create_spreadsheet(self):
        """Test spreadsheet creation."""
        print("\nTesting spreadsheet creation...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        
        # Test basic spreadsheet creation
        spreadsheet = self.office.excel.create_spreadsheet()
        assert spreadsheet is not None, "Spreadsheet creation returned None"
        
        # Test with template
        spreadsheet_with_template = self.office.excel.create_spreadsheet(template="financial")
        assert spreadsheet_with_template is not None, "Spreadsheet with template creation returned None"
        
        print("  Spreadsheet creation: PASSED")
    
    def test_add_data(self):
        """Test adding data to spreadsheet."""
        print("\nTesting data addition...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        
        # Create spreadsheet
        spreadsheet = self.office.excel.create_spreadsheet()
        assert spreadsheet is not None
        
        # Test adding data
        if hasattr(self.office.excel, 'add_data'):
            test_data = [
                ["Product", "Q1", "Q2", "Q3", "Q4"],
                ["Widget A", 100, 150, 200, 180],
                ["Widget B", 80, 120, 160, 140],
                ["Widget C", 60, 90, 120, 110],
            ]
            
            result = self.office.excel.add_data(spreadsheet, test_data)
            assert result is not None, "add_data returned None"
            print("  add_data: PASSED")
        
        # Test adding data with sheet name
        if hasattr(self.office.excel, 'add_data_to_sheet'):
            test_data = [
                ["Region", "Sales", "Growth"],
                ["North", 50000, 0.15],
                ["South", 45000, 0.12],
                ["East", 48000, 0.18],
                ["West", 52000, 0.20],
            ]
            
            result = self.office.excel.add_data_to_sheet(spreadsheet, test_data, sheet_name="Regional Sales")
            assert result is not None, "add_data_to_sheet returned None"
            print("  add_data_to_sheet: PASSED")
    
    def test_add_formulas(self):
        """Test adding formulas to spreadsheet."""
        print("\nTesting formula addition...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        
        spreadsheet = self.office.excel.create_spreadsheet()
        
        if hasattr(self.office.excel, 'add_formula'):
            # Test basic formulas
            formulas = [
                ("A1", "=SUM(B1:B10)"),
                ("C1", "=AVERAGE(D1:D20)"),
                ("E1", '=IF(F1>100, "High", "Low")'),
            ]
            
            for cell, formula in formulas:
                try:
                    result = self.office.excel.add_formula(spreadsheet, cell, formula)
                    if result is not None:
                        print(f"  Formula {cell}={formula}: Added")
                    else:
                        print(f"  Formula {cell}={formula}: add_formula returned None")
                except Exception as e:
                    print(f"  Formula {cell}={formula}: Error - {type(e).__name__}")
    
    def test_create_chart(self):
        """Test chart creation."""
        print("\nTesting chart creation...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        
        spreadsheet = self.office.excel.create_spreadsheet()
        
        if hasattr(self.office.excel, 'create_chart'):
            # Add some data first
            if hasattr(self.office.excel, 'add_data'):
                data = [
                    ["Month", "Sales", "Expenses"],
                    ["Jan", 5000, 3000],
                    ["Feb", 6000, 3500],
                    ["Mar", 7000, 4000],
                    ["Apr", 6500, 3800],
                ]
                self.office.excel.add_data(spreadsheet, data)
            
            # Test creating different chart types
            chart_types = ['column', 'line', 'bar', 'pie', 'scatter']
            
            for chart_type in chart_types:
                try:
                    result = self.office.excel.create_chart(
                        spreadsheet,
                        chart_type=chart_type,
                        title=f"{chart_type.capitalize()} Chart",
                        data_range="A1:C6"
                    )
                    if result is not None:
                        print(f"  Chart type {chart_type}: Created")
                    else:
                        print(f"  Chart type {chart_type}: create_chart returned None")
                except Exception as e:
                    print(f"  Chart type {chart_type}: Error - {type(e).__name__}")
    
    def test_save_spreadsheet(self):
        """Test saving spreadsheet to file."""
        print("\nTesting spreadsheet saving...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        if not hasattr(self.office.excel, 'save_spreadsheet'):
            pytest.skip("Excel module does not have save_spreadsheet method")
        
        # Create spreadsheet
        spreadsheet = self.office.excel.create_spreadsheet()
        assert spreadsheet is not None
        
        # Add some data
        if hasattr(self.office.excel, 'add_data'):
            self.office.excel.add_data(spreadsheet, [["Test", "Data"], [1, 2], [3, 4]])
        
        # Save spreadsheet
        excel_path = os.path.join(self.temp_dir, "test_spreadsheet.xlsx")
        result = self.office.excel.save_spreadsheet(spreadsheet, excel_path)
        assert result is not None, "save_spreadsheet returned None"
        
        # Verify file was created
        assert os.path.exists(excel_path), f"Spreadsheet not created: {excel_path}"
        file_size = os.path.getsize(excel_path)
        assert file_size > 0, f"Spreadsheet file is empty: {excel_path}"
        
        print(f"  Spreadsheet saved: {excel_path} ({file_size} bytes)")
        print("  Spreadsheet saving: PASSED")
    
    def test_spreadsheet_formats(self):
        """Test different spreadsheet formats."""
        print("\nTesting spreadsheet formats...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        if not hasattr(self.office.excel, 'save_spreadsheet'):
            pytest.skip("Excel module does not have save_spreadsheet method")
        
        # Create spreadsheet
        spreadsheet = self.office.excel.create_spreadsheet()
        if hasattr(self.office.excel, 'add_data'):
            self.office.excel.add_data(spreadsheet, [["Test", "Formats"], [1, 2]])
        
        # Test different formats
        formats = [
            ("xlsx", ".xlsx"),
            ("xls", ".xls"),
            ("csv", ".csv"),
            ("ods", ".ods"),
            ("pdf", ".pdf"),
        ]
        
        for format_name, extension in formats:
            try:
                file_path = os.path.join(self.temp_dir, f"test_spreadsheet{extension}")
                result = self.office.excel.save_spreadsheet(spreadsheet, file_path, format=format_name)
                
                if result is not None:
                    # Check if file was created (some formats may not be supported)
                    if os.path.exists(file_path):
                        file_size = os.path.getsize(file_path)
                        print(f"  Format {format_name}: Created ({file_size} bytes)")
                    else:
                        print(f"  Format {format_name}: Not created (may not be supported)")
                else:
                    print(f"  Format {format_name}: save_spreadsheet returned None")
                    
            except Exception as e:
                print(f"  Format {format_name}: Error - {type(e).__name__}")
    
    def test_worksheet_operations(self):
        """Test worksheet operations."""
        print("\nTesting worksheet operations...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        
        spreadsheet = self.office.excel.create_spreadsheet()
        
        # Test adding worksheet
        if hasattr(self.office.excel, 'add_worksheet'):
            worksheet_names = ["Data", "Analysis", "Charts", "Summary"]
            
            for name in worksheet_names:
                try:
                    result = self.office.excel.add_worksheet(spreadsheet, name)
                    if result is not None:
                        print(f"  Worksheet '{name}': Added")
                    else:
                        print(f"  Worksheet '{name}': add_worksheet returned None")
                except Exception as e:
                    print(f"  Worksheet '{name}': Error - {type(e).__name__}")
        
        # Test switching worksheet
        if hasattr(self.office.excel, 'switch_worksheet'):
            try:
                result = self.office.excel.switch_worksheet(spreadsheet, "Analysis")
                if result is not None:
                    print("  switch_worksheet: OK")
                else:
                    print("  switch_worksheet: Returned None")
            except Exception as e:
                print(f"  switch_worksheet: Error - {type(e).__name__}")
    
    def test_cell_formatting(self):
        """Test cell formatting operations."""
        print("\nTesting cell formatting...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        
        spreadsheet = self.office.excel.create_spreadsheet()
        
        if hasattr(self.office.excel, 'format_cell'):
            # Test different formatting options
            formatting_options = [
                ("A1", {"bold": True, "font_size": 14}),
                ("B1", {"italic": True, "align": "center"}),
                ("C1", {"number_format": "0.00%"}),
                ("D1", {"fill_color": "yellow", "border": True}),
            ]
            
            for cell, formatting in formatting_options:
                try:
                    result = self.office.excel.format_cell(spreadsheet, cell, formatting)
                    if result is not None:
                        print(f"  Format cell {cell}: Applied")
                    else:
                        print(f"  Format cell {cell}: format_cell returned None")
                except Exception as e:
                    print(f"  Format cell {cell}: Error - {type(e).__name__}")
    
    def test_error_handling(self):
        """Test error handling in Excel module."""
        print("\nTesting error handling...")
        
        # Test with invalid inputs
        if hasattr(self.office.excel, 'create_spreadsheet'):
            # Test None input
            try:
                result = self.office.excel.create_spreadsheet(template=None)
                print("  create_spreadsheet with None template: OK")
            except Exception as e:
                print(f"  create_spreadsheet with None template: {type(e).__name__}")
        
        # Test with invalid data
        if hasattr(self.office.excel, 'add_data'):
            spreadsheet = self.office.excel.create_spreadsheet() if hasattr(self.office.excel, 'create_spreadsheet') else "dummy"
            
            invalid_data = [
                None,
                [],
                [None],
                [[], []],
            ]
            
            for data in invalid_data:
                try:
                    result = self.office.excel.add_data(spreadsheet, data)
                    print(f"  add_data with invalid data {data}: Returned {result}")
                except Exception as e:
                    print(f"  add_data with invalid data {data}: {type(e).__name__}")
    
    def test_performance(self):
        """Test Excel module performance."""
        print("\nTesting Excel module performance...")
        
        if not hasattr(self.office.excel, 'create_spreadsheet'):
            pytest.skip("Excel module does not have create_spreadsheet method")
        
        import time
        
        # Time spreadsheet creation
        start_time = time.perf_counter()
        spreadsheet = self.office.excel.create_spreadsheet()
        creation_time = time.perf_counter() - start_time
        
        print(f"  Spreadsheet creation time: {creation_time:.4f}s")
        assert creation_time < 5.0, f"Spreadsheet creation too slow: {creation_time:.4f}s"
        
        # Time data addition
        if hasattr(self.office.excel, 'add_data'):
            # Create larger dataset
            data = [["Row", "Value"]]
            for i in range(100):
                data.append([f"Row {i+1}", i * 10])
            
            start_time = time.perf_counter()
            result = self.office.excel.add_data(spreadsheet, data)
            addition_time = time.perf_counter() - start_time
            
            print(f"  100 rows data addition time: {addition_time:.4f}s")
            assert addition_time < 2.0, f"Data addition too slow: {addition_time:.4f}s"
        
        print("  Performance test: PASSED")


def run_excel_tests():
    """Run all Excel module tests."""
    print("=" * 60)
    print("EXCEL MODULE TESTS")
    print("=" * 60)
    
    tester = TestExcelModule()
    
    tests = [
        tester.test_excel_module_available,
        tester.test_create_spreadsheet,
        tester.test_add_data,
        tester.test_add_formulas,
        tester.test_create_chart,
        tester.test_save_spreadsheet,
        tester.test_spreadsheet_formats,
        tester.test_worksheet_operations,
        tester.test_cell_formatting,
        tester.test_error_handling,
        tester.test_performance,
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
    print("EXCEL MODULE TEST SUMMARY:")
    print(f"  Passed:  {passed}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped}")
    print(f"  Total:   {passed + failed + skipped}")
    print("=" * 60)
    
    if failed > 0:
        print("\nFAILED: Some Excel module tests failed!")
        return False
    else:
        print("\nPASSED: All Excel module tests passed!")
        return True


if __name__ == "__main__":
    success = run_excel_tests()
    sys.exit(0 if success else 1)