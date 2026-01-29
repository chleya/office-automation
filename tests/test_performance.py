"""
Performance Tests for Office Automation
======================================

These tests measure performance characteristics of the library.
"""

import os
import time
import tempfile
import pytest
from pathlib import Path
import sys

# Import the module
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from office_automation import OfficeAutomation


class TestOfficeAutomationPerformance:
    """Performance tests for Office Automation."""
    
    def setup_method(self):
        """Setup for each test."""
        self.office = OfficeAutomation()
        self.temp_dir = tempfile.mkdtemp(prefix="perf_test_")
    
    def teardown_method(self):
        """Cleanup after each test."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_initialization_performance(self):
        """Test that initialization is fast."""
        print("\nTesting initialization performance...")
        
        # Measure initialization time
        times = []
        for i in range(5):  # Run multiple times for average
            start_time = time.perf_counter()
            office = OfficeAutomation()
            end_time = time.perf_counter()
            times.append(end_time - start_time)
        
        avg_time = sum(times) / len(times)
        max_time = max(times)
        
        print(f"  Initialization times: {times}")
        print(f"  Average: {avg_time:.4f}s, Max: {max_time:.4f}s")
        
        # Performance requirements
        assert avg_time < 1.0, f"Initialization too slow: {avg_time:.4f}s average"
        assert max_time < 2.0, f"Initialization too slow: {max_time:.4f}s max"
        
        print("  Initialization performance: PASSED")
    
    def test_get_info_performance(self):
        """Test that get_info() is fast."""
        print("\nTesting get_info() performance...")
        
        # Measure get_info time
        times = []
        for i in range(10):  # Run multiple times
            start_time = time.perf_counter()
            info = self.office.get_info()
            end_time = time.perf_counter()
            times.append(end_time - start_time)
            
            # Verify we got valid info
            assert 'version' in info
            assert 'modules' in info
        
        avg_time = sum(times) / len(times)
        max_time = max(times)
        
        print(f"  get_info() times: {times}")
        print(f"  Average: {avg_time:.4f}s, Max: {max_time:.4f}s")
        
        # Performance requirements
        assert avg_time < 0.5, f"get_info() too slow: {avg_time:.4f}s average"
        assert max_time < 1.0, f"get_info() too slow: {max_time:.4f}s max"
        
        print("  get_info() performance: PASSED")
    
    def test_document_creation_performance(self):
        """Test document creation performance."""
        print("\nTesting document creation performance...")
        
        if not hasattr(self.office.word, 'available') or not self.office.word.available:
            pytest.skip("Word module not available")
        
        # Create multiple documents
        doc_count = 5
        creation_times = []
        
        for i in range(doc_count):
            doc_path = os.path.join(self.temp_dir, f"perf_doc_{i}.docx")
            
            start_time = time.perf_counter()
            
            # Create document
            doc = self.office.word.create_document()
            
            # Add some content
            for j in range(10):  # Add 10 paragraphs
                self.office.word.add_paragraph(
                    doc,
                    text=f"Performance test paragraph {j} for document {i}",
                    style="Normal"
                )
            
            # Save document
            self.office.word.save_document(doc, doc_path)
            
            end_time = time.perf_counter()
            creation_times.append(end_time - start_time)
            
            # Verify file was created
            assert os.path.exists(doc_path), f"Document not created: {doc_path}"
        
        avg_time = sum(creation_times) / len(creation_times)
        total_time = sum(creation_times)
        
        print(f"  Created {doc_count} documents with 10 paragraphs each")
        print(f"  Creation times: {creation_times}")
        print(f"  Average per document: {avg_time:.4f}s")
        print(f"  Total time: {total_time:.4f}s")
        
        # Performance requirements
        assert avg_time < 5.0, f"Document creation too slow: {avg_time:.4f}s average"
        
        print("  Document creation performance: PASSED")
    
    def test_memory_usage(self):
        """Test memory usage (basic check)."""
        print("\nTesting memory usage...")
        
        import psutil
        import os
        
        # Get initial memory usage
        process = psutil.Process(os.getpid())
        initial_memory = process.memory_info().rss / 1024 / 1024  # MB
        
        print(f"  Initial memory: {initial_memory:.2f} MB")
        
        # Create multiple OfficeAutomation instances
        instances = []
        for i in range(10):
            office = OfficeAutomation()
            info = office.get_info()
            instances.append(office)
        
        # Get memory after creating instances
        current_memory = process.memory_info().rss / 1024 / 1024  # MB
        memory_increase = current_memory - initial_memory
        
        print(f"  Memory after 10 instances: {current_memory:.2f} MB")
        print(f"  Memory increase: {memory_increase:.2f} MB")
        
        # Memory requirements
        assert memory_increase < 50.0, f"Memory usage too high: {memory_increase:.2f} MB increase"
        
        print("  Memory usage: PASSED")
    
    def test_concurrent_operations(self):
        """Test performance with concurrent-like operations."""
        print("\nTesting concurrent operations...")
        
        # Simulate concurrent operations by running them sequentially
        # but measuring total time
        
        operations = [
            ("get_info", lambda: self.office.get_info()),
            ("create_doc", lambda: self.office.word.create_document() if hasattr(self.office.word, 'create_document') else None),
            ("create_excel", lambda: self.office.excel.create_spreadsheet() if hasattr(self.office.excel, 'create_spreadsheet') else None),
            ("create_ppt", lambda: self.office.powerpoint.create_presentation() if hasattr(self.office.powerpoint, 'create_presentation') else None),
        ]
        
        operation_times = {}
        
        start_total = time.perf_counter()
        
        for op_name, op_func in operations:
            op_start = time.perf_counter()
            
            # Run operation multiple times
            for i in range(3):
                try:
                    result = op_func()
                    if result is not None:
                        pass  # Operation succeeded
                except:
                    pass  # Operation not available or failed
            
            op_end = time.perf_counter()
            operation_times[op_name] = op_end - op_start
        
        end_total = time.perf_counter()
        total_time = end_total - start_total
        
        print(f"  Operation times:")
        for op_name, op_time in operation_times.items():
            print(f"    {op_name}: {op_time:.4f}s")
        
        print(f"  Total time for all operations: {total_time:.4f}s")
        
        # Performance requirements
        assert total_time < 10.0, f"Concurrent operations too slow: {total_time:.4f}s total"
        
        print("  Concurrent operations performance: PASSED")
    
    def test_file_io_performance(self):
        """Test file I/O performance."""
        print("\nTesting file I/O performance...")
        
        # Create test files of different sizes
        file_sizes = [1024, 10240, 102400]  # 1KB, 10KB, 100KB
        io_times = []
        
        for size in file_sizes:
            file_path = os.path.join(self.temp_dir, f"test_{size}.txt")
            
            # Create file with dummy content
            start_time = time.perf_counter()
            
            with open(file_path, 'w') as f:
                f.write('X' * size)
            
            # Read file back
            with open(file_path, 'r') as f:
                content = f.read()
            
            end_time = time.perf_counter()
            io_time = end_time - start_time
            
            io_times.append((size, io_time))
            
            # Verify
            assert len(content) == size, f"File size mismatch: {len(content)} != {size}"
        
        print(f"  File I/O performance:")
        for size, io_time in io_times:
            speed = size / io_time / 1024  # KB/s
            print(f"    {size} bytes: {io_time:.4f}s ({speed:.2f} KB/s)")
        
        # Performance requirements
        for size, io_time in io_times:
            assert io_time < 1.0, f"File I/O too slow for {size} bytes: {io_time:.4f}s"
        
        print("  File I/O performance: PASSED")


def run_performance_tests():
    """Run all performance tests and generate report."""
    print("=" * 60)
    print("OFFICE AUTOMATION PERFORMANCE TESTS")
    print("=" * 60)
    
    tester = TestOfficeAutomationPerformance()
    
    tests = [
        tester.test_initialization_performance,
        tester.test_get_info_performance,
        tester.test_document_creation_performance,
        tester.test_memory_usage,
        tester.test_concurrent_operations,
        tester.test_file_io_performance,
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
    
    # Generate performance report
    report_dir = Path(tester.temp_dir)
    report_dir.mkdir(parents=True, exist_ok=True)
    report_path = report_dir / "performance_report.md"
    with open(report_path, 'w') as f:
        f.write("# Office Automation Performance Report\n\n")
        f.write(f"Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write("## Test Results\n\n")
        f.write("| Test | Status | Notes |\n")
        f.write("|------|--------|-------|\n")
        
        for test_name, status, notes in results:
            f.write(f"| {test_name} | {status} | {notes} |\n")
        
        f.write("\n## System Information\n")
        f.write(f"- Python: {sys.version}\n")
        f.write(f"- Platform: {sys.platform}\n")
        f.write(f"- CPU Cores: {os.cpu_count()}\n")
        
        try:
            import psutil
            memory = psutil.virtual_memory()
            f.write(f"- Total Memory: {memory.total / 1024 / 1024 / 1024:.2f} GB\n")
            f.write(f"- Available Memory: {memory.available / 1024 / 1024 / 1024:.2f} GB\n")
        except:
            f.write("- Memory info: Not available\n")
    
    print(f"\nPerformance report saved to: {report_path}")
    
    # Summary
    passed = sum(1 for _, status, _ in results if status == "PASSED")
    failed = sum(1 for _, status, _ in results if status == "FAILED")
    skipped = sum(1 for _, status, _ in results if status == "SKIPPED")
    
    print("\n" + "=" * 60)
    print("PERFORMANCE TEST SUMMARY:")
    print(f"  Passed:  {passed}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped}")
    print(f"  Total:   {passed + failed + skipped}")
    print("=" * 60)
    
    if failed > 0:
        print("\nFAILED: Some performance tests failed!")
        return False
    else:
        print("\nPASSED: All performance tests passed!")
        return True


if __name__ == "__main__":
    success = run_performance_tests()
    sys.exit(0 if success else 1)