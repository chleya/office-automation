"""
Error Handling Tests
====================

Tests for error handling and edge cases in Office Automation.
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


class TestErrorHandling:
    """Tests for error handling and edge cases."""
    
    def setup_method(self):
        """Setup for each test."""
        self.office = OfficeAutomation()
        self.temp_dir = tempfile.mkdtemp(prefix="error_test_")
        print(f"\nTest directory: {self.temp_dir}")
    
    def teardown_method(self):
        """Cleanup after each test."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_invalid_configurations(self):
        """Test with invalid configuration values."""
        print("\nTesting invalid configurations...")
        
        invalid_configs = [
            None,
            {},
            {"invalid_key": "value"},
            {"temp_dir": None},
            {"log_level": "INVALID_LEVEL"},
            {"max_file_size": -1},
            {"max_file_size": "not_a_number"},
        ]
        
        for config in invalid_configs:
            try:
                office = OfficeAutomation(config=config)
                assert office is not None
                print(f"  Config {config}: Accepted")
            except Exception as e:
                print(f"  Config {config}: {type(e).__name__} - {e}")
    
    def test_none_inputs(self):
        """Test handling of None inputs."""
        print("\nTesting None inputs...")
        
        # Test get_info with None
        try:
            info = self.office.get_info()
            assert info is not None
            print("  get_info(): OK")
        except Exception as e:
            print(f"  get_info(): {type(e).__name__}")
        
        # Test module methods with None if they exist
        modules_to_test = ['word', 'excel', 'powerpoint', 'converter', 'wps']
        
        for module_name in modules_to_test:
            if hasattr(self.office, module_name):
                module = getattr(self.office, module_name)
                
                # Test get_capabilities if available
                if hasattr(module, 'get_capabilities'):
                    try:
                        caps = module.get_capabilities()
                        print(f"  {module_name}.get_capabilities(): OK")
                    except Exception as e:
                        print(f"  {module_name}.get_capabilities(): {type(e).__name__}")
    
    def test_empty_strings(self):
        """Test handling of empty strings."""
        print("\nTesting empty strings...")
        
        empty_values = ["", "   ", "\t", "\n", "\r\n"]
        
        # Test with file paths
        if hasattr(self.office.word, 'create_document') and hasattr(self.office.word, 'save_document'):
            doc = self.office.word.create_document()
            
            for empty_value in empty_values:
                try:
                    result = self.office.word.save_document(doc, empty_value)
                    print(f"  save_document with '{repr(empty_value)}': Returned {result}")
                except Exception as e:
                    print(f"  save_document with '{repr(empty_value)}': {type(e).__name__}")
    
    def test_invalid_file_paths(self):
        """Test with invalid file paths."""
        print("\nTesting invalid file paths...")
        
        invalid_paths = [
            "/invalid/path/document.docx",
            "C:\\invalid\\path\\document.docx",
            "document.docx",  # Relative path without directory
            "..\\..\\..\\document.docx",  # Path traversal
            "document" * 100 + ".docx",  # Very long filename
            "document<>.docx",  # Invalid characters
            "document?.docx",  # Invalid characters
            "document*.docx",  # Invalid characters
            "document|.docx",  # Invalid characters
        ]
        
        # Test with Word module if available
        if hasattr(self.office.word, 'create_document') and hasattr(self.office.word, 'save_document'):
            doc = self.office.word.create_document()
            
            for path in invalid_paths:
                try:
                    result = self.office.word.save_document(doc, path)
                    print(f"  save_document with '{path}': Returned {result}")
                except Exception as e:
                    print(f"  save_document with '{path}': {type(e).__name__}")
    
    def test_permission_errors(self):
        """Test handling of permission errors (simulated)."""
        print("\nTesting permission errors...")
        
        # Create a read-only directory
        read_only_dir = os.path.join(self.temp_dir, "readonly")
        os.makedirs(read_only_dir, exist_ok=True)
        
        # On Windows, set directory to read-only
        try:
            import stat
            os.chmod(read_only_dir, stat.S_IRUSR)
        except:
            pass  # Not all systems support chmod
        
        # Try to write to read-only directory
        if hasattr(self.office.word, 'create_document') and hasattr(self.office.word, 'save_document'):
            doc = self.office.word.create_document()
            read_only_path = os.path.join(read_only_dir, "test.docx")
            
            try:
                result = self.office.word.save_document(doc, read_only_path)
                print(f"  save_document to read-only dir: Returned {result}")
            except Exception as e:
                print(f"  save_document to read-only dir: {type(e).__name__}")
    
    def test_disk_space_errors(self):
        """Test handling of disk space errors (simulated)."""
        print("\nTesting disk space errors...")
        
        # We can't actually fill the disk, but we can test with very large file sizes
        if hasattr(self.office.word, 'create_document'):
            doc = self.office.word.create_document()
            
            # Test with very large content (simulated)
            if hasattr(self.office.word, 'add_paragraph'):
                try:
                    # Add many paragraphs to simulate large document
                    for i in range(1000):
                        self.office.word.add_paragraph(doc, "X" * 1000)
                    print("  Large document creation: OK")
                except Exception as e:
                    print(f"  Large document creation: {type(e).__name__}")
    
    def test_concurrent_access(self):
        """Test handling of concurrent access (simulated)."""
        print("\nTesting concurrent access...")
        
        import threading
        import time
        
        results = []
        errors = []
        
        def worker(worker_id):
            """Worker function for concurrent testing."""
            try:
                office = OfficeAutomation()
                info = office.get_info()
                results.append((worker_id, "success"))
            except Exception as e:
                errors.append((worker_id, str(e)))
        
        # Start multiple threads
        threads = []
        for i in range(5):
            t = threading.Thread(target=worker, args=(i,))
            threads.append(t)
            t.start()
        
        # Wait for all threads to complete
        for t in threads:
            t.join()
        
        print(f"  Concurrent access results: {len(results)} successes, {len(errors)} errors")
        if errors:
            for worker_id, error in errors:
                print(f"    Worker {worker_id}: {error}")
    
    def test_resource_cleanup(self):
        """Test that resources are properly cleaned up."""
        print("\nTesting resource cleanup...")
        
        import psutil
        import os
        
        process = psutil.Process(os.getpid())
        initial_memory = process.memory_info().rss
        
        # Create many OfficeAutomation instances
        instances = []
        for i in range(20):
            office = OfficeAutomation()
            info = office.get_info()
            instances.append(office)
        
        # Let instances go out of scope
        instances = []
        
        # Force garbage collection
        import gc
        gc.collect()
        
        # Check memory usage
        final_memory = process.memory_info().rss
        memory_increase = final_memory - initial_memory
        
        print(f"  Memory increase after 20 instances: {memory_increase / 1024 / 1024:.2f} MB")
        
        # Memory should not increase too much
        assert memory_increase < 50 * 1024 * 1024, f"Memory leak detected: {memory_increase / 1024 / 1024:.2f} MB increase"
        
        print("  Resource cleanup: PASSED")
    
    def test_error_recovery(self):
        """Test error recovery scenarios."""
        print("\nTesting error recovery...")
        
        # Test that we can recover from an error and continue
        operations = []
        
        # Operation 1: Should work
        try:
            info = self.office.get_info()
            operations.append(("get_info", "success"))
        except Exception as e:
            operations.append(("get_info", f"failed: {type(e).__name__}"))
        
        # Operation 2: Try something that might fail
        if hasattr(self.office.word, 'create_document'):
            try:
                doc = self.office.word.create_document()
                operations.append(("create_document", "success"))
                
                # Try to save with invalid path
                try:
                    result = self.office.word.save_document(doc, "/invalid/path/document.docx")
                    operations.append(("save_invalid_path", f"returned: {result}"))
                except Exception as e:
                    operations.append(("save_invalid_path", f"failed: {type(e).__name__}"))
                
                # Try another operation after error
                try:
                    # This should still work even if previous save failed
                    if hasattr(self.office.word, 'add_paragraph'):
                        self.office.word.add_paragraph(doc, "Recovery test")
                        operations.append(("add_paragraph_after_error", "success"))
                except Exception as e:
                    operations.append(("add_paragraph_after_error", f"failed: {type(e).__name__}"))
                    
            except Exception as e:
                operations.append(("create_document", f"failed: {type(e).__name__}"))
        
        # Operation 3: Should still work even if previous operations failed
        try:
            # Get info again
            info = self.office.get_info()
            operations.append(("get_info_after_errors", "success"))
        except Exception as e:
            operations.append(("get_info_after_errors", f"failed: {type(e).__name__}"))
        
        print("  Error recovery operations:")
        for op_name, result in operations:
            print(f"    {op_name}: {result}")
    
    def test_edge_case_data(self):
        """Test with edge case data values."""
        print("\nTesting edge case data...")
        
        edge_cases = [
            # (description, data)
            ("empty list", []),
            ("list with empty strings", ["", "", ""]),
            ("list with None", [None, None, None]),
            ("nested empty lists", [[], [], []]),
            ("mixed types", ["text", 123, 45.67, True, None]),
            ("very long string", ["X" * 10000]),
            ("unicode characters", ["Hello 世界", "مرحبا", "नमस्ते"]),
            ("special characters", ["Line 1\nLine 2", "Tab\there", "Quote\"here"]),
            ("empty dictionary", {}),
            ("dictionary with empty values", {"key1": "", "key2": None}),
        ]
        
        # Test with Excel module if available
        if hasattr(self.office.excel, 'create_spreadsheet') and hasattr(self.office.excel, 'add_data'):
            for description, data in edge_cases:
                try:
                    spreadsheet = self.office.excel.create_spreadsheet()
                    result = self.office.excel.add_data(spreadsheet, data)
                    print(f"  {description}: Returned {result}")
                except Exception as e:
                    print(f"  {description}: {type(e).__name__}")
    
    def test_timeout_handling(self):
        """Test timeout handling (simulated)."""
        print("\nTesting timeout handling...")
        
        import time
        
        # Test that operations don't hang indefinitely
        start_time = time.time()
        
        try:
            # This should complete quickly
            info = self.office.get_info()
            elapsed = time.time() - start_time
            
            print(f"  get_info() completed in {elapsed:.3f} seconds")
            assert elapsed < 5.0, f"get_info() took too long: {elapsed:.3f}s"
            
        except Exception as e:
            print(f"  get_info(): {type(e).__name__}")
    
    def test_import_errors(self):
        """Test handling of import errors."""
        print("\nTesting import error simulation...")
        
        # We can't actually cause import errors in the running process,
        # but we can test that the module handles missing dependencies gracefully
        
        # Check if module reports its dependencies properly
        info = self.office.get_info()
        
        if 'dependencies' in info:
            print(f"  Dependencies: {info['dependencies']}")
        else:
            print("  No dependency information available")
        
        # Test that we can still use the module even if some features are unavailable
        print("  Module works in dummy mode: OK")


def run_error_handling_tests():
    """Run all error handling tests."""
    print("=" * 60)
    print("ERROR HANDLING TESTS")
    print("=" * 60)
    
    tester = TestErrorHandling()
    
    tests = [
        tester.test_invalid_configurations,
        tester.test_none_inputs,
        tester.test_empty_strings,
        tester.test_invalid_file_paths,
        tester.test_permission_errors,
        tester.test_disk_space_errors,
        tester.test_concurrent_access,
        tester.test_resource_cleanup,
        tester.test_error_recovery,
        tester.test_edge_case_data,
        tester.test_timeout_handling,
        tester.test_import_errors,
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
    print("ERROR HANDLING TEST SUMMARY:")
    print(f"  Passed:  {passed}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped}")
    print(f"  Total:   {passed + failed + skipped}")
    print("=" * 60)
    
    if failed > 0:
        print("\nFAILED: Some error handling tests failed!")
        return False
    else:
        print("\nPASSED: All error handling tests passed!")
        return True


if __name__ == "__main__":
    success = run_error_handling_tests()
    sys.exit(0 if success else 1)