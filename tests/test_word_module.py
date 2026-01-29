"""
Word Module Tests
=================

Tests for the Word document automation module.
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


class TestWordModule:
    """Tests for the Word document module."""
    
    def setup_method(self):
        """Setup for each test."""
        self.office = OfficeAutomation()
        self.temp_dir = tempfile.mkdtemp(prefix="word_test_")
        print(f"\nTest directory: {self.temp_dir}")
    
    def teardown_method(self):
        """Cleanup after each test."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_word_module_available(self):
        """Test that Word module is available."""
        print("\nTesting Word module availability...")
        
        assert hasattr(self.office, 'word'), "Word module not found"
        assert self.office.word is not None, "Word module is None"
        
        # Check capabilities
        if hasattr(self.office.word, 'get_capabilities'):
            capabilities = self.office.word.get_capabilities()
            assert isinstance(capabilities, dict), "Capabilities should be a dict"
            assert 'available' in capabilities, "Capabilities missing 'available' key"
            print(f"  Word capabilities: {capabilities}")
        else:
            print("  Word module does not have get_capabilities method")
    
    def test_create_document(self):
        """Test document creation."""
        print("\nTesting document creation...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        
        # Test basic document creation
        doc = self.office.word.create_document()
        assert doc is not None, "Document creation returned None"
        
        # Test with template
        doc_with_template = self.office.word.create_document(template="business")
        assert doc_with_template is not None, "Document with template creation returned None"
        
        print("  Document creation: PASSED")
    
    def test_add_content(self):
        """Test adding content to document."""
        print("\nTesting content addition...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        
        # Create document
        doc = self.office.word.create_document()
        assert doc is not None
        
        # Test adding paragraph
        if hasattr(self.office.word, 'add_paragraph'):
            result = self.office.word.add_paragraph(
                doc,
                text="Test paragraph for Word module testing",
                style="Normal"
            )
            assert result is not None, "add_paragraph returned None"
            print("  add_paragraph: PASSED")
        
        # Test adding heading
        if hasattr(self.office.word, 'add_heading'):
            result = self.office.word.add_heading(
                doc,
                text="Test Heading",
                level=1
            )
            assert result is not None, "add_heading returned None"
            print("  add_heading: PASSED")
        
        # Test adding table
        if hasattr(self.office.word, 'add_table'):
            data = [
                ["Name", "Age", "Score"],
                ["Alice", "25", "95"],
                ["Bob", "30", "88"]
            ]
            result = self.office.word.add_table(doc, data)
            assert result is not None, "add_table returned None"
            print("  add_table: PASSED")
    
    def test_save_document(self):
        """Test saving document to file."""
        print("\nTesting document saving...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        if not hasattr(self.office.word, 'save_document'):
            pytest.skip("Word module does not have save_document method")
        
        # Create document
        doc = self.office.word.create_document()
        assert doc is not None
        
        # Add some content
        if hasattr(self.office.word, 'add_paragraph'):
            self.office.word.add_paragraph(doc, "Test content for saving")
        
        # Save document
        doc_path = os.path.join(self.temp_dir, "test_document.docx")
        result = self.office.word.save_document(doc, doc_path)
        assert result is not None, "save_document returned None"
        
        # Verify file was created
        assert os.path.exists(doc_path), f"Document not created: {doc_path}"
        file_size = os.path.getsize(doc_path)
        assert file_size > 0, f"Document file is empty: {doc_path}"
        
        print(f"  Document saved: {doc_path} ({file_size} bytes)")
        print("  Document saving: PASSED")
    
    def test_document_formats(self):
        """Test different document formats."""
        print("\nTesting document formats...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        if not hasattr(self.office.word, 'save_document'):
            pytest.skip("Word module does not have save_document method")
        
        # Create document
        doc = self.office.word.create_document()
        if hasattr(self.office.word, 'add_paragraph'):
            self.office.word.add_paragraph(doc, "Test document formats")
        
        # Test different formats
        formats = [
            ("docx", ".docx"),
            ("doc", ".doc"),
            ("pdf", ".pdf"),
            ("rtf", ".rtf"),
            ("txt", ".txt"),
        ]
        
        for format_name, extension in formats:
            try:
                doc_path = os.path.join(self.temp_dir, f"test_document{extension}")
                result = self.office.word.save_document(doc, doc_path, format=format_name)
                
                if result is not None:
                    # Check if file was created (some formats may not be supported)
                    if os.path.exists(doc_path):
                        file_size = os.path.getsize(doc_path)
                        print(f"  Format {format_name}: Created ({file_size} bytes)")
                    else:
                        print(f"  Format {format_name}: Not created (may not be supported)")
                else:
                    print(f"  Format {format_name}: save_document returned None")
                    
            except Exception as e:
                print(f"  Format {format_name}: Error - {type(e).__name__}")
    
    def test_document_metadata(self):
        """Test document metadata operations."""
        print("\nTesting document metadata...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        
        doc = self.office.word.create_document()
        
        # Test setting metadata if available
        if hasattr(self.office.word, 'set_metadata'):
            metadata = {
                'title': 'Test Document',
                'author': 'Test Author',
                'subject': 'Testing Word Module',
                'keywords': ['test', 'word', 'automation'],
                'category': 'Testing',
            }
            
            result = self.office.word.set_metadata(doc, metadata)
            assert result is not None, "set_metadata returned None"
            print("  set_metadata: PASSED")
        
        # Test getting metadata if available
        if hasattr(self.office.word, 'get_metadata'):
            metadata = self.office.word.get_metadata(doc)
            assert metadata is not None, "get_metadata returned None"
            assert isinstance(metadata, dict), "Metadata should be a dict"
            print(f"  get_metadata: {metadata}")
    
    def test_document_styles(self):
        """Test document style operations."""
        print("\nTesting document styles...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        
        doc = self.office.word.create_document()
        
        # Test adding styled content
        if hasattr(self.office.word, 'add_paragraph'):
            styles = ['Normal', 'Title', 'Heading 1', 'Heading 2', 'Quote']
            
            for style in styles:
                try:
                    result = self.office.word.add_paragraph(
                        doc,
                        text=f"Paragraph with style: {style}",
                        style=style
                    )
                    if result is not None:
                        print(f"  Style {style}: Added")
                    else:
                        print(f"  Style {style}: add_paragraph returned None")
                except Exception as e:
                    print(f"  Style {style}: Error - {type(e).__name__}")
    
    def test_error_handling(self):
        """Test error handling in Word module."""
        print("\nTesting error handling...")
        
        # Test with invalid inputs
        if hasattr(self.office.word, 'create_document'):
            # Test None input
            try:
                result = self.office.word.create_document(template=None)
                print("  create_document with None template: OK")
            except Exception as e:
                print(f"  create_document with None template: {type(e).__name__}")
        
        # Test with invalid file paths
        if hasattr(self.office.word, 'save_document'):
            doc = self.office.word.create_document() if hasattr(self.office.word, 'create_document') else "dummy"
            
            invalid_paths = [
                "/invalid/path/document.docx",
                "",
                None,
                "   ",
            ]
            
            for path in invalid_paths:
                try:
                    result = self.office.word.save_document(doc, path)
                    print(f"  save_document with invalid path '{path}': Returned {result}")
                except Exception as e:
                    print(f"  save_document with invalid path '{path}': {type(e).__name__}")
    
    def test_batch_operations(self):
        """Test batch document operations."""
        print("\nTesting batch operations...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        if not hasattr(self.office.word, 'save_document'):
            pytest.skip("Word module does not have save_document method")
        
        # Create multiple documents
        doc_count = 3
        documents = []
        
        for i in range(doc_count):
            doc = self.office.word.create_document()
            if hasattr(self.office.word, 'add_paragraph'):
                self.office.word.add_paragraph(doc, f"Document {i+1} content")
            documents.append(doc)
        
        # Save all documents
        for i, doc in enumerate(documents):
            doc_path = os.path.join(self.temp_dir, f"batch_doc_{i+1}.docx")
            result = self.office.word.save_document(doc, doc_path)
            assert result is not None, f"Failed to save document {i+1}"
            assert os.path.exists(doc_path), f"Document {i+1} not created"
        
        print(f"  Created and saved {doc_count} documents")
        print("  Batch operations: PASSED")
    
    def test_performance(self):
        """Test Word module performance."""
        print("\nTesting Word module performance...")
        
        if not hasattr(self.office.word, 'create_document'):
            pytest.skip("Word module does not have create_document method")
        
        import time
        
        # Time document creation
        start_time = time.perf_counter()
        doc = self.office.word.create_document()
        creation_time = time.perf_counter() - start_time
        
        print(f"  Document creation time: {creation_time:.4f}s")
        assert creation_time < 5.0, f"Document creation too slow: {creation_time:.4f}s"
        
        # Time content addition
        if hasattr(self.office.word, 'add_paragraph'):
            start_time = time.perf_counter()
            for i in range(10):
                self.office.word.add_paragraph(doc, f"Paragraph {i+1}")
            addition_time = time.perf_counter() - start_time
            
            print(f"  10 paragraph addition time: {addition_time:.4f}s")
            assert addition_time < 2.0, f"Paragraph addition too slow: {addition_time:.4f}s"
        
        print("  Performance test: PASSED")


def run_word_tests():
    """Run all Word module tests."""
    print("=" * 60)
    print("WORD MODULE TESTS")
    print("=" * 60)
    
    tester = TestWordModule()
    
    tests = [
        tester.test_word_module_available,
        tester.test_create_document,
        tester.test_add_content,
        tester.test_save_document,
        tester.test_document_formats,
        tester.test_document_metadata,
        tester.test_document_styles,
        tester.test_error_handling,
        tester.test_batch_operations,
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
    print("WORD MODULE TEST SUMMARY:")
    print(f"  Passed:  {passed}")
    print(f"  Failed:  {failed}")
    print(f"  Skipped: {skipped}")
    print(f"  Total:   {passed + failed + skipped}")
    print("=" * 60)
    
    if failed > 0:
        print("\nFAILED: Some Word module tests failed!")
        return False
    else:
        print("\nPASSED: All Word module tests passed!")
        return True


if __name__ == "__main__":
    success = run_word_tests()
    sys.exit(0 if success else 1)