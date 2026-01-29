#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS Office Integration Demo - Fixed Version
===========================================

This example demonstrates WPS Office integration features.
"""

import os
import sys
import json
from pathlib import Path
from datetime import datetime

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Simple WPS integration demo
class WPSIntegrationDemo:
    def __init__(self):
        self.wps_available = self.detect_wps()
        self.wps_version = self.get_wps_version() if self.wps_available else None
        print("WPS Integration Demo initialized")
    
    def detect_wps(self):
        """Detect if WPS Office is installed."""
        print("Detecting WPS Office installation...")
        
        # Common WPS installation paths
        wps_paths = [
            "C:\\Program Files\\WPS Office",
            "C:\\Program Files (x86)\\WPS Office",
            os.path.expandvars("%ProgramFiles%\\WPS Office"),
            os.path.expandvars("%ProgramFiles(x86)%\\WPS Office"),
            os.path.expandvars("%LOCALAPPDATA%\\Kingsoft\\WPS Office"),
        ]
        
        for path in wps_paths:
            if os.path.exists(path):
                print(f"  Found WPS at: {path}")
                return True
        
        print("  WPS Office not detected")
        return False
    
    def get_wps_version(self):
        """Get WPS Office version if available."""
        print("Getting WPS version...")
        version = "Unknown (detected but version not found)"
        print(f"  {version}")
        return version
    
    def check_wps_compatibility(self, file_path):
        """Check if a file is compatible with WPS Office."""
        print(f"Checking WPS compatibility for: {file_path}")
        
        if not os.path.exists(file_path):
            print("  File does not exist")
            return {'compatible': False, 'reason': 'File not found'}
        
        file_ext = os.path.splitext(file_path)[1].lower()
        
        # WPS supported formats
        wps_formats = {
            '.doc': 'Word Document',
            '.docx': 'Word Document',
            '.wps': 'WPS Document',
            '.et': 'WPS Spreadsheet',
            '.xls': 'Excel Spreadsheet',
            '.xlsx': 'Excel Spreadsheet',
            '.dps': 'WPS Presentation',
            '.ppt': 'PowerPoint Presentation',
            '.pptx': 'PowerPoint Presentation',
            '.pdf': 'PDF Document',
            '.txt': 'Text File',
        }
        
        if file_ext in wps_formats:
            format_name = wps_formats[file_ext]
            print(f"  Format: {format_name} ({file_ext})")
            print(f"  WPS compatibility: YES")
            return {
                'compatible': True,
                'format': format_name,
                'extension': file_ext,
                'notes': 'Fully supported by WPS Office'
            }
        else:
            print(f"  Format: {file_ext}")
            print(f"  WPS compatibility: MAYBE (unknown format)")
            return {
                'compatible': True,
                'format': 'Unknown',
                'extension': file_ext,
                'notes': 'Format not in known list, may still work'
            }
    
    def optimize_for_wps(self, input_path, output_path=None):
        """Optimize a document for better WPS compatibility."""
        print(f"Optimizing document for WPS: {input_path}")
        
        if not self.wps_available:
            print("  WPS not available, cannot optimize")
            return {'success': False, 'reason': 'WPS not installed'}
        
        if not os.path.exists(input_path):
            print("  Input file does not exist")
            return {'success': False, 'reason': 'File not found'}
        
        # Create output path if not provided
        if output_path is None:
            input_dir = os.path.dirname(input_path)
            input_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(input_dir, f"{input_name}_wps_optimized.docx")
        
        print(f"  Output: {output_path}")
        
        # For demo, just copy the file
        try:
            import shutil
            shutil.copy2(input_path, output_path)
            
            # Add optimization metadata
            meta_path = output_path + ".meta.json"
            metadata = {
                'original_file': input_path,
                'optimized_file': output_path,
                'optimization_date': datetime.now().isoformat(),
                'wps_version': self.wps_version,
                'optimizations_applied': [
                    'File copied (demo mode)',
                    'In real implementation: format optimization',
                    'In real implementation: compatibility fixes',
                    'In real implementation: WPS feature enhancement'
                ]
            }
            
            with open(meta_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            
            print("  Optimization complete (demo mode)")
            print(f"  Metadata saved: {meta_path}")
            
            return {
                'success': True,
                'original': input_path,
                'optimized': output_path,
                'metadata': meta_path,
                'notes': 'Demo mode - file copied with metadata'
            }
            
        except Exception as e:
            print(f"  Optimization failed: {e}")
            return {'success': False, 'reason': str(e)}
    
    def compare_wps_vs_msoffice(self):
        """Compare WPS Office vs Microsoft Office features."""
        print("Comparing WPS Office vs Microsoft Office...")
        
        comparison = {
            'document_formats': {
                'wps': ['.wps', '.doc', '.docx', '.pdf'],
                'msoffice': ['.doc', '.docx', '.dot', '.dotx']
            },
            'spreadsheet_formats': {
                'wps': ['.et', '.xls', '.xlsx'],
                'msoffice': ['.xls', '.xlsx', '.xlt', '.xltx']
            },
            'presentation_formats': {
                'wps': ['.dps', '.ppt', '.pptx'],
                'msoffice': ['.ppt', '.pptx', '.pot', '.potx']
            },
            'key_differences': [
                'WPS has smaller installation size',
                'WPS has better PDF support built-in',
                'MS Office has more advanced collaboration features',
                'WPS has better compatibility with Chinese documents',
                'MS Office has more third-party integrations',
                'WPS is generally more affordable'
            ],
            'compatibility_notes': [
                'Most basic documents work interchangeably',
                'Complex macros may need adjustment',
                'Advanced formatting may render differently',
                'Cloud integration differs between platforms'
            ]
        }
        
        print("  Feature comparison:")
        for category, data in comparison.items():
            if isinstance(data, dict):
                print(f"    {category}:")
                for app, formats in data.items():
                    print(f"      {app}: {', '.join(formats)}")
            elif isinstance(data, list):
                print(f"    {category}:")
                for item in data:
                    print(f"      - {item}")
        
        return comparison
    
    def generate_wps_report(self, output_dir):
        """Generate a comprehensive WPS compatibility report."""
        print(f"Generating WPS compatibility report in: {output_dir}")
        
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(output_dir, f"wps_compatibility_report_{timestamp}.md")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("# WPS Office Compatibility Report\n\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            f.write("## System Information\n")
            f.write(f"- WPS Detected: {'Yes' if self.wps_available else 'No'}\n")
            if self.wps_available:
                f.write(f"- WPS Version: {self.wps_version}\n")
            f.write(f"- Python Version: {sys.version}\n")
            f.write(f"- Platform: {sys.platform}\n\n")
            
            f.write("## Feature Comparison\n")
            comparison = self.compare_wps_vs_msoffice()
            
            f.write("### Document Formats\n")
            for app, formats in comparison['document_formats'].items():
                f.write(f"- **{app}**: {', '.join(formats)}\n")
            f.write("\n")
            
            f.write("### Key Differences\n")
            for diff in comparison['key_differences']:
                f.write(f"- {diff}\n")
            f.write("\n")
            
            f.write("### Compatibility Notes\n")
            for note in comparison['compatibility_notes']:
                f.write(f"- {note}\n")
            f.write("\n")
            
            f.write("## Recommendations\n")
            f.write("1. **For basic documents**: WPS and MS Office are highly compatible\n")
            f.write("2. **For complex documents**: Test compatibility before deployment\n")
            f.write("3. **For macros/VBA**: May require adjustments for WPS\n")
            f.write("4. **For cloud collaboration**: Consider platform-specific features\n")
            f.write("5. **For Chinese documents**: WPS may have better support\n")
            f.write("\n")
            
            f.write("## Next Steps\n")
            f.write("1. Install WPS Office if not already installed\n")
            f.write("2. Test your specific document types\n")
            f.write("3. Use the optimization tools if needed\n")
            f.write("4. Monitor for any compatibility issues\n")
        
        print(f"  Report generated: {report_path}")
        return report_path


def main():
    """Main function to run WPS integration demo."""
    print("=" * 60)
    print("WPS OFFICE INTEGRATION DEMO")
    print("=" * 60)
    
    # Initialize demo
    demo = WPSIntegrationDemo()
    
    # Create test directory
    test_dir = project_root / "output" / "wps_tests"
    test_dir.mkdir(parents=True, exist_ok=True)
    
    # Create a test document
    test_doc = test_dir / "test_document.docx"
    with open(test_doc, 'w', encoding='utf-8') as f:
        f.write("This is a test document for WPS compatibility testing.\n")
        f.write("Created: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
    
    print(f"\n1. Created test document: {test_doc}")
    
    # Check compatibility
    print("\n2. Checking WPS compatibility...")
    compatibility = demo.check_wps_compatibility(str(test_doc))
    
    # Compare WPS vs MS Office
    print("\n3. Comparing WPS vs Microsoft Office...")
    comparison = demo.compare_wps_vs_msoffice()
    
    # Optimize for WPS (if available)
    if demo.wps_available:
        print("\n4. Optimizing document for WPS...")
        optimized = demo.optimize_for_wps(str(test_doc))
    else:
        print("\n4. Skipping optimization (WPS not available)")
        optimized = {'success': False, 'reason': 'WPS not installed'}
    
    # Generate report
    print("\n5. Generating compatibility report...")
    report_path = demo.generate_wps_report(str(test_dir))
    
    print("\n" + "=" * 60)
    print("DEMO COMPLETE")
    print("=" * 60)
    
    print(f"\nSummary:")
    print(f"- WPS Available: {'Yes' if demo.wps_available else 'No'}")
    if demo.wps_available:
        print(f"- WPS Version: {demo.wps_version}")
    print(f"- Test Document: {test_doc}")
    print(f"- WPS Compatible: {'Yes' if compatibility['compatible'] else 'No'}")
    print(f"- Optimization: {'Success' if optimized['success'] else 'Failed'}")
    print(f"- Report: {report_path}")
    
    print(f"\nFiles created in '{test_dir}':")
    for file in os.listdir(test_dir):
        file_path = os.path.join(test_dir, file)
        if os.path.isfile(file_path):
            size = os.path.getsize(file_path)
            print(f"  - {file} ({size} bytes)")
    
    print("\nNext steps:")
    print("1. Install WPS Office for full functionality")
    print("2. Test with real Office documents")
    print("3. Implement actual optimization algorithms")
    print("4. Add more WPS-specific features")
    
    return True


if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\nError in demo: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)