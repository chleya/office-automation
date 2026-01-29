"""
WPS Office Integration

This module provides integration with WPS Office for enhanced compatibility
and access to WPS-specific features when available.
"""

import os
import sys
import subprocess
from typing import Optional, Dict, Any, List, Union
from pathlib import Path
import warnings

from ..utils.error_handler import (
    WPSNotAvailableError,
    DocumentCreationError,
    FormatConversionError,
    create_error_context,
)


class WPSIntegration:
    """Integration with WPS Office for enhanced document processing."""
    
    def __init__(self):
        """Initialize WPS Integration."""
        self.wps_available = self._check_wps_installation()
        self.wps_version = self._get_wps_version() if self.wps_available else None
        self.wps_features = self._detect_wps_features() if self.wps_available else {}
    
    def _check_wps_installation(self) -> bool:
        """
        Check if WPS Office is installed on the system.
        
        Returns:
            True if WPS Office is detected, False otherwise
        """
        # Check common WPS installation paths on Windows
        wps_paths = [
            # Common installation paths
            r"C:\Program Files\WPS Office",
            r"C:\Program Files (x86)\WPS Office",
            r"C:\Program Files\Kingsoft\WPS Office",
            # User installation
            os.path.expanduser(r"~\AppData\Local\Kingsoft\WPS Office"),
        ]
        
        for path in wps_paths:
            if os.path.exists(path):
                return True
        
        # Check registry (Windows specific)
        if sys.platform == "win32":
            try:
                import winreg
                
                registry_paths = [
                    r"SOFTWARE\Kingsoft\Office",
                    r"SOFTWARE\WOW6432Node\Kingsoft\Office",
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WPS Office",
                ]
                
                for reg_path in registry_paths:
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                        winreg.CloseKey(key)
                        return True
                    except WindowsError:
                        continue
                        
            except ImportError:
                # winreg not available (not Windows or Python issue)
                pass
        
        # Check PATH for WPS executables
        wps_executables = ["wps.exe", "et.exe", "wpp.exe", "wpsoffice.exe"]
        
        for executable in wps_executables:
            try:
                subprocess.run([executable, "--version"], 
                             capture_output=True, 
                             timeout=2,
                             shell=True)
                return True
            except (subprocess.SubprocessError, FileNotFoundError):
                continue
        
        return False
    
    def _get_wps_version(self) -> Optional[str]:
        """
        Get WPS Office version if installed.
        
        Returns:
            Version string or None if not available
        """
        if not self.wps_available:
            return None
        
        try:
            # Try to get version from registry (Windows)
            if sys.platform == "win32":
                import winreg
                
                try:
                    key = winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE,
                        r"SOFTWARE\Kingsoft\Office\6.0\common"
                    )
                    version, _ = winreg.QueryValueEx(key, "Version")
                    winreg.CloseKey(key)
                    return version
                except WindowsError:
                    pass
            
            # Try to get version from executable
            wps_paths = [
                r"C:\Program Files\WPS Office\11.2.0.12388\office6\wps.exe",
                r"C:\Program Files (x86)\WPS Office\11.2.0.12388\office6\wps.exe",
            ]
            
            for exe_path in wps_paths:
                if os.path.exists(exe_path):
                    try:
                        result = subprocess.run(
                            [exe_path, "--version"],
                            capture_output=True,
                            text=True,
                            timeout=5
                        )
                        if result.returncode == 0:
                            return result.stdout.strip()
                    except (subprocess.SubprocessError, FileNotFoundError):
                        continue
            
            return "Unknown (WPS detected but version not found)"
            
        except Exception:
            return "Unknown"
    
    def _detect_wps_features(self) -> Dict[str, bool]:
        """
        Detect available WPS-specific features.
        
        Returns:
            Dictionary of feature availability
        """
        features = {
            "pdf_export": True,  # WPS has built-in PDF export
            "template_gallery": True,  # WPS has template gallery
            "cloud_integration": True,  # WPS Cloud integration
            "collaboration": True,  # Real-time collaboration
            "ai_features": False,  # AI features (may require subscription)
            "mobile_sync": True,  # Mobile device sync
            "chinese_support": True,  # Enhanced Chinese language support
        }
        
        # Try to detect AI features
        try:
            # Check for AI-related components
            ai_paths = [
                r"C:\Program Files\WPS Office\ai",
                r"C:\Program Files (x86)\WPS Office\ai",
            ]
            
            for path in ai_paths:
                if os.path.exists(path):
                    features["ai_features"] = True
                    break
        except Exception:
            pass
        
        return features
    
    def convert_to_wps_format(
        self,
        input_file: Union[str, Path],
        output_file: Optional[Union[str, Path]] = None,
        format: str = "wps"
    ) -> str:
        """
        Convert document to WPS-specific format.
        
        Args:
            input_file: Path to input file
            output_file: Path for output file (optional)
            format: Target format ("wps", "et", "dps" for WPS equivalents)
            
        Returns:
            Path to converted file
            
        Raises:
            WPSNotAvailableError: If WPS is not installed
            FormatConversionError: If conversion fails
        """
        if not self.wps_available:
            raise WPSNotAvailableError(
                "WPS Office is not installed. Cannot convert to WPS format."
            )
        
        context = create_error_context(
            "convert_to_wps_format",
            file_path=input_file,
            format=format
        )
        
        try:
            input_file = Path(input_file)
            if not input_file.exists():
                raise FileNotFoundError(f"Input file not found: {input_file}")
            
            # Determine output filename
            if output_file is None:
                # Map format to extension
                format_extensions = {
                    "wps": ".wps",  # WPS Writer
                    "et": ".et",    # WPS Spreadsheets
                    "dps": ".dps",  # WPS Presentation
                }
                ext = format_extensions.get(format.lower(), f".{format}")
                output_file = input_file.with_suffix(ext)
            else:
                output_file = Path(output_file)
            
            # In a real implementation, this would call WPS Office APIs
            # or use WPS command-line tools to perform the conversion
            
            # For now, we'll create a placeholder implementation
            # that copies the file and changes extension
            
            import shutil
            shutil.copy2(input_file, output_file)
            
            warnings.warn(
                f"WPS format conversion is simulated. "
                f"File copied with {output_file.suffix} extension. "
                f"In production, use WPS Office APIs for actual conversion."
            )
            
            return str(output_file)
            
        except Exception as e:
            raise FormatConversionError(
                f"Failed to convert to WPS format: {e}",
                context
            )
    
    def use_wps_template(
        self,
        template_name: str,
        output_file: Optional[Union[str, Path]] = None,
        data: Optional[Dict[str, Any]] = None
    ) -> str:
        """
        Create document using WPS template.
        
        Args:
            template_name: Name of WPS template
            output_file: Path for output file (optional)
            data: Data to fill in template
            
        Returns:
            Path to created document
            
        Raises:
            WPSNotAvailableError: If WPS is not installed
            DocumentCreationError: If template creation fails
        """
        if not self.wps_available:
            raise WPSNotAvailableError(
                "WPS Office is not installed. Cannot use WPS templates."
            )
        
        context = create_error_context(
            "use_wps_template",
            template=template_name
        )
        
        try:
            # Determine output filename
            if output_file is None:
                # Guess output format based on template name
                if template_name.endswith((".wpt", ".dot", ".dotx")):
                    output_file = Path("document.wps")
                elif template_name.endswith((".xlt", ".xltx")):
                    output_file = Path("spreadsheet.et")
                elif template_name.endswith((".pot", ".potx")):
                    output_file = Path("presentation.dps")
                else:
                    output_file = Path("document.wps")
            else:
                output_file = Path(output_file)
            
            # In a real implementation, this would:
            # 1. Locate WPS template (from WPS template gallery or local storage)
            # 2. Use WPS APIs to create document from template
            # 3. Fill template with provided data
            # 4. Save to output file
            
            # For now, create a placeholder file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"WPS Template: {template_name}\n")
                f.write("=" * 40 + "\n")
                if data:
                    for key, value in data.items():
                        f.write(f"{key}: {value}\n")
                f.write("\n[This is a placeholder for WPS template-based document]\n")
            
            warnings.warn(
                f"WPS template usage is simulated. "
                f"Created placeholder file at {output_file}. "
                f"In production, use WPS Office APIs for template-based document creation."
            )
            
            return str(output_file)
            
        except Exception as e:
            raise DocumentCreationError(
                f"Failed to create document from WPS template: {e}",
                context
            )
    
    def execute_wps_macro(
        self,
        document_file: Union[str, Path],
        macro_name: str,
        parameters: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Execute a WPS macro on a document.
        
        Args:
            document_file: Path to document file
            macro_name: Name of macro to execute
            parameters: Macro parameters
            
        Returns:
            Dictionary with execution results
            
        Raises:
            WPSNotAvailableError: If WPS is not installed
        """
        if not self.wps_available:
            raise WPSNotAvailableError(
                "WPS Office is not installed. Cannot execute WPS macros."
            )
        
        # In a real implementation, this would:
        # 1. Load document in WPS
        # 2. Execute specified macro
        # 3. Capture results
        
        warnings.warn(
            f"WPS macro execution is not implemented. "
            f"Macro '{macro_name}' would be executed on {document_file}. "
            f"In production, use WPS Office automation APIs."
        )
        
        return {
            "success": False,
            "message": "WPS macro execution not implemented",
            "macro": macro_name,
            "document": str(document_file),
            "parameters": parameters,
        }
    
    def get_wps_cloud_documents(self) -> List[Dict[str, Any]]:
        """
        Get list of documents from WPS Cloud.
        
        Returns:
            List of document information dictionaries
            
        Raises:
            WPSNotAvailableError: If WPS is not installed
        """
        if not self.wps_available:
            raise WPSNotAvailableError(
                "WPS Office is not installed. Cannot access WPS Cloud."
            )
        
        # In a real implementation, this would:
        # 1. Authenticate with WPS Cloud
        # 2. Retrieve document list
        # 3. Return structured document information
        
        warnings.warn(
            "WPS Cloud integration is not implemented. "
            "In production, use WPS Cloud APIs."
        )
        
        return [
            {
                "name": "Sample Document.wps",
                "type": "document",
                "size": "15KB",
                "modified": "2024-01-28",
                "cloud_url": "wpscloud://sample/document",
            },
            {
                "name": "Sample Spreadsheet.et",
                "type": "spreadsheet",
                "size": "25KB",
                "modified": "2024-01-27",
                "cloud_url": "wpscloud://sample/spreadsheet",
            },
        ]
    
    def optimize_for_wps(
        self,
        document_file: Union[str, Path],
        options: Optional[Dict[str, Any]] = None
    ) -> str:
        """
        Optimize document for best compatibility with WPS Office.
        
        Args:
            document_file: Path to document file
            options: Optimization options
            
        Returns:
            Path to optimized document
            
        Raises:
            WPSNotAvailableError: If WPS is not installed
        """
        if not self.wps_available:
            raise WPSNotAvailableError(
                "WPS Office is not installed. Cannot optimize for WPS."
            )
        
        # In a real implementation, this would:
        # 1. Analyze document for WPS compatibility issues
        # 2. Apply optimizations (remove unsupported features, convert formats, etc.)
        # 3. Save optimized version
        
        document_path = Path(document_file)
        optimized_path = document_path.with_stem(f"{document_path.stem}_wps_optimized")
        
        warnings.warn(
            f"WPS optimization is simulated. "
            f"Would optimize {document_file} for WPS compatibility. "
            f"In production, use WPS compatibility tools."
        )
        
        # For now, just copy the file
        import shutil
        shutil.copy2(document_file, optimized_path)
        
        return str(optimized_path)
    
    def get_wps_info(self) -> Dict[str, Any]:
        """
        Get information about WPS Office installation.
        
        Returns:
            Dictionary with WPS information
        """
        return {
            "available": self.wps_available,
            "version": self.wps_version,
            "features": self.wps_features,
            "platform": sys.platform,
            "detection_method": "path_and_registry" if sys.platform == "win32" else "path_only",
        }
    
    def is_feature_available(self, feature_name: str) -> bool:
        """
        Check if a specific WPS feature is available.
        
        Args:
            feature_name: Name of feature to check
            
        Returns:
            True if feature is available, False otherwise
        """
        if not self.wps_available:
            return False
        
        return self.wps_features.get(feature_name, False)
    
    def get_recommended_format(self, document_type: str) -> str:
        """
        Get recommended file format for WPS compatibility.
        
        Args:
            document_type: Type of document ("word", "excel", "powerpoint")
            
        Returns:
            Recommended file extension
        """
        recommendations = {
            "word": ".docx",  # .docx has best WPS compatibility
            "excel": ".xlsx",  # .xlsx has best WPS compatibility
            "powerpoint": ".pptx",  # .pptx has best WPS compatibility
            "text": ".txt",  # Plain text is always compatible
            "pdf": ".pdf",  # PDF is universally compatible
        }
        
        doc_type = document_type.lower()
        if doc_type in recommendations:
            return recommendations[doc_type]
        elif doc_type in ["doc", "document", "writer"]:
            return ".docx"
        elif doc_type in ["xls", "spreadsheet", "sheets"]:
            return ".xlsx"
        elif doc_type in ["ppt", "presentation", "slides"]:
            return ".pptx"
        else:
            return ".docx"  # Default to Word format