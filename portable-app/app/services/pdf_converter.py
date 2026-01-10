"""
PDF Converter Service
Handles converting Word documents to PDF using LibreOffice.
No Microsoft Word dependency.
"""

import os
import subprocess
import shutil
import platform
from pathlib import Path
from typing import Optional, Tuple, List


class PDFConverter:
    """
    Converts Word documents to PDF using LibreOffice headless mode.
    Works on Windows, macOS, and Linux without Microsoft Word.
    """
    
    def __init__(
        self,
        libreoffice_path: Optional[str] = None,
        timeout_seconds: int = 60
    ):
        """
        Initialize the PDF converter.
        
        Args:
            libreoffice_path: Path to LibreOffice executable (auto-detect if None)
            timeout_seconds: Timeout for conversion process
        """
        self.timeout_seconds = timeout_seconds
        self.libreoffice_path = libreoffice_path or self._find_libreoffice()
        
        if not self.libreoffice_path:
            raise RuntimeError(
                "LibreOffice not found. Please install LibreOffice or provide the path."
            )
    
    def _find_libreoffice(self) -> Optional[str]:
        """
        Auto-detect LibreOffice installation.
        
        Returns:
            Path to LibreOffice executable, or None if not found
        """
        system = platform.system()
        
        # Common paths to check
        possible_paths = []
        
        if system == "Windows":
            possible_paths = [
                # Bundled LibreOffice (relative to app folder - from build)
                Path(__file__).parent.parent.parent / "libreoffice" / "program" / "soffice.exe",
                # Bundled LibreOffice Portable structure
                Path(__file__).parent.parent.parent / "libreoffice" / "App" / "libreoffice" / "program" / "soffice.exe",
                # Standard installations
                Path(os.environ.get("PROGRAMFILES", "C:\\Program Files")) / "LibreOffice" / "program" / "soffice.exe",
                Path(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)")) / "LibreOffice" / "program" / "soffice.exe",
                Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
            ]
        elif system == "Darwin":  # macOS
            possible_paths = [
                Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
                Path.home() / "Applications" / "LibreOffice.app" / "Contents" / "MacOS" / "soffice",
            ]
        else:  # Linux
            possible_paths = [
                Path("/usr/bin/soffice"),
                Path("/usr/bin/libreoffice"),
                Path("/usr/local/bin/soffice"),
            ]
        
        # Check each path
        for path in possible_paths:
            if path.exists():
                return str(path)
        
        # Try finding in PATH
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            return soffice
        
        return None
    
    def convert(self, docx_path: str, output_dir: str) -> Tuple[bool, str, Optional[str]]:
        """
        Convert a Word document to PDF.
        
        Args:
            docx_path: Path to the .docx file
            output_dir: Directory to save the PDF
        
        Returns:
            Tuple of (success, pdf_path or error_message, error_details)
        """
        docx_path = Path(docx_path)
        output_dir = Path(output_dir)
        
        if not docx_path.exists():
            return False, f"File not found: {docx_path}", None
        
        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Build LibreOffice command
        # --headless: No GUI
        # --convert-to pdf: Convert to PDF format
        # --outdir: Output directory
        cmd = [
            self.libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(docx_path)
        ]
        
        try:
            # Run conversion
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.timeout_seconds,
                cwd=str(output_dir)  # Set working directory
            )
            
            # Check for expected output file
            expected_pdf = output_dir / (docx_path.stem + ".pdf")
            
            if expected_pdf.exists():
                return True, str(expected_pdf), None
            
            # Conversion might have failed
            error_msg = result.stderr or result.stdout or "Unknown error"
            return False, f"Conversion failed: {error_msg}", result.stderr
            
        except subprocess.TimeoutExpired:
            return False, f"Conversion timed out after {self.timeout_seconds} seconds", None
        except Exception as e:
            return False, f"Conversion error: {str(e)}", str(e)
    
    def convert_batch(
        self,
        docx_files: List[str],
        output_dir: str,
        progress_callback: Optional[callable] = None
    ) -> List[Tuple[str, bool, str]]:
        """
        Convert multiple Word documents to PDF (one at a time).
        
        Args:
            docx_files: List of paths to .docx files
            output_dir: Directory to save PDFs
            progress_callback: Optional callback(current, total, filename)
        
        Returns:
            List of (filename, success, pdf_path_or_error)
        """
        results = []
        total = len(docx_files)
        
        for i, docx_path in enumerate(docx_files):
            filename = Path(docx_path).stem
            
            if progress_callback:
                progress_callback(i + 1, total, filename)
            
            success, result, _ = self.convert(docx_path, output_dir)
            results.append((filename, success, result))
        
        return results
    
    def convert_batch_single_call(
        self,
        docx_files: List[str],
        output_dir: str
    ) -> dict:
        """
        Convert ALL documents in a SINGLE LibreOffice call.
        Much faster than converting one at a time!
        
        Args:
            docx_files: List of paths to .docx files
            output_dir: Directory to save PDFs
        
        Returns:
            Dict with success_count, failure_count, errors
        """
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        result = {
            "success_count": 0,
            "failure_count": 0,
            "errors": []
        }
        
        if not docx_files:
            return result
        
        # Build single command with ALL files
        cmd = [
            self.libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_path)
        ] + [str(f) for f in docx_files]
        
        try:
            # Single call converts ALL files!
            subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.timeout_seconds * len(docx_files),  # Scale timeout
                cwd=str(output_path)
            )
            
            # Check which files were created
            for docx_path in docx_files:
                expected_pdf = output_path / (Path(docx_path).stem + ".pdf")
                if expected_pdf.exists():
                    result["success_count"] += 1
                else:
                    result["failure_count"] += 1
                    result["errors"].append(f"Failed to convert: {Path(docx_path).name}")
        
        except subprocess.TimeoutExpired:
            result["errors"].append(f"Batch conversion timed out")
            result["failure_count"] = len(docx_files)
        except Exception as e:
            result["errors"].append(f"Batch conversion error: {str(e)}")
            result["failure_count"] = len(docx_files)
        
        return result
    
    def is_available(self) -> bool:
        """Check if LibreOffice is available and working."""
        if not self.libreoffice_path:
            return False
        
        try:
            result = subprocess.run(
                [self.libreoffice_path, "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            return result.returncode == 0
        except Exception:
            return False
    
    def get_version(self) -> Optional[str]:
        """Get LibreOffice version string."""
        try:
            result = subprocess.run(
                [self.libreoffice_path, "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                return result.stdout.strip()
        except Exception:
            pass
        return None

