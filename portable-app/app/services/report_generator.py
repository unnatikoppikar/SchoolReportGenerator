"""
Report Generator Service
Orchestrates the entire report card generation process.
"""

import json
import tempfile
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Callable, Any

from .data_processor import DataProcessor
from .template_filler import TemplateFiller, sanitize_filename
from .pdf_converter import PDFConverter


class ReportGenerator:
    """
    Main orchestrator for generating report cards.
    Coordinates data processing, template filling, and PDF conversion.
    """
    
    def __init__(self, settings: Optional[Dict[str, Any]] = None):
        """
        Initialize the report generator with settings.
        
        Args:
            settings: Configuration dictionary (loaded from settings.json)
        """
        self.settings = settings or self._default_settings()
    
    def _default_settings(self) -> Dict[str, Any]:
        """Default settings if none provided."""
        return {
            "header_rows_to_skip": 4,
            "placeholder_prefix": "{{",
            "placeholder_suffix": "}}",
            "default_null_value": "---",
            "null_indicators": ["NAN", "NONE", "NA", "NULL", ""],
            "libreoffice_timeout_seconds": 60
        }
    
    def generate(
        self,
        excel_path: str,
        template_path: str,
        mapping_path: str,
        class_name: str,
        output_dir: str,
        progress_callback: Optional[Callable[[int, int, str, str], None]] = None
    ) -> Dict[str, Any]:
        """
        Generate report cards for all students.
        
        Args:
            excel_path: Path to Excel file with student data
            template_path: Path to Word template
            mapping_path: Path to JSON mapping file
            class_name: Name of the class
            output_dir: Directory to save generated PDFs
            progress_callback: Optional callback(current, total, student_name, status)
        
        Returns:
            Dictionary with generation results
        """
        results = {
            "success": False,
            "total_students": 0,
            "generated": 0,
            "failed": 0,
            "errors": [],
            "output_dir": output_dir
        }
        
        try:
            # Initialize components
            data_processor = DataProcessor(
                excel_path=excel_path,
                mapping_path=mapping_path,
                header_rows_to_skip=self.settings.get("header_rows_to_skip", 4),
                null_indicators=self.settings.get("null_indicators"),
                default_null_value=self.settings.get("default_null_value", "---")
            )
            
            # Validate data
            validation_errors = data_processor.validate()
            if validation_errors:
                results["errors"] = validation_errors
                return results
            
            template_filler = TemplateFiller(
                template_path=template_path,
                placeholder_prefix=self.settings.get("placeholder_prefix", "{{"),
                placeholder_suffix=self.settings.get("placeholder_suffix", "}}")
            )
            
            pdf_converter = PDFConverter(
                timeout_seconds=self.settings.get("libreoffice_timeout_seconds", 60)
            )
            
            # Check LibreOffice availability
            if not pdf_converter.is_available():
                results["errors"].append("LibreOffice not found or not working")
                return results
            
            # Create temp directory for intermediate docx files
            temp_dir = Path(tempfile.mkdtemp(prefix="reportgen_"))
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)
            
            try:
                # Get all student rows
                student_rows = list(data_processor.get_student_rows())
                total_students = len(student_rows)
                results["total_students"] = total_students
                
                if total_students == 0:
                    results["errors"].append("No student data found in Excel file")
                    return results
                
                # PHASE 1: Fill all templates first (fast)
                docx_files = []
                for i, row in enumerate(student_rows):
                    current = i + 1
                    
                    # Process student data
                    student_data = data_processor.process_student_data(row, class_name)
                    
                    # Get student name for filename
                    student_name = student_data.get("name", f"Student_{current}")
                    safe_name = sanitize_filename(student_name)
                    
                    if progress_callback:
                        progress_callback(current, total_students, student_name, "Filling template...")
                    
                    try:
                        # Fill template
                        docx_path = temp_dir / f"{safe_name}.docx"
                        template_filler.fill_template(student_data, str(docx_path))
                        docx_files.append(str(docx_path))
                    
                    except Exception as e:
                        results["failed"] += 1
                        results["errors"].append(f"{student_name}: {str(e)}")
                
                # PHASE 2: Batch convert all docx to PDF (single LibreOffice call!)
                if docx_files:
                    if progress_callback:
                        progress_callback(total_students, total_students, "All students", "Converting to PDF (batch)...")
                    
                    # Use batch conversion - much faster!
                    batch_results = pdf_converter.convert_batch_single_call(docx_files, str(output_path))
                    
                    results["generated"] = batch_results.get("success_count", 0)
                    results["failed"] += batch_results.get("failure_count", 0)
                    if batch_results.get("errors"):
                        results["errors"].extend(batch_results["errors"])
                
                results["success"] = results["failed"] == 0
                
            finally:
                # Cleanup temp directory
                try:
                    shutil.rmtree(temp_dir)
                except Exception:
                    pass
        
        except Exception as e:
            results["errors"].append(f"Unexpected error: {str(e)}")
        
        return results
    
    def validate_inputs(
        self,
        excel_path: str,
        template_path: str,
        mapping_path: str
    ) -> List[str]:
        """
        Validate input files before generation.
        
        Args:
            excel_path: Path to Excel file
            template_path: Path to Word template
            mapping_path: Path to JSON mapping
        
        Returns:
            List of validation error messages
        """
        errors = []
        
        # Check file existence
        if not Path(excel_path).exists():
            errors.append(f"Excel file not found: {excel_path}")
        
        if not Path(template_path).exists():
            errors.append(f"Template file not found: {template_path}")
        
        if not Path(mapping_path).exists():
            errors.append(f"Mapping file not found: {mapping_path}")
        
        # Check file extensions
        excel_ext = Path(excel_path).suffix.lower()
        if excel_ext not in self.settings.get("allowed_excel_extensions", [".xlsx", ".xls"]):
            errors.append(f"Invalid Excel file extension: {excel_ext}")
        
        template_ext = Path(template_path).suffix.lower()
        if template_ext not in self.settings.get("allowed_template_extensions", [".docx"]):
            errors.append(f"Invalid template file extension: {template_ext}")
        
        mapping_ext = Path(mapping_path).suffix.lower()
        if mapping_ext not in self.settings.get("allowed_mapping_extensions", [".json"]):
            errors.append(f"Invalid mapping file extension: {mapping_ext}")
        
        # Validate mapping JSON
        if Path(mapping_path).exists():
            try:
                with open(mapping_path, 'r') as f:
                    mapping = json.load(f)
                if not isinstance(mapping, dict):
                    errors.append("Mapping file must contain a JSON object")
                elif not mapping:
                    errors.append("Mapping file is empty")
            except json.JSONDecodeError as e:
                errors.append(f"Invalid JSON in mapping file: {e}")
        
        return errors


def load_settings(settings_path: str) -> Dict[str, Any]:
    """
    Load settings from JSON file.
    
    Args:
        settings_path: Path to settings.json
    
    Returns:
        Settings dictionary
    """
    try:
        with open(settings_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

