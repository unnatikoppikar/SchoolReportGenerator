"""
Template Filler Service
Handles filling Word templates with student data using docxtpl.
"""

import os
import re
from pathlib import Path
from typing import Dict, List
from docxtpl import DocxTemplate


class TemplateFiller:
    """
    Fills Word document templates with placeholder data.
    Uses docxtpl for robust template handling.
    """
    
    def __init__(
        self,
        template_path: str,
        placeholder_prefix: str = "{{",
        placeholder_suffix: str = "}}"
    ):
        """
        Initialize the template filler.
        
        Args:
            template_path: Path to the Word template file
            placeholder_prefix: Prefix for placeholders (default: "{{")
            placeholder_suffix: Suffix for placeholders (default: "}}")
        """
        self.template_path = Path(template_path)
        self.placeholder_prefix = placeholder_prefix
        self.placeholder_suffix = placeholder_suffix
        
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template file not found: {self.template_path}")
    
    def fill_template(self, data: Dict[str, str], output_path: str) -> str:
        """
        Fill the template with data and save to output path.
        
        Args:
            data: Dictionary with placeholder keys and values
            output_path: Path to save the filled document
        
        Returns:
            Path to the saved document
        """
        # Load template fresh for each document
        doc = DocxTemplate(str(self.template_path))
        
        # Render with data
        doc.render(data)
        
        # Ensure output directory exists
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Save filled document
        doc.save(str(output_path))
        
        return str(output_path)
    
    def get_placeholders(self) -> List[str]:
        """
        Extract all placeholders from the template.
        Useful for validation.
        
        Returns:
            List of placeholder names found in template
        """
        doc = DocxTemplate(str(self.template_path))
        
        # Get all variables from the template
        # docxtpl uses Jinja2 syntax, so we look for {{ variable }}
        placeholders = set()
        
        # Get the undeclared variables
        try:
            variables = doc.get_undeclared_template_variables()
            placeholders.update(variables)
        except Exception:
            pass
        
        return list(placeholders)
    
    def validate_data(self, data: Dict[str, str]) -> List[str]:
        """
        Validate that data contains all required placeholders.
        
        Args:
            data: Dictionary with placeholder data
        
        Returns:
            List of missing placeholder names
        """
        template_placeholders = self.get_placeholders()
        missing = []
        
        for placeholder in template_placeholders:
            if placeholder not in data:
                missing.append(placeholder)
        
        return missing


def sanitize_filename(filename: str) -> str:
    """
    Sanitize a filename by removing invalid characters.
    
    Args:
        filename: Original filename
    
    Returns:
        Sanitized filename safe for all operating systems
    """
    # Characters not allowed in filenames on Windows
    invalid_chars = r'[<>:"/\\|?*]'
    
    # Replace invalid characters with underscore
    sanitized = re.sub(invalid_chars, '_', filename)
    
    # Remove leading/trailing whitespace and dots
    sanitized = sanitized.strip(' .')
    
    # Limit length (Windows max path component is 255)
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    
    # Fallback if empty
    if not sanitized:
        sanitized = "unnamed"
    
    return sanitized

