"""
Data Processor Service
Handles reading Excel files and processing student data.
No hardcoded values - all configuration from settings.
"""

import json
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Optional, Generator


def column_letter_to_index(column_letter: str) -> int:
    """
    Convert Excel column letter (A, B, AA, etc.) to 0-based index.
    
    Args:
        column_letter: Column letter like 'A', 'B', 'AA'
    
    Returns:
        0-based column index
    """
    index = 0
    for char in column_letter.upper():
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1  # Convert to 0-based


class DataProcessor:
    """
    Processes Excel data for report card generation.
    Configurable header rows, null handling, etc.
    """
    
    def __init__(
        self,
        excel_path: str,
        mapping_path: str,
        header_rows_to_skip: int = 4,
        null_indicators: Optional[List[str]] = None,
        default_null_value: str = "---"
    ):
        """
        Initialize the data processor.
        
        Args:
            excel_path: Path to Excel file
            mapping_path: Path to JSON mapping file
            header_rows_to_skip: Number of header rows to skip (configurable)
            null_indicators: List of values to treat as null
            default_null_value: Value to replace nulls with
        """
        self.excel_path = Path(excel_path)
        self.mapping_path = Path(mapping_path)
        self.header_rows_to_skip = header_rows_to_skip
        self.null_indicators = [v.upper() for v in (null_indicators or ["NAN", "NONE", "NA", "NULL", ""])]
        self.default_null_value = default_null_value
        
        # Load data
        self.column_map = self._load_mapping()
        self.df = self._load_dataframe()
    
    def _load_mapping(self) -> Dict[str, str]:
        """Load column mapping from JSON file."""
        with open(self.mapping_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _load_dataframe(self) -> pd.DataFrame:
        """
        Load Excel file and skip header rows.
        Uses header=None to read raw data without interpreting first row as headers.
        Automatically finds the first non-empty sheet if multiple sheets exist.
        """
        # Get all sheet names
        xl = pd.ExcelFile(self.excel_path)
        sheet_names = xl.sheet_names
        
        # Find first non-empty sheet
        df = pd.DataFrame()
        for sheet_name in sheet_names:
            temp_df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)
            if not temp_df.empty and temp_df.shape[0] > 0 and temp_df.shape[1] > 0:
                df = temp_df
                break
        
        # Skip header rows
        if self.header_rows_to_skip > 0 and not df.empty:
            df = df.iloc[self.header_rows_to_skip:]
            df = df.reset_index(drop=True)
        
        return df
    
    def get_total_students(self) -> int:
        """Get total number of student rows."""
        return len(self.df)
    
    def get_student_rows(self) -> Generator[pd.Series, None, None]:
        """
        Yield each student row as a pandas Series.
        
        Yields:
            Student data as pandas Series
        """
        for _, row in self.df.iterrows():
            # Skip completely empty rows
            if row.notna().any():
                yield row
    
    def process_student_data(self, row: pd.Series, class_name: str) -> Dict[str, str]:
        """
        Process a single student row into a dictionary for template filling.
        
        Args:
            row: Pandas Series containing student data
            class_name: Class name to include in output
        
        Returns:
            Dictionary with placeholder keys and values
        """
        field_dict = {}
        
        # Process each field from mapping
        for placeholder_key, column_letter in self.column_map.items():
            column_index = column_letter_to_index(column_letter)
            
            try:
                raw_value = row.iloc[column_index]
            except IndexError:
                raw_value = None
            
            # Handle null/empty values
            field_dict[placeholder_key] = self._clean_value(raw_value)
        
        # Add class name (replace underscores with spaces)
        field_dict['class'] = class_name.replace('_', ' ')
        
        return field_dict
    
    def _clean_value(self, value: Any) -> str:
        """
        Clean a cell value, handling nulls and formatting.
        
        Args:
            value: Raw cell value
        
        Returns:
            Cleaned string value
        """
        if value is None:
            return self.default_null_value
        
        if pd.isna(value):
            return self.default_null_value
        
        str_value = str(value).strip()
        
        if str_value.upper().replace(' ', '') in self.null_indicators:
            return self.default_null_value
        
        return str_value
    
    def validate(self) -> List[str]:
        """
        Validate the data and mapping configuration.
        
        Returns:
            List of validation error messages (empty if valid)
        """
        errors = []
        
        # Check if Excel file exists
        if not self.excel_path.exists():
            errors.append(f"Excel file not found: {self.excel_path}")
        
        # Check if mapping file exists
        if not self.mapping_path.exists():
            errors.append(f"Mapping file not found: {self.mapping_path}")
        
        # Check if dataframe has data
        if self.df.empty:
            errors.append("Excel file has no data after skipping header rows")
        
        # Check if mapping is not empty
        if not self.column_map:
            errors.append("Mapping file is empty")
        
        # Validate column letters in mapping
        max_col = len(self.df.columns) if not self.df.empty else 0
        for key, col_letter in self.column_map.items():
            col_index = column_letter_to_index(col_letter)
            if col_index >= max_col:
                errors.append(f"Column '{col_letter}' for '{key}' exceeds data columns ({max_col} columns available)")
        
        return errors

