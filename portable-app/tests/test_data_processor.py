"""
Tests for Data Processor Service
"""

import sys
import os
import json
import tempfile
from pathlib import Path

# Add app to path
sys.path.insert(0, str(Path(__file__).parent.parent / "app"))

import pandas as pd
from services.data_processor import DataProcessor, column_letter_to_index


def test_column_letter_to_index():
    """Test column letter to index conversion."""
    print("Testing column_letter_to_index...")
    
    assert column_letter_to_index("A") == 0, "A should be 0"
    assert column_letter_to_index("B") == 1, "B should be 1"
    assert column_letter_to_index("Z") == 25, "Z should be 25"
    assert column_letter_to_index("AA") == 26, "AA should be 26"
    assert column_letter_to_index("AB") == 27, "AB should be 27"
    assert column_letter_to_index("AZ") == 51, "AZ should be 51"
    
    # Case insensitive
    assert column_letter_to_index("a") == 0, "a should be 0"
    assert column_letter_to_index("aa") == 26, "aa should be 26"
    
    print("  [OK] All column_letter_to_index tests passed!")


def test_data_processor_with_test_data():
    """Test DataProcessor with realistic test data."""
    print("Testing DataProcessor...")
    
    # Create temp directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        
        # Create test Excel file
        excel_path = temp_dir / "test_students.xlsx"
        df = pd.DataFrame({
            0: ["School Name", "Class Info", "Date", "Roll No", "1", "2", "3"],
            1: ["", "Class Teacher", "", "Name", "Rahul", "Priya", "Amit"],
            2: ["", "", "", "English", "85", "92", "45"],
            3: ["", "", "", "Math", "78", "88", "52"],
            4: ["", "", "", "Result", "Pass", "Pass", "Fail"]
        })
        df.to_excel(excel_path, index=False, header=False)
        
        # Create test mapping
        mapping_path = temp_dir / "mapping.json"
        mapping = {
            "rollno": "A",
            "name": "B",
            "english": "C",
            "math": "D",
            "result": "E"
        }
        with open(mapping_path, 'w') as f:
            json.dump(mapping, f)
        
        # Test DataProcessor
        processor = DataProcessor(
            excel_path=str(excel_path),
            mapping_path=str(mapping_path),
            header_rows_to_skip=4
        )
        
        # Test validation
        errors = processor.validate()
        assert len(errors) == 0, f"Validation failed: {errors}"
        
        # Test total students
        total = processor.get_total_students()
        assert total == 3, f"Expected 3 students, got {total}"
        
        # Test student data processing
        students = list(processor.get_student_rows())
        assert len(students) == 3, f"Expected 3 student rows, got {len(students)}"
        
        # Test first student
        student1 = processor.process_student_data(students[0], "Class_5A")
        assert student1["name"] == "Rahul", f"Expected 'Rahul', got {student1['name']}"
        assert student1["english"] == "85", f"Expected '85', got {student1['english']}"
        assert student1["class"] == "Class 5A", f"Expected 'Class 5A', got {student1['class']}"
        
        print("  [OK] All DataProcessor tests passed!")


def test_null_handling():
    """Test null value handling."""
    print("Testing null handling...")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        
        # Create Excel with null values
        excel_path = temp_dir / "test_nulls.xlsx"
        df = pd.DataFrame({
            0: ["Header", "1", "2"],
            1: ["Name", "John", ""],
            2: ["Score", "NaN", "NA"]
        })
        df.to_excel(excel_path, index=False, header=False)
        
        mapping_path = temp_dir / "mapping.json"
        mapping = {"rollno": "A", "name": "B", "score": "C"}
        with open(mapping_path, 'w') as f:
            json.dump(mapping, f)
        
        processor = DataProcessor(
            excel_path=str(excel_path),
            mapping_path=str(mapping_path),
            header_rows_to_skip=1,
            default_null_value="---"
        )
        
        students = list(processor.get_student_rows())
        
        # Second student has empty name and NA score
        student2 = processor.process_student_data(students[1], "Test")
        assert student2["name"] == "---", f"Expected '---' for empty, got {student2['name']}"
        assert student2["score"] == "---", f"Expected '---' for NA, got {student2['score']}"
        
        print("  [OK] All null handling tests passed!")


if __name__ == "__main__":
    print("\n" + "="*50)
    print("Running Data Processor Tests")
    print("="*50 + "\n")
    
    try:
        test_column_letter_to_index()
        test_data_processor_with_test_data()
        test_null_handling()
        
        print("\n" + "="*50)
        print("[OK] ALL TESTS PASSED!")
        print("="*50 + "\n")
    except AssertionError as e:
        print(f"\n[FAIL] TEST FAILED: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[FAIL] ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

