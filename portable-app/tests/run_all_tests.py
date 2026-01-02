"""
Run all tests for Report Card Generator
"""

import sys
import subprocess
from pathlib import Path


def run_test(test_file: str) -> bool:
    """Run a single test file."""
    print(f"\n{'='*60}")
    print(f"Running: {test_file}")
    print('='*60)
    
    result = subprocess.run(
        [sys.executable, test_file],
        cwd=str(Path(__file__).parent)
    )
    
    return result.returncode == 0


def main():
    test_dir = Path(__file__).parent
    
    tests = [
        test_dir / "test_data_processor.py",
        test_dir / "test_template_filler.py",
    ]
    
    passed = 0
    failed = 0
    
    for test in tests:
        if test.exists():
            if run_test(str(test)):
                passed += 1
            else:
                failed += 1
        else:
            print(f"Warning: Test file not found: {test}")
    
    print("\n" + "="*60)
    print(f"SUMMARY: {passed} passed, {failed} failed")
    print("="*60)
    
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())

