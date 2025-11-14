"""Diagnostic tool to check TEST_LAB directory structure and file matching.

This script scans the TEST_LAB/CARICHI NOMINALI directory and:
1. Lists all year subdirectories found
2. Shows all Excel files in each directory
3. Tests which test numbers would match each file using the loader's normalization rules
4. Helps diagnose why specific test numbers aren't being found

Usage:
    python -m src.tools.check_testlab_files [test_number1] [test_number2] ...

Example:
    python -m src.tools.check_testlab_files 26178A 26529A
"""
from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, Optional

from src import directory_config


def normalize_test_number(test_number: str) -> str:
    """Apply the same normalization logic as TestLabSummaryLoader."""
    return re.sub(r"[^0-9A-Za-z]+", "", test_number or "").upper()


def get_test_lab_directory() -> Optional[Path]:
    """Get the configured TEST_LAB directory."""
    test_lab_dir = getattr(directory_config, "TEST_LAB_CARICHI_DIR", None)
    if not test_lab_dir:
        return None
    path = Path(test_lab_dir)
    return path if path.exists() else None


def scan_directory(directory: Path) -> List[Path]:
    """Find all Excel files in a directory."""
    if not directory.is_dir():
        return []
    return [
        f
        for f in directory.iterdir()
        if f.is_file() and f.suffix.lower() in {".xlsx", ".xls"}
    ]


def would_match(file_stem: str, test_number: str, allow_prefix: bool = False) -> bool:
    """Check if a file would match the given test number using loader logic."""
    file_stem_lower = file_stem.lower()
    test_lower = test_number.lower()
    
    # Exact match
    if file_stem_lower == test_lower:
        return True
    
    # Prefix match (if allowed)
    if allow_prefix and file_stem_lower.startswith(test_lower):
        return True
    
    return False


def main():
    """Run the diagnostic scan."""
    print("=" * 80)
    print("TEST_LAB Directory Diagnostic Tool")
    print("=" * 80)
    print()
    
    # Get the TEST_LAB directory
    test_lab_dir = get_test_lab_directory()
    if not test_lab_dir:
        print("❌ ERROR: TEST_LAB directory not configured or does not exist")
        print(f"   Configured path: {getattr(directory_config, 'TEST_LAB_CARICHI_DIR', 'NOT SET')}")
        return 1
    
    print(f"✓ TEST_LAB Directory: {test_lab_dir}")
    print()
    
    # Get test numbers from command line
    test_numbers = sys.argv[1:] if len(sys.argv) > 1 else []
    if test_numbers:
        print(f"Searching for test numbers: {', '.join(test_numbers)}")
        normalized = [normalize_test_number(t) for t in test_numbers]
        print(f"Normalized forms: {', '.join(normalized)}")
        print()
    
    # Scan subdirectories
    subdirs = [p for p in test_lab_dir.iterdir() if p.is_dir()]
    subdirs_sorted = sorted(subdirs, key=lambda p: p.name, reverse=True)
    
    print(f"Found {len(subdirs_sorted)} subdirectories:")
    for subdir in subdirs_sorted:
        print(f"  - {subdir.name}")
    print()
    
    # Scan each directory for Excel files
    all_directories = subdirs_sorted + [test_lab_dir]
    total_files = 0
    matches_found = {}
    
    for directory in all_directories:
        excel_files = scan_directory(directory)
        total_files += len(excel_files)
        
        if not excel_files:
            print(f"📁 {directory.name}/")
            print(f"   No Excel files found")
            print()
            continue
        
        print(f"📁 {directory.name}/")
        print(f"   Found {len(excel_files)} Excel file(s):")
        
        for file in sorted(excel_files, key=lambda f: f.name):
            file_stem = file.stem
            print(f"      • {file.name}")
            
            # Check if this file matches any requested test numbers
            if test_numbers:
                for test_num in test_numbers:
                    norm_test = normalize_test_number(test_num)
                    norm_stem = normalize_test_number(file_stem)
                    
                    exact_match = would_match(norm_stem, norm_test, allow_prefix=False)
                    prefix_match = would_match(norm_stem, norm_test, allow_prefix=True)
                    
                    if exact_match:
                        print(f"        ✓ EXACT MATCH for {test_num} (normalized: {norm_test})")
                        matches_found.setdefault(test_num, []).append((str(file), "exact"))
                    elif prefix_match:
                        print(f"        ≈ PREFIX MATCH for {test_num} (normalized: {norm_test})")
                        matches_found.setdefault(test_num, []).append((str(file), "prefix"))
        
        print()
    
    # Summary
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"Total directories scanned: {len(all_directories)}")
    print(f"Total Excel files found: {total_files}")
    print()
    
    if test_numbers:
        print("Match Results:")
        for test_num in test_numbers:
            norm_test = normalize_test_number(test_num)
            print(f"  {test_num} (normalized: {norm_test}):")
            
            if test_num in matches_found:
                for file_path, match_type in matches_found[test_num]:
                    print(f"    ✓ {match_type.upper()}: {file_path}")
            else:
                print(f"    ❌ NOT FOUND")
                
                # Suggest potential fallback candidates
                is_alias = norm_test.endswith("A")
                if is_alias:
                    fallback = norm_test.rstrip("A")
                    print(f"       Loader would also search for fallback: {fallback}")
                else:
                    fallback = f"{norm_test}A"
                    print(f"       Loader would also search for fallback: {fallback}")
        print()
    
    if total_files == 0:
        print("⚠️  WARNING: No Excel files found in any directory!")
        print("   Possible reasons:")
        print("   - Files are stored in a different location")
        print("   - Files have different extensions (not .xlsx or .xls)")
        print("   - Directory permissions prevent reading")
        print()
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
