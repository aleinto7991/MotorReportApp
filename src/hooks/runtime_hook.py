import sys
import os

# Force xlrd to be available
try:
    import xlrd
    print(f"Runtime hook: xlrd imported successfully from {xlrd.__file__}")
except ImportError as e:
    print(f"Runtime hook: Failed to import xlrd: {e}")
