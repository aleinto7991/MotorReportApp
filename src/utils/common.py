import re
import os
import sys
import subprocess
import logging
from pathlib import Path
from typing import Optional, Union

logger = logging.getLogger(__name__)

EXCEL_MAX_SHEET_NAME_LENGTH = 31

# SAP Code normalization settings
SAP_CODE_NORMALIZE_TO_UPPER = True  # Set to False if you want case-sensitive SAP codes

def normalize_sap_code(sap_code: str) -> str:
    """
    Normalize SAP code for consistent handling throughout the application.
    
    Args:
        sap_code: The SAP code to normalize
        
    Returns:
        Normalized SAP code (uppercase and stripped)
    """
    if not sap_code:
        return ""
    
    normalized = str(sap_code).strip()
    if SAP_CODE_NORMALIZE_TO_UPPER:
        normalized = normalized.upper()
    
    return normalized

def validate_file_path(file_path: Union[str, Path]) -> Optional[Path]:
    """
    Validate and sanitize file path to prevent security issues.
    
    Args:
        file_path: Path to validate
        
    Returns:
        Validated Path object or None if invalid
    """
    if not file_path:
        return None
        
    try:
        # Convert to Path object and resolve
        path_obj = Path(str(file_path)).resolve()
        
        # Check for path traversal attempts
        if '..' in str(path_obj):
            logger.warning(f"Path traversal detected in: {file_path}")
            return None
            
        # Check if path is too long (Windows limit is 260 chars)
        if len(str(path_obj)) > 250:
            logger.warning(f"Path too long: {file_path}")
            return None
            
        # Check for invalid characters
        invalid_chars = '<>"|?*'
        if any(char in str(path_obj) for char in invalid_chars):
            logger.warning(f"Invalid characters in path: {file_path}")
            return None
            
        return path_obj
        
    except (ValueError, OSError) as e:
        logger.warning(f"Invalid path: {file_path}, error: {e}")
        return None

def validate_directory_path(dir_path: Union[str, Path]) -> Optional[Path]:
    """
    Validate directory path and ensure it exists or can be created.
    
    Args:
        dir_path: Directory path to validate
        
    Returns:
        Validated Path object or None if invalid
    """
    validated_path = validate_file_path(dir_path)
    if not validated_path:
        return None
        
    try:
        # Ensure directory exists
        validated_path.mkdir(parents=True, exist_ok=True)
        return validated_path
        
    except (PermissionError, OSError) as e:
        logger.error(f"Cannot create directory {validated_path}: {e}")
        return None

def sanitize_filename(filename: str, max_length: int = 255) -> str:
    """
    Sanitize filename to remove invalid characters and limit length.
    
    Args:
        filename: Original filename
        max_length: Maximum allowed length
        
    Returns:
        Sanitized filename
    """
    if not filename:
        return "untitled"
        
    # Remove invalid characters
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', str(filename))
    
    # Remove control characters
    sanitized = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', sanitized)
    
    # Limit length
    if len(sanitized) > max_length:
        name, ext = os.path.splitext(sanitized)
        available_length = max_length - len(ext)
        sanitized = name[:available_length] + ext
        
    # Ensure not empty
    return sanitized if sanitized.strip() else "untitled"

def sanitize_sheet_name(name: str) -> str:
    """Sanitize Excel sheet names with enhanced validation."""
    if not name:
        return "Sheet1"
    sanitized = re.sub(r'[:\\/?*\[\]]', '_', str(name))
    sanitized = re.sub(r'\s+', '_', sanitized.strip())
    if sanitized and sanitized[0].isdigit():
        sanitized = f"S_{sanitized}"
    sanitized = sanitized[:EXCEL_MAX_SHEET_NAME_LENGTH]
    return sanitized if sanitized else "Sheet1"

def open_file_externally(file_path: str) -> bool:
    """Open a file with the default application (cross-platform)."""
    try:
        file_path_obj = Path(file_path).resolve()
        if not file_path_obj.exists():
            logger.error(f"File does not exist: {file_path_obj}")
            return False
        if os.name == 'nt':
            os.startfile(str(file_path_obj))
        elif os.name == 'posix':
            subprocess.call(['open' if sys.platform == 'darwin' else 'xdg-open', str(file_path_obj)])
        else:
            logger.warning(f"Unsupported OS for auto-opening files: {os.name}")
            return False
        logger.info(f"Opened file: {file_path_obj}")
        return True
    except Exception as e:
        logger.error(f"Error opening file {file_path}: {e}")
        return False
