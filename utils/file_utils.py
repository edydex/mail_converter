"""
File Utilities Module

Common file operations and helpers.
"""

import os
import re
import hashlib
import shutil
from pathlib import Path
from typing import Union, Optional
import logging

logger = logging.getLogger(__name__)


def sanitize_filename(filename: str, max_length: int = 200, replacement: str = '_') -> str:
    """
    Sanitize a filename by removing or replacing invalid characters.
    
    Args:
        filename: Original filename
        max_length: Maximum length for the filename
        replacement: Character to replace invalid chars with
        
    Returns:
        Sanitized filename
    """
    if not filename:
        return "unnamed"
    
    # Characters not allowed in filenames (Windows is most restrictive)
    invalid_chars = '<>:"/\\|?*\x00'
    
    # Also remove control characters
    for i in range(32):
        invalid_chars += chr(i)
    
    # Replace invalid characters
    sanitized = filename
    for char in invalid_chars:
        sanitized = sanitized.replace(char, replacement)
    
    # Replace multiple consecutive replacements with single
    sanitized = re.sub(f'{re.escape(replacement)}+', replacement, sanitized)
    
    # Remove leading/trailing dots, spaces, and replacement chars
    sanitized = sanitized.strip(f'. {replacement}')
    
    # Limit length while preserving extension
    if len(sanitized) > max_length:
        name, ext = os.path.splitext(sanitized)
        max_name_length = max_length - len(ext)
        sanitized = name[:max_name_length] + ext
    
    # Handle empty result
    if not sanitized:
        return "unnamed"
    
    # Handle reserved Windows names
    reserved_names = {
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    }
    
    name_without_ext = os.path.splitext(sanitized)[0].upper()
    if name_without_ext in reserved_names:
        sanitized = f"_{sanitized}"
    
    return sanitized


def ensure_dir(path: Union[str, Path]) -> Path:
    """
    Ensure a directory exists, creating it if necessary.
    
    Args:
        path: Directory path
        
    Returns:
        Path object for the directory
    """
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_file_hash(filepath: Union[str, Path], algorithm: str = 'md5') -> str:
    """
    Calculate hash of a file.
    
    Args:
        filepath: Path to the file
        algorithm: Hash algorithm ('md5', 'sha1', 'sha256')
        
    Returns:
        Hex digest of the file hash
    """
    hash_func = hashlib.new(algorithm)
    
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            hash_func.update(chunk)
    
    return hash_func.hexdigest()


def get_unique_filepath(filepath: Union[str, Path]) -> Path:
    """
    Get a unique filepath by appending a number if file exists.
    
    Args:
        filepath: Desired filepath
        
    Returns:
        Unique filepath that doesn't exist
    """
    filepath = Path(filepath)
    
    if not filepath.exists():
        return filepath
    
    counter = 1
    stem = filepath.stem
    suffix = filepath.suffix
    parent = filepath.parent
    
    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def safe_copy(src: Union[str, Path], dst: Union[str, Path], overwrite: bool = False) -> Path:
    """
    Safely copy a file with optional overwrite control.
    
    Args:
        src: Source file path
        dst: Destination file path
        overwrite: Whether to overwrite if destination exists
        
    Returns:
        Path to the copied file
    """
    src = Path(src)
    dst = Path(dst)
    
    if not src.exists():
        raise FileNotFoundError(f"Source file not found: {src}")
    
    dst.parent.mkdir(parents=True, exist_ok=True)
    
    if dst.exists() and not overwrite:
        dst = get_unique_filepath(dst)
    
    shutil.copy2(src, dst)
    return dst


def get_file_size_str(size_bytes: int) -> str:
    """
    Convert file size in bytes to human-readable string.
    
    Args:
        size_bytes: Size in bytes
        
    Returns:
        Human-readable size string
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    
    return f"{size_bytes:.1f} PB"


def clean_directory(directory: Union[str, Path], pattern: str = "*") -> int:
    """
    Remove files matching a pattern from a directory.
    
    Args:
        directory: Directory to clean
        pattern: Glob pattern for files to remove
        
    Returns:
        Number of files removed
    """
    directory = Path(directory)
    count = 0
    
    if not directory.exists():
        return 0
    
    for filepath in directory.glob(pattern):
        if filepath.is_file():
            try:
                filepath.unlink()
                count += 1
            except Exception as e:
                logger.warning(f"Could not remove {filepath}: {e}")
    
    return count


def get_extension(filepath: Union[str, Path], include_dot: bool = True) -> str:
    """
    Get the file extension, handling compound extensions.
    
    Args:
        filepath: File path
        include_dot: Whether to include the leading dot
        
    Returns:
        File extension
    """
    filepath = Path(filepath)
    ext = filepath.suffix.lower()
    
    # Handle compound extensions like .tar.gz
    if ext in ('.gz', '.bz2', '.xz'):
        stem_ext = Path(filepath.stem).suffix.lower()
        if stem_ext:
            ext = stem_ext + ext
    
    if not include_dot and ext.startswith('.'):
        ext = ext[1:]
    
    return ext
