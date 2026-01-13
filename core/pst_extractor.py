"""
PST Extractor Module

Extracts EML files from PST using libpst (readpst command).
Handles large PST files with progress tracking.
"""

import os
import sys
import subprocess
import shutil
import platform
import logging
from pathlib import Path
from typing import Optional, Callable, List, Tuple
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class ExtractionResult:
    """Result of PST extraction"""
    success: bool
    eml_count: int
    output_dir: Path
    errors: List[str]
    warnings: List[str]


class PSTExtractor:
    """
    Extracts emails from PST files using libpst/readpst.
    
    Supports progress tracking for large files.
    """
    
    def __init__(self, progress_callback: Optional[Callable[[int, int, str], None]] = None):
        """
        Initialize PST Extractor.
        
        Args:
            progress_callback: Optional callback function(current, total, message)
        """
        self.progress_callback = progress_callback
        self.readpst_path = self._find_readpst()
        
    def _find_readpst(self) -> Optional[str]:
        """Find readpst executable on the system."""
        system = platform.system()
        
        # Check if running as PyInstaller bundle first
        if hasattr(sys, '_MEIPASS'):
            # Running as compiled executable - look for bundled readpst
            if system == "Windows":
                bundled_path = os.path.join(sys._MEIPASS, 'bin', 'readpst.exe')
            else:
                bundled_path = os.path.join(sys._MEIPASS, 'bin', 'readpst')
            
            if os.path.isfile(bundled_path):
                logger.info(f"Using bundled readpst at: {bundled_path}")
                return bundled_path
            else:
                logger.warning(f"Bundled readpst not found at: {bundled_path}, falling back to system")
        
        # Fall back to system installation
        possible_paths = []
        
        if system == "Darwin":  # macOS
            possible_paths = [
                "/opt/homebrew/bin/readpst",  # Apple Silicon
                "/usr/local/bin/readpst",     # Intel
                shutil.which("readpst")
            ]
        elif system == "Windows":
            # Check relative to script location (for development)
            script_dir = os.path.dirname(os.path.abspath(__file__))
            possible_paths = [
                os.path.join(script_dir, "..", "build", "bin", "readpst.exe"),
                os.path.join(script_dir, "..", "bin", "readpst.exe"),
                shutil.which("readpst.exe"),
                shutil.which("readpst")
            ]
        else:  # Linux
            possible_paths = [
                "/usr/bin/readpst",
                "/usr/local/bin/readpst",
                shutil.which("readpst")
            ]
        
        for path in possible_paths:
            if path and os.path.isfile(path):
                logger.info(f"Found readpst at: {path}")
                return path
                
        logger.warning("readpst not found in common locations")
        return shutil.which("readpst")  # Last resort
    
    def is_available(self) -> bool:
        """Check if readpst is available on the system."""
        return self.readpst_path is not None
    
    def get_pst_info(self, pst_path: str) -> Tuple[int, List[str]]:
        """
        Get information about a PST file (folder count, structure).
        
        Args:
            pst_path: Path to PST file
            
        Returns:
            Tuple of (estimated email count, list of folder names)
        """
        if not self.is_available():
            raise RuntimeError("readpst is not available. Please install libpst.")
        
        # Use readpst -d to get debug info about structure
        try:
            result = subprocess.run(
                [self.readpst_path, "-d", "/dev/null", pst_path],
                capture_output=True,
                text=True,
                timeout=60
            )
            
            # Parse output to estimate email count
            # This is a rough estimate based on folder structure
            folders = []
            email_count = 0
            
            for line in result.stderr.split('\n'):
                if 'folder' in line.lower():
                    folders.append(line.strip())
                if 'email' in line.lower() or 'message' in line.lower():
                    # Try to extract count
                    parts = line.split()
                    for part in parts:
                        if part.isdigit():
                            email_count += int(part)
            
            return email_count if email_count > 0 else 100, folders  # Default estimate
            
        except subprocess.TimeoutExpired:
            logger.warning("Timeout getting PST info, using defaults")
            return 100, []
        except Exception as e:
            logger.warning(f"Error getting PST info: {e}")
            return 100, []
    
    def extract(
        self,
        pst_path: str,
        output_dir: str,
        preserve_structure: bool = True
    ) -> ExtractionResult:
        """
        Extract EML files from a PST file.
        
        Args:
            pst_path: Path to the PST file
            output_dir: Directory to extract EML files to
            preserve_structure: If True, preserve folder structure from PST
            
        Returns:
            ExtractionResult with extraction details
        """
        errors = []
        warnings = []
        
        # Validate inputs
        if not os.path.isfile(pst_path):
            return ExtractionResult(
                success=False,
                eml_count=0,
                output_dir=Path(output_dir),
                errors=[f"PST file not found: {pst_path}"],
                warnings=[]
            )
        
        if not self.is_available():
            return ExtractionResult(
                success=False,
                eml_count=0,
                output_dir=Path(output_dir),
                errors=["readpst is not available. Please install libpst."],
                warnings=[]
            )
        
        # Create output directory
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Report initial progress
        if self.progress_callback:
            self.progress_callback(0, 100, "Starting PST extraction...")
        
        # Build readpst command
        # Using same flags as working outlook_forensics:
        # -e: Extract emails to separate files
        # -o: Output directory
        # -D: Don't create subdirectories (flat output)
        # -M: MH mode - don't extract attachments as separate files
        cmd = [
            self.readpst_path,
            "-e",  # Extract emails
            "-o", str(output_path),  # Output directory
            "-D",  # Don't create subdirectories
            "-M",  # MH mode - attachments stay in email
            pst_path
        ]
        
        logger.info(f"Running: {' '.join(cmd)}")
        
        if self.progress_callback:
            self.progress_callback(10, 100, "Extracting emails from PST...")
        
        try:
            # Run readpst
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            # Monitor output for progress
            stdout, stderr = process.communicate()
            
            if process.returncode != 0:
                errors.append(f"readpst error (code {process.returncode}): {stderr}")
                if "unable to open" in stderr.lower():
                    errors.append("Unable to open PST file. It may be corrupted or password protected.")
            
            # Parse any warnings from stderr
            if stderr:
                for line in stderr.split('\n'):
                    if line.strip() and 'warn' in line.lower():
                        warnings.append(line.strip())
            
        except subprocess.TimeoutExpired:
            errors.append("PST extraction timed out")
            return ExtractionResult(
                success=False,
                eml_count=0,
                output_dir=output_path,
                errors=errors,
                warnings=warnings
            )
        except Exception as e:
            errors.append(f"Extraction error: {str(e)}")
            return ExtractionResult(
                success=False,
                eml_count=0,
                output_dir=output_path,
                errors=errors,
                warnings=warnings
            )
        
        if self.progress_callback:
            self.progress_callback(80, 100, "Counting extracted emails...")
        
        # Count extracted email files
        # readpst creates numbered files (1, 2, 3...) without extensions
        # Files are in nested subdirectories by folder structure
        email_files = []
        for item in output_path.rglob('*'):
            if item.is_file():
                # Only include files that look like emails (numbered or .eml)
                # Skip things like .DS_Store
                if item.name.isdigit() or item.suffix.lower() == '.eml':
                    email_files.append(item)
        
        eml_count = len(email_files)
        
        if eml_count == 0:
            warnings.append("No EML files were extracted. The PST may be empty or in an unsupported format.")
        
        if self.progress_callback:
            self.progress_callback(100, 100, f"Extracted {eml_count} emails")
        
        logger.info(f"Extracted {eml_count} emails to {output_path}")
        
        return ExtractionResult(
            success=len(errors) == 0,
            eml_count=eml_count,
            output_dir=output_path,
            errors=errors,
            warnings=warnings
        )
    
    def get_extracted_emls(self, output_dir: str) -> List[Path]:
        """
        Get list of all extracted email files.
        
        Args:
            output_dir: Directory containing extracted emails
            
        Returns:
            List of Path objects to email files
        """
        output_path = Path(output_dir)
        email_files = []
        
        # readpst creates numbered files in nested subdirectories
        for item in output_path.rglob('*'):
            if item.is_file() and (item.name.isdigit() or item.suffix.lower() == '.eml'):
                email_files.append(item)
        
        return sorted(email_files)
