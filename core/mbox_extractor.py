"""
MBOX Extractor Module

Extracts emails from MBOX files and converts them to EML format.
MBOX is a standard format for storing multiple emails in a single file.
"""

import os
import mailbox
import logging
from pathlib import Path
from typing import Optional, Callable, List, Tuple
from dataclasses import dataclass, field
from email import policy
from email.generator import BytesGenerator
import io

logger = logging.getLogger(__name__)


@dataclass
class MboxExtractionResult:
    """Result of MBOX extraction."""
    success: bool
    email_count: int  # Number of emails extracted
    output_dir: str
    extracted_files: List[str] = field(default_factory=list)  # Paths to extracted EML files
    folder_structure: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    @property
    def emails_extracted(self) -> int:
        """Alias for email_count for backwards compatibility."""
        return self.email_count


class MBOXExtractor:
    """
    Extracts emails from MBOX files.
    
    MBOX is a common format for storing email messages, used by:
    - Thunderbird
    - Apple Mail
    - Gmail exports (Google Takeout)
    - Many other email clients
    """
    
    def __init__(
        self,
        progress_callback: Optional[Callable[[str, str], None]] = None
    ):
        """
        Initialize the MBOX extractor.
        
        Args:
            progress_callback: Optional callback(current_file, message) for progress updates
        """
        self.progress_callback = progress_callback
    
    def _report_progress(self, current_file: str, message: str):
        """Report progress to callback if set."""
        if self.progress_callback:
            self.progress_callback(current_file, message)
    
    def get_mbox_info(self, mbox_path: str) -> Tuple[int, List[str]]:
        """
        Get information about an MBOX file without extracting.
        
        Args:
            mbox_path: Path to MBOX file
            
        Returns:
            Tuple of (email_count, folder_list)
        """
        try:
            mbox = mailbox.mbox(mbox_path)
            email_count = len(mbox)
            # MBOX files don't have folder structure, but we use the filename as folder
            folder_name = Path(mbox_path).stem
            mbox.close()
            return email_count, [folder_name]
        except Exception as e:
            logger.error(f"Failed to read MBOX info: {e}")
            return 0, []
    
    def extract(
        self,
        mbox_path: str,
        output_dir: str,
        preserve_structure: bool = True,
        rename_emls: bool = True
    ) -> MboxExtractionResult:
        """
        Extract emails from MBOX file to EML files.
        
        Args:
            mbox_path: Path to the MBOX file
            output_dir: Directory to extract emails to
            preserve_structure: If True, create folder based on MBOX filename
            rename_emls: If True, rename EMLs to YYYYMMDD_HHMMSS_subject.eml
            
        Returns:
            MboxExtractionResult with extraction details
        """
        result = MboxExtractionResult(
            success=False,
            email_count=0,
            output_dir=output_dir
        )
        
        if not os.path.isfile(mbox_path):
            result.errors.append(f"MBOX file not found: {mbox_path}")
            return result
        
        try:
            # Create output directory
            output_path = Path(output_dir)
            
            # Use MBOX filename as folder name if preserving structure
            if preserve_structure:
                folder_name = Path(mbox_path).stem
                output_path = output_path / folder_name
                result.folder_structure.append(folder_name)
            
            output_path.mkdir(parents=True, exist_ok=True)
            
            # Open MBOX file
            self._report_progress(mbox_path, "Opening MBOX file...")
            mbox = mailbox.mbox(mbox_path)
            
            total_emails = len(mbox)
            logger.info(f"Found {total_emails} emails in MBOX: {mbox_path}")
            self._report_progress(mbox_path, f"Found {total_emails} emails")
            
            # Extract each email
            for i, message in enumerate(mbox):
                try:
                    # Generate filename
                    if rename_emls:
                        filename = self._generate_eml_filename(message, i)
                    else:
                        filename = f"email_{i:06d}.eml"
                    
                    eml_path = output_path / filename
                    
                    # Avoid overwriting existing files
                    if eml_path.exists():
                        base = eml_path.stem
                        suffix = 1
                        while eml_path.exists():
                            eml_path = output_path / f"{base}_{suffix}.eml"
                            suffix += 1
                    
                    # Write email to EML file
                    with open(eml_path, 'wb') as f:
                        gen = BytesGenerator(f, policy=policy.default)
                        gen.flatten(message)
                    
                    result.email_count += 1
                    result.extracted_files.append(str(eml_path))
                    
                    if (i + 1) % 100 == 0:
                        self._report_progress(
                            mbox_path, 
                            f"Extracted {i + 1}/{total_emails} emails"
                        )
                        
                except Exception as e:
                    error_msg = f"Failed to extract email {i}: {e}"
                    logger.warning(error_msg)
                    result.warnings.append(error_msg)
            
            mbox.close()
            
            result.success = True
            logger.info(
                f"Successfully extracted {result.email_count} emails "
                f"from {mbox_path}"
            )
            
        except Exception as e:
            error_msg = f"MBOX extraction failed: {e}"
            logger.error(error_msg)
            result.errors.append(error_msg)
        
        return result
    
    def _generate_eml_filename(self, message, index: int) -> str:
        """
        Generate a descriptive filename for an email.
        
        Format: YYYYMMDD_HHMMSS_subject.eml
        """
        from email.utils import parsedate_to_datetime
        import re
        
        # Get date
        date_str = message.get('Date', '')
        try:
            dt = parsedate_to_datetime(date_str)
            date_prefix = dt.strftime('%Y%m%d_%H%M%S')
        except Exception:
            date_prefix = f"00000000_{index:06d}"
        
        # Get subject
        subject = message.get('Subject', '') or 'No_Subject'
        
        # Decode subject if needed
        from email.header import decode_header
        try:
            decoded_parts = decode_header(subject)
            subject = ''
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    subject += part.decode(encoding or 'utf-8', errors='replace')
                else:
                    subject += part
        except Exception:
            pass
        
        # Sanitize subject for filename
        subject = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', subject)
        subject = subject.strip()[:50]  # Limit length
        
        if not subject:
            subject = 'No_Subject'
        
        return f"{date_prefix}_{subject}.eml"


def extract_mbox_to_emls(
    mbox_path: str,
    output_dir: str,
    progress_callback: Optional[Callable[[str, str], None]] = None
) -> MboxExtractionResult:
    """
    Convenience function to extract MBOX file.
    
    Args:
        mbox_path: Path to MBOX file
        output_dir: Output directory for EML files
        progress_callback: Optional progress callback
        
    Returns:
        MboxExtractionResult
    """
    extractor = MBOXExtractor(progress_callback=progress_callback)
    return extractor.extract(mbox_path, output_dir)
