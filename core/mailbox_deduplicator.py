"""
Mailbox Deduplicator Module

Removes duplicate emails from a single mailbox.
"""

import os
import logging
from pathlib import Path
from typing import List, Optional, Callable
from dataclasses import dataclass, field
import shutil
import tempfile

from .email_fingerprint import (
    FingerprintIndex,
    FingerprintMatch,
    create_fingerprint_from_parsed_email
)
from .mailbox_writer import MailboxWriter, OutputFormat
from .pst_extractor import PSTExtractor
from .mbox_extractor import MBOXExtractor
from .eml_parser import EMLParser

logger = logging.getLogger(__name__)


@dataclass
class DedupeConfig:
    """Configuration for deduplication"""
    # Matching options
    use_message_id: bool = True
    use_content: bool = True
    timestamp_tolerance_seconds: int = 15
    
    # Output options
    output_format: OutputFormat = OutputFormat.MBOX
    keep_duplicates: bool = False  # If True, also output duplicates separately


@dataclass
class DedupeResult:
    """Result of deduplication operation"""
    success: bool
    output_path: str = ""
    duplicates_path: Optional[str] = None
    
    # Counts
    total_emails: int = 0
    unique_emails: int = 0
    duplicates_found: int = 0
    
    # Detailed matches
    duplicate_matches: List[FingerprintMatch] = field(default_factory=list)
    
    # Errors and warnings
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


class MailboxDeduplicator:
    """
    Removes duplicate emails from a mailbox.
    
    Outputs a clean mailbox with duplicates removed,
    and optionally a separate file with the duplicates.
    """
    
    def __init__(
        self,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ):
        """
        Initialize the deduplicator.
        
        Args:
            progress_callback: Optional callback(current, total, message)
        """
        self.progress_callback = progress_callback
        self.pst_extractor = PSTExtractor()
        self.mbox_extractor = MBOXExtractor()
        self.eml_parser = EMLParser()
        self.writer = MailboxWriter(progress_callback)
    
    def _report_progress(self, current: int, total: int, message: str):
        """Report progress to callback."""
        if self.progress_callback:
            self.progress_callback(current, total, message)
    
    def _detect_input_type(self, path: str) -> str:
        """Detect input type."""
        p = Path(path)
        if p.is_dir():
            return "eml_folder"
        ext = p.suffix.lower()
        if ext == ".pst":
            return "pst"
        elif ext == ".mbox":
            return "mbox"
        return "eml_folder"
    
    def _collect_email_files(self, directory: Path) -> List[str]:
        """
        Collect all email files from a directory.
        
        readpst creates numbered files (1, 2, 3...) without extensions.
        Also handles standard .eml files.
        """
        email_files = []
        for item in directory.rglob('*'):
            if item.is_file():
                if item.name.isdigit() or item.suffix.lower() in ['.eml', '.msg', '.email']:
                    email_files.append(str(item))
        return email_files
    
    def _extract_mailbox(
        self,
        input_path: str,
        temp_dir: Path
    ) -> tuple[List[str], List[str]]:
        """Extract emails from mailbox."""
        input_type = self._detect_input_type(input_path)
        eml_paths = []
        warnings = []
        
        output_subdir = temp_dir / "extracted"
        output_subdir.mkdir(parents=True, exist_ok=True)
        
        if input_type == "pst":
            result = self.pst_extractor.extract(
                input_path,
                str(output_subdir),
                preserve_structure=True
            )
            if result.success:
                # readpst creates numbered files without .eml extension
                eml_paths = self._collect_email_files(output_subdir)
            warnings.extend(result.errors + result.warnings)
            
        elif input_type == "mbox":
            result = self.mbox_extractor.extract(
                input_path,
                str(output_subdir),
                preserve_structure=False
            )
            if result.success:
                eml_paths = result.extracted_files
            warnings.extend(result.errors + result.warnings)
            
        elif input_type == "eml_folder":
            input_dir = Path(input_path)
            eml_paths = self._collect_email_files(input_dir)
        
        return eml_paths, warnings
    
    def deduplicate(
        self,
        input_path: str,
        output_path: str,
        config: DedupeConfig = None
    ) -> DedupeResult:
        """
        Remove duplicates from a mailbox.
        
        Args:
            input_path: Path to input mailbox
            output_path: Path for deduplicated output
            config: Deduplication configuration
            
        Returns:
            DedupeResult with details
        """
        if config is None:
            config = DedupeConfig()
        
        result = DedupeResult(success=False, output_path=output_path)
        
        temp_dir = Path(tempfile.mkdtemp(prefix="mailbox_dedupe_"))
        
        try:
            # Step 1: Extract mailbox
            self._report_progress(0, 3, "Extracting mailbox...")
            eml_paths, warnings = self._extract_mailbox(input_path, temp_dir)
            result.warnings.extend(warnings)
            result.total_emails = len(eml_paths)
            
            if not eml_paths:
                result.errors.append("No emails found in mailbox")
                return result
            
            # Step 2: Find duplicates
            self._report_progress(1, 3, "Finding duplicates...")
            
            index = FingerprintIndex(
                timestamp_tolerance_seconds=config.timestamp_tolerance_seconds
            )
            
            unique_paths: List[str] = []
            duplicate_paths: List[str] = []
            
            for i, eml_path in enumerate(eml_paths):
                if i % 100 == 0:
                    self._report_progress(
                        i, result.total_emails,
                        f"Scanning: {i}/{result.total_emails}"
                    )
                
                try:
                    email_data = self.eml_parser.parse_file(eml_path)
                    
                    fingerprint = create_fingerprint_from_parsed_email(
                        email_data,
                        f"email_{i}",
                        source_file=eml_path
                    )
                    
                    match = index.find_match(
                        fingerprint,
                        use_message_id=config.use_message_id,
                        use_content=config.use_content
                    )
                    
                    if match:
                        duplicate_paths.append(eml_path)
                        result.duplicate_matches.append(match)
                    else:
                        index.add(fingerprint)
                        unique_paths.append(eml_path)
                        
                except Exception as e:
                    unique_paths.append(eml_path)
                    result.warnings.append(f"Parse error {eml_path}: {e}")
            
            result.unique_emails = len(unique_paths)
            result.duplicates_found = len(duplicate_paths)
            
            logger.info(
                f"Found {result.duplicates_found} duplicates, "
                f"{result.unique_emails} unique"
            )
            
            # Step 3: Write output
            self._report_progress(2, 3, "Writing output...")
            
            # Write unique emails
            write_result = self.writer.write(
                unique_paths,
                output_path,
                config.output_format,
                folder_name="Deduplicated"
            )
            
            result.warnings.extend(write_result.warnings)
            if not write_result.success:
                result.errors.extend(write_result.errors)
                return result
            
            # Optionally write duplicates
            if config.keep_duplicates and duplicate_paths:
                dup_output = str(Path(output_path).parent / "duplicates")
                if config.output_format == OutputFormat.MBOX:
                    dup_output += ".mbox"
                elif config.output_format == OutputFormat.PST:
                    dup_output += ".pst"
                
                dup_result = self.writer.write(
                    duplicate_paths,
                    dup_output,
                    config.output_format,
                    folder_name="Duplicates"
                )
                
                if dup_result.success:
                    result.duplicates_path = dup_output
                result.warnings.extend(dup_result.warnings)
            
            result.success = True
            self._report_progress(3, 3, "Deduplication complete!")
            
        except Exception as e:
            result.errors.append(f"Deduplication failed: {e}")
            logger.exception(f"Deduplication failed: {e}")
        
        finally:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass
        
        return result
    
    def get_dedupe_summary(self, result: DedupeResult) -> str:
        """Generate human-readable summary."""
        lines = [
            "=" * 50,
            "DEDUPLICATION SUMMARY",
            "=" * 50,
            "",
            f"Total emails:        {result.total_emails}",
            f"Unique emails:       {result.unique_emails}",
            f"Duplicates removed:  {result.duplicates_found}",
            "",
            f"Output: {result.output_path}",
        ]
        
        if result.duplicates_path:
            lines.append(f"Duplicates saved to: {result.duplicates_path}")
        
        if result.errors:
            lines.extend(["", "ERRORS:"])
            for e in result.errors:
                lines.append(f"  - {e}")
        
        return "\n".join(lines)
