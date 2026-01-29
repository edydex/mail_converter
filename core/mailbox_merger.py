"""
Mailbox Merger Module

Merges multiple mailboxes into one, with optional deduplication.
"""

import os
import logging
from pathlib import Path
from typing import List, Optional, Callable, Dict
from dataclasses import dataclass, field
import shutil
import tempfile

from .email_fingerprint import (
    FingerprintIndex,
    create_fingerprint_from_parsed_email
)
from .mailbox_writer import MailboxWriter, OutputFormat, WriteResult
from .pst_extractor import PSTExtractor
from .mbox_extractor import MBOXExtractor
from .eml_parser import EMLParser

logger = logging.getLogger(__name__)


@dataclass
class MergeConfig:
    """Configuration for mailbox merging"""
    # Deduplication options
    deduplicate: bool = True
    use_message_id: bool = True
    use_content: bool = True
    timestamp_tolerance_seconds: int = 15
    
    # Output options
    output_format: OutputFormat = OutputFormat.MBOX


@dataclass
class MergeResult:
    """Result of mailbox merge operation"""
    success: bool
    output_path: str = ""
    
    # Counts
    total_input_emails: int = 0
    emails_written: int = 0
    duplicates_removed: int = 0
    
    # Errors and warnings
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


class MailboxMerger:
    """
    Merges multiple mailboxes into a single mailbox.
    
    Supports:
    - Multiple input formats (PST, MBOX, EML folder)
    - Optional deduplication during merge
    - Multiple output formats
    """
    
    def __init__(
        self,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ):
        """
        Initialize the merger.
        
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
        """Detect input type from path."""
        p = Path(path)
        
        if p.is_dir():
            return "eml_folder"
        
        ext = p.suffix.lower()
        if ext == ".pst":
            return "pst"
        elif ext == ".mbox":
            return "mbox"
        elif ext == ".eml":
            return "eml_file"
        else:
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
        temp_dir: Path,
        index: int
    ) -> tuple[List[str], List[str]]:
        """
        Extract emails from a mailbox to temp directory.
        
        Returns:
            Tuple of (list of EML paths, list of warnings)
        """
        input_type = self._detect_input_type(input_path)
        eml_paths = []
        warnings = []
        
        output_subdir = temp_dir / f"mailbox_{index}"
        output_subdir.mkdir(parents=True, exist_ok=True)
        
        if input_type == "pst":
            result = self.pst_extractor.extract(
                input_path,
                str(output_subdir),
                preserve_structure=False
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
            
        elif input_type == "eml_file":
            # Single EML file
            eml_paths = [input_path]
        
        return eml_paths, warnings
    
    def merge(
        self,
        input_paths: List[str],
        output_path: str,
        config: MergeConfig = None
    ) -> MergeResult:
        """
        Merge multiple mailboxes into one.
        
        Args:
            input_paths: List of paths to mailboxes (PST, MBOX, or EML folders)
            output_path: Output file/folder path
            config: Merge configuration
            
        Returns:
            MergeResult with details
        """
        if config is None:
            config = MergeConfig()
        
        result = MergeResult(success=False, output_path=output_path)
        
        # Create temp directory
        temp_dir = Path(tempfile.mkdtemp(prefix="mailbox_merge_"))
        
        try:
            # Step 1: Extract all mailboxes
            all_eml_paths: List[str] = []
            
            for i, input_path in enumerate(input_paths):
                self._report_progress(
                    i, len(input_paths), 
                    f"Extracting mailbox {i+1}/{len(input_paths)}..."
                )
                
                eml_paths, warnings = self._extract_mailbox(
                    input_path, temp_dir, i
                )
                result.warnings.extend(warnings)
                all_eml_paths.extend(eml_paths)
            
            result.total_input_emails = len(all_eml_paths)
            
            if not all_eml_paths:
                result.errors.append("No emails found in any input mailbox")
                return result
            
            logger.info(f"Found {result.total_input_emails} total emails to merge")
            
            # Step 2: Deduplicate if enabled
            if config.deduplicate:
                self._report_progress(
                    0, result.total_input_emails,
                    "Deduplicating emails..."
                )
                
                unique_paths = self._deduplicate(
                    all_eml_paths, config
                )
                
                result.duplicates_removed = len(all_eml_paths) - len(unique_paths)
                all_eml_paths = unique_paths
                
                logger.info(
                    f"Removed {result.duplicates_removed} duplicates, "
                    f"{len(all_eml_paths)} unique emails"
                )
            
            # Step 3: Write merged output
            self._report_progress(
                0, len(all_eml_paths),
                f"Writing merged output..."
            )
            
            write_result = self.writer.write(
                all_eml_paths,
                output_path,
                config.output_format,
                folder_name="Merged"
            )
            
            result.emails_written = write_result.emails_written
            result.warnings.extend(write_result.warnings)
            
            if not write_result.success:
                result.errors.extend(write_result.errors)
                return result
            
            result.success = True
            self._report_progress(
                len(all_eml_paths), len(all_eml_paths),
                f"Merge complete! {result.emails_written} emails written."
            )
            
        except Exception as e:
            result.errors.append(f"Merge failed: {e}")
            logger.exception(f"Merge failed: {e}")
        
        finally:
            # Cleanup temp directory
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass
        
        return result
    
    def _deduplicate(
        self,
        eml_paths: List[str],
        config: MergeConfig
    ) -> List[str]:
        """
        Remove duplicates from list of EML paths.
        
        Returns:
            List of unique EML paths
        """
        index = FingerprintIndex(
            timestamp_tolerance_seconds=config.timestamp_tolerance_seconds
        )
        
        unique_paths: List[str] = []
        
        for i, eml_path in enumerate(eml_paths):
            if i % 100 == 0:
                self._report_progress(
                    i, len(eml_paths),
                    f"Checking for duplicates: {i}/{len(eml_paths)}"
                )
            
            try:
                email_data = self.eml_parser.parse_file(eml_path)
                
                fingerprint = create_fingerprint_from_parsed_email(
                    email_data,
                    f"email_{i}",
                    source_file=eml_path
                )
                
                # Check if duplicate
                match = index.find_match(
                    fingerprint,
                    use_message_id=config.use_message_id,
                    use_content=config.use_content
                )
                
                if match:
                    # Duplicate found, skip
                    continue
                
                # Not a duplicate, add to index and keep
                index.add(fingerprint)
                unique_paths.append(eml_path)
                
            except Exception as e:
                # Keep files we can't parse (let writer handle them)
                unique_paths.append(eml_path)
                logger.warning(f"Failed to parse {eml_path}: {e}")
        
        return unique_paths
    
    def get_merge_summary(self, result: MergeResult) -> str:
        """Generate human-readable summary of merge."""
        lines = [
            "=" * 50,
            "MAILBOX MERGE SUMMARY",
            "=" * 50,
            "",
            f"Total input emails:   {result.total_input_emails}",
            f"Duplicates removed:   {result.duplicates_removed}",
            f"Emails written:       {result.emails_written}",
            "",
            f"Output: {result.output_path}",
        ]
        
        if result.errors:
            lines.extend(["", "ERRORS:"])
            for e in result.errors:
                lines.append(f"  - {e}")
        
        if result.warnings:
            lines.extend(["", f"WARNINGS ({len(result.warnings)}):"])
            for w in result.warnings[:10]:
                lines.append(f"  - {w}")
            if len(result.warnings) > 10:
                lines.append(f"  ... and {len(result.warnings) - 10} more")
        
        return "\n".join(lines)
