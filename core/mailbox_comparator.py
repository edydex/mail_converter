"""
Mailbox Comparator Module

Compares two mailboxes and identifies common and unique emails.
Outputs: Common, Unique to A, Unique to B
"""

import os
import logging
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict
from dataclasses import dataclass, field
from enum import Enum
import shutil

from .email_fingerprint import (
    EmailFingerprint,
    FingerprintIndex,
    FingerprintMatch,
    MatchCertainty,
    create_fingerprint_from_parsed_email
)
from .mailbox_writer import MailboxWriter, OutputFormat, WriteResult
from .pst_extractor import PSTExtractor
from .mbox_extractor import MBOXExtractor
from .eml_parser import EMLParser

logger = logging.getLogger(__name__)


@dataclass
class ComparisonConfig:
    """Configuration for mailbox comparison"""
    # Matching options
    use_message_id: bool = True
    use_content: bool = True
    timestamp_tolerance_seconds: int = 15
    
    # Output options
    output_format: OutputFormat = OutputFormat.EML_FOLDER
    
    # What to output
    output_common: bool = True
    output_unique_a: bool = True
    output_unique_b: bool = True


@dataclass
class ComparisonResult:
    """Result of mailbox comparison"""
    success: bool
    
    # Counts
    total_in_a: int = 0
    total_in_b: int = 0
    common_count: int = 0
    unique_to_a_count: int = 0
    unique_to_b_count: int = 0
    
    # Output paths (set after writing)
    common_output_path: Optional[str] = None
    unique_a_output_path: Optional[str] = None
    unique_b_output_path: Optional[str] = None
    
    # Errors and warnings
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    # Detailed matches (for debugging/reporting)
    matches: List[FingerprintMatch] = field(default_factory=list)


class MailboxComparator:
    """
    Compares two mailboxes and identifies common/unique emails.
    
    Supports PST, MBOX, and EML folder inputs.
    Outputs to MBOX, EML folder, or PST (Windows only).
    """
    
    def __init__(
        self,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ):
        """
        Initialize the comparator.
        
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
        """Detect if input is PST, MBOX, or EML folder."""
        p = Path(path)
        
        if p.is_dir():
            return "eml_folder"
        
        ext = p.suffix.lower()
        if ext == ".pst":
            return "pst"
        elif ext == ".mbox":
            return "mbox"
        else:
            # Assume EML file, but we expect folders
            return "eml_folder"
    
    def _extract_to_temp(
        self, 
        input_path: str, 
        temp_dir: Path,
        label: str
    ) -> Tuple[List[str], List[str]]:
        """
        Extract emails from input to temp directory.
        
        Args:
            input_path: Path to PST, MBOX, or EML folder
            temp_dir: Temporary directory for extracted EMLs
            label: Label for this input (A or B)
            
        Returns:
            Tuple of (list of EML paths, list of warnings)
        """
        input_type = self._detect_input_type(input_path)
        eml_paths = []
        warnings = []
        
        self._report_progress(0, 1, f"Extracting mailbox {label}...")
        logger.info(f"Extracting mailbox {label} from: {input_path}")
        logger.info(f"Input type detected: {input_type}")
        
        if input_type == "pst":
            # Extract PST
            output_dir = temp_dir / f"extracted_{label}"
            output_dir.mkdir(parents=True, exist_ok=True)
            
            logger.info(f"Extracting PST to: {output_dir}")
            
            result = self.pst_extractor.extract(
                input_path,
                str(output_dir),
                preserve_structure=False  # Flatten for comparison
            )
            
            if not result.success:
                logger.error(f"PST extraction failed: {result.errors}")
                warnings.extend(result.errors)
                return [], warnings
            
            warnings.extend(result.warnings)
            
            # Collect all email files - readpst creates numbered files WITHOUT .eml extension
            # Look for both numbered files (1, 2, 3...) and .eml files
            eml_paths = self._collect_email_files(output_dir)
            logger.info(f"Found {len(eml_paths)} email files in extracted PST")
            
        elif input_type == "mbox":
            # Extract MBOX
            output_dir = temp_dir / f"extracted_{label}"
            output_dir.mkdir(parents=True, exist_ok=True)
            
            logger.info(f"Extracting MBOX to: {output_dir}")
            
            result = self.mbox_extractor.extract(
                input_path,
                str(output_dir),
                preserve_structure=False
            )
            
            if not result.success:
                logger.error(f"MBOX extraction failed: {result.errors}")
                warnings.extend(result.errors)
                return [], warnings
            
            warnings.extend(result.warnings)
            eml_paths = result.extracted_files
            logger.info(f"MBOX extraction returned {len(eml_paths)} files")
            
        elif input_type == "eml_folder":
            # EML folder - collect all email files
            input_dir = Path(input_path)
            eml_paths = self._collect_email_files(input_dir)
            logger.info(f"Found {len(eml_paths)} email files in folder")
        
        logger.info(f"Extracted {len(eml_paths)} emails from {label}")
        if len(eml_paths) > 0:
            logger.debug(f"Sample paths: {eml_paths[:3]}")
        
        return eml_paths, warnings
    
    def _collect_email_files(self, directory: Path) -> List[str]:
        """
        Collect all email files from a directory.
        
        readpst creates numbered files (1, 2, 3...) without extensions.
        Also handles standard .eml files.
        
        Args:
            directory: Directory to search
            
        Returns:
            List of email file paths
        """
        email_files = []
        
        logger.debug(f"Scanning directory: {directory}")
        
        for item in directory.rglob('*'):
            if item.is_file():
                # Check if it's a numbered file (readpst output) or .eml file
                if item.name.isdigit() or item.suffix.lower() == '.eml':
                    email_files.append(str(item))
                # Also check for common email file patterns
                elif item.suffix.lower() in ['.msg', '.email']:
                    email_files.append(str(item))
        
        # Log what we found
        if email_files:
            logger.debug(f"Found {len(email_files)} email files")
        else:
            # Log directory contents for debugging
            all_files = list(directory.rglob('*'))
            logger.warning(f"No email files found! Directory contains {len(all_files)} items:")
            for f in all_files[:20]:  # Show first 20
                logger.warning(f"  - {f} (is_file={f.is_file()}, name={f.name})")
        
        return email_files
    
    def _build_fingerprint_index(
        self,
        eml_paths: List[str],
        source_label: str,
        config: ComparisonConfig
    ) -> Tuple[FingerprintIndex, Dict[str, str], List[str]]:
        """
        Build fingerprint index from EML files.
        
        Args:
            eml_paths: List of EML file paths
            source_label: Label for source (A or B)
            config: Comparison configuration
            
        Returns:
            Tuple of (FingerprintIndex, {fingerprint_id: eml_path}, warnings)
        """
        index = FingerprintIndex(
            timestamp_tolerance_seconds=config.timestamp_tolerance_seconds
        )
        id_to_path: Dict[str, str] = {}
        warnings = []
        
        total = len(eml_paths)
        for i, eml_path in enumerate(eml_paths):
            if i % 100 == 0:
                self._report_progress(i, total, f"Indexing {source_label}: {i}/{total}")
            
            try:
                email_data = self.eml_parser.parse_file(eml_path)
                fingerprint_id = f"{source_label}_{i}_{Path(eml_path).name}"
                
                fingerprint = create_fingerprint_from_parsed_email(
                    email_data,
                    fingerprint_id,
                    source_file=eml_path,
                    folder_path=""
                )
                
                index.add(fingerprint)
                id_to_path[fingerprint_id] = eml_path
                
            except Exception as e:
                warnings.append(f"Failed to parse {eml_path}: {e}")
                logger.warning(f"Failed to parse {eml_path}: {e}")
        
        return index, id_to_path, warnings
    
    def compare(
        self,
        mailbox_a_path: str,
        mailbox_b_path: str,
        output_dir: str,
        config: ComparisonConfig = None
    ) -> ComparisonResult:
        """
        Compare two mailboxes and output common/unique emails.
        
        Args:
            mailbox_a_path: Path to first mailbox (PST, MBOX, or EML folder)
            mailbox_b_path: Path to second mailbox
            output_dir: Directory to write output
            config: Comparison configuration
            
        Returns:
            ComparisonResult with details
        """
        if config is None:
            config = ComparisonConfig()
        
        result = ComparisonResult(success=False)
        
        # Create temp directory for extraction
        import tempfile
        temp_dir = Path(tempfile.mkdtemp(prefix="mailbox_compare_"))
        
        try:
            # Step 1: Extract both mailboxes
            self._report_progress(0, 4, "Step 1/4: Extracting mailbox A...")
            logger.info(f"Starting comparison: A={mailbox_a_path}, B={mailbox_b_path}")
            
            eml_paths_a, warnings_a = self._extract_to_temp(
                mailbox_a_path, temp_dir, "A"
            )
            result.warnings.extend(warnings_a)
            result.total_in_a = len(eml_paths_a)
            
            if not eml_paths_a:
                # More detailed error message
                input_type = self._detect_input_type(mailbox_a_path)
                error_msg = (
                    f"No emails found in mailbox A ({input_type}). "
                    f"Path: {mailbox_a_path}. "
                    f"Check logs/mail_converter.log for details."
                )
                if warnings_a:
                    error_msg += f" Warnings: {'; '.join(warnings_a)}"
                result.errors.append(error_msg)
                logger.error(error_msg)
                return result
            
            self._report_progress(1, 4, "Step 2/4: Extracting mailbox B...")
            eml_paths_b, warnings_b = self._extract_to_temp(
                mailbox_b_path, temp_dir, "B"
            )
            result.warnings.extend(warnings_b)
            result.total_in_b = len(eml_paths_b)
            
            if not eml_paths_b:
                input_type = self._detect_input_type(mailbox_b_path)
                error_msg = (
                    f"No emails found in mailbox B ({input_type}). "
                    f"Path: {mailbox_b_path}. "
                    f"Check logs/mail_converter.log for details."
                )
                if warnings_b:
                    error_msg += f" Warnings: {'; '.join(warnings_b)}"
                result.errors.append(error_msg)
                logger.error(error_msg)
                return result
            
            # Step 2: Build fingerprint indexes
            self._report_progress(2, 4, "Step 3/4: Building indexes and comparing...")
            
            index_a, id_to_path_a, warnings_idx_a = self._build_fingerprint_index(
                eml_paths_a, "A", config
            )
            result.warnings.extend(warnings_idx_a)
            
            index_b, id_to_path_b, warnings_idx_b = self._build_fingerprint_index(
                eml_paths_b, "B", config
            )
            result.warnings.extend(warnings_idx_b)
            
            # Step 3: Compare
            # For each email in A, check if it exists in B
            common_from_a: List[str] = []  # EML paths
            unique_to_a: List[str] = []
            matched_in_b: set = set()  # Fingerprint IDs that matched
            
            for fp_a in index_a.get_all():
                match = index_b.find_match(
                    fp_a,
                    use_message_id=config.use_message_id,
                    use_content=config.use_content
                )
                
                if match:
                    common_from_a.append(id_to_path_a[fp_a.id])
                    matched_in_b.add(match.fingerprint_b.id)
                    result.matches.append(match)
                else:
                    unique_to_a.append(id_to_path_a[fp_a.id])
            
            # Emails in B that weren't matched are unique to B
            unique_to_b: List[str] = []
            for fp_b in index_b.get_all():
                if fp_b.id not in matched_in_b:
                    unique_to_b.append(id_to_path_b[fp_b.id])
            
            result.common_count = len(common_from_a)
            result.unique_to_a_count = len(unique_to_a)
            result.unique_to_b_count = len(unique_to_b)
            
            logger.info(
                f"Comparison complete: {result.common_count} common, "
                f"{result.unique_to_a_count} unique to A, "
                f"{result.unique_to_b_count} unique to B"
            )
            
            # Step 4: Write output
            self._report_progress(3, 4, "Step 4/4: Writing output...")
            output_base = Path(output_dir)
            output_base.mkdir(parents=True, exist_ok=True)
            
            categories = {}
            if config.output_common and common_from_a:
                categories["common"] = common_from_a
            if config.output_unique_a and unique_to_a:
                categories["unique_to_A"] = unique_to_a
            if config.output_unique_b and unique_to_b:
                categories["unique_to_B"] = unique_to_b
            
            write_results = self.writer.write_categorized(
                categories, output_dir, config.output_format
            )
            
            # Set output paths
            for category, wr in write_results.items():
                if category == "common":
                    result.common_output_path = wr.output_path
                elif category == "unique_to_A":
                    result.unique_a_output_path = wr.output_path
                elif category == "unique_to_B":
                    result.unique_b_output_path = wr.output_path
                
                if not wr.success:
                    result.warnings.extend(wr.errors)
            
            result.success = True
            self._report_progress(4, 4, "Comparison complete!")
            
        except Exception as e:
            result.errors.append(f"Comparison failed: {e}")
            logger.exception(f"Comparison failed: {e}")
        
        finally:
            # Cleanup temp directory
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass
        
        return result
    
    def get_comparison_summary(self, result: ComparisonResult) -> str:
        """Generate a human-readable summary of the comparison."""
        lines = [
            "=" * 50,
            "MAILBOX COMPARISON SUMMARY",
            "=" * 50,
            "",
            f"Mailbox A: {result.total_in_a} emails",
            f"Mailbox B: {result.total_in_b} emails",
            "",
            f"Common (in both):     {result.common_count}",
            f"Unique to A:          {result.unique_to_a_count}",
            f"Unique to B:          {result.unique_to_b_count}",
            "",
        ]
        
        if result.common_output_path:
            lines.append(f"Common output:     {result.common_output_path}")
        if result.unique_a_output_path:
            lines.append(f"Unique A output:   {result.unique_a_output_path}")
        if result.unique_b_output_path:
            lines.append(f"Unique B output:   {result.unique_b_output_path}")
        
        if result.errors:
            lines.extend(["", "ERRORS:"])
            for e in result.errors:
                lines.append(f"  - {e}")
        
        if result.warnings:
            lines.extend(["", f"WARNINGS ({len(result.warnings)}):"])
            for w in result.warnings[:10]:  # Show first 10
                lines.append(f"  - {w}")
            if len(result.warnings) > 10:
                lines.append(f"  ... and {len(result.warnings) - 10} more")
        
        return "\n".join(lines)
