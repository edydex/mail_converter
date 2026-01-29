"""
Mailbox Filter Module

Filter emails from a mailbox by sender/recipient.
Extracts emails to/from specific email addresses.
"""

import os
import re
import logging
from pathlib import Path
from typing import List, Optional, Callable, Set
from dataclasses import dataclass, field
import shutil
import tempfile

from .mailbox_writer import MailboxWriter, OutputFormat
from .pst_extractor import PSTExtractor
from .mbox_extractor import MBOXExtractor
from .eml_parser import EMLParser

logger = logging.getLogger(__name__)


@dataclass
class FilterConfig:
    """Configuration for email filtering"""
    # Filter criteria (OR logic - email matches if ANY criteria matches)
    sender_emails: List[str] = field(default_factory=list)      # Match if sender in list
    sender_domains: List[str] = field(default_factory=list)     # Match if sender domain in list
    recipient_emails: List[str] = field(default_factory=list)   # Match if any recipient in list
    recipient_domains: List[str] = field(default_factory=list)  # Match if any recipient domain in list
    
    # Logic
    match_mode: str = "any"  # "any" = OR, "all" = AND (for multiple criteria)
    include_cc: bool = True  # Include CC recipients in matching
    include_bcc: bool = True # Include BCC recipients in matching
    
    # Output
    output_format: OutputFormat = OutputFormat.EML_FOLDER
    output_non_matching: bool = False  # Also output non-matching emails


@dataclass
class FilterResult:
    """Result of filter operation"""
    success: bool
    matched_output_path: str = ""
    non_matched_output_path: Optional[str] = None
    
    # Counts
    total_emails: int = 0
    matched_emails: int = 0
    non_matched_emails: int = 0
    
    # Errors and warnings
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


class MailboxFilter:
    """
    Filters emails from a mailbox by sender/recipient.
    
    Useful for extracting correspondence with specific
    email addresses or domains.
    """
    
    def __init__(
        self,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ):
        """
        Initialize the filter.
        
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
    
    def _extract_email_address(self, email_str: str) -> str:
        """Extract just the email address from a string like 'Name <email@domain.com>'."""
        if not email_str:
            return ""
        
        # Try to extract from angle brackets
        match = re.search(r'<([^>]+)>', email_str)
        if match:
            return match.group(1).lower().strip()
        
        # Otherwise return cleaned string
        return email_str.lower().strip()
    
    def _extract_domain(self, email: str) -> str:
        """Extract domain from email address."""
        email = self._extract_email_address(email)
        if '@' in email:
            return email.split('@')[1]
        return ""
    
    def _matches_filter(self, email_data, config: FilterConfig) -> bool:
        """Check if an email matches the filter criteria."""
        # Collect all addresses to check
        sender_email = self._extract_email_address(email_data.sender_email)
        sender_domain = self._extract_domain(sender_email)
        
        all_recipients: List[str] = []
        all_recipient_domains: Set[str] = set()
        
        for r in email_data.recipients_to:
            email = self._extract_email_address(r)
            all_recipients.append(email)
            domain = self._extract_domain(email)
            if domain:
                all_recipient_domains.add(domain)
        
        if config.include_cc:
            for r in email_data.recipients_cc:
                email = self._extract_email_address(r)
                all_recipients.append(email)
                domain = self._extract_domain(email)
                if domain:
                    all_recipient_domains.add(domain)
        
        if config.include_bcc:
            for r in getattr(email_data, 'recipients_bcc', []):
                email = self._extract_email_address(r)
                all_recipients.append(email)
                domain = self._extract_domain(email)
                if domain:
                    all_recipient_domains.add(domain)
        
        # Check criteria
        matches = []
        
        # Sender email match
        if config.sender_emails:
            sender_match = sender_email in [e.lower() for e in config.sender_emails]
            matches.append(sender_match)
        
        # Sender domain match
        if config.sender_domains:
            domain_match = sender_domain in [d.lower() for d in config.sender_domains]
            matches.append(domain_match)
        
        # Recipient email match
        if config.recipient_emails:
            filter_emails = [e.lower() for e in config.recipient_emails]
            recipient_match = any(r in filter_emails for r in all_recipients)
            matches.append(recipient_match)
        
        # Recipient domain match
        if config.recipient_domains:
            filter_domains = [d.lower() for d in config.recipient_domains]
            domain_match = any(d in filter_domains for d in all_recipient_domains)
            matches.append(domain_match)
        
        # If no criteria specified, don't match anything
        if not matches:
            return False
        
        # Apply match mode
        if config.match_mode == "all":
            return all(matches)
        else:  # "any"
            return any(matches)
    
    def filter(
        self,
        input_path: str,
        output_path: str,
        config: FilterConfig
    ) -> FilterResult:
        """
        Filter emails from a mailbox by sender/recipient.
        
        Args:
            input_path: Path to input mailbox
            output_path: Path for filtered output
            config: Filter configuration
            
        Returns:
            FilterResult with details
        """
        result = FilterResult(success=False, matched_output_path=output_path)
        
        # Validate config
        has_criteria = (
            config.sender_emails or 
            config.sender_domains or 
            config.recipient_emails or 
            config.recipient_domains
        )
        if not has_criteria:
            result.errors.append("No filter criteria specified")
            return result
        
        temp_dir = Path(tempfile.mkdtemp(prefix="mailbox_filter_"))
        
        try:
            # Step 1: Extract mailbox
            self._report_progress(0, 3, "Extracting mailbox...")
            eml_paths, warnings = self._extract_mailbox(input_path, temp_dir)
            result.warnings.extend(warnings)
            result.total_emails = len(eml_paths)
            
            if not eml_paths:
                result.errors.append("No emails found in mailbox")
                return result
            
            # Step 2: Filter emails
            self._report_progress(1, 3, "Filtering emails...")
            
            matched_paths: List[str] = []
            non_matched_paths: List[str] = []
            
            for i, eml_path in enumerate(eml_paths):
                if i % 100 == 0:
                    self._report_progress(
                        i, result.total_emails,
                        f"Filtering: {i}/{result.total_emails}"
                    )
                
                try:
                    email_data = self.eml_parser.parse_file(eml_path)
                    
                    if self._matches_filter(email_data, config):
                        matched_paths.append(eml_path)
                    else:
                        non_matched_paths.append(eml_path)
                        
                except Exception as e:
                    result.warnings.append(f"Parse error {eml_path}: {e}")
                    non_matched_paths.append(eml_path)
            
            result.matched_emails = len(matched_paths)
            result.non_matched_emails = len(non_matched_paths)
            
            logger.info(
                f"Filter complete: {result.matched_emails} matched, "
                f"{result.non_matched_emails} not matched"
            )
            
            # Step 3: Write output
            self._report_progress(2, 3, "Writing output...")
            
            if matched_paths:
                write_result = self.writer.write(
                    matched_paths,
                    output_path,
                    config.output_format,
                    folder_name="Filtered"
                )
                result.warnings.extend(write_result.warnings)
                if not write_result.success:
                    result.errors.extend(write_result.errors)
                    return result
            else:
                result.warnings.append("No emails matched the filter criteria")
            
            # Optionally write non-matching
            if config.output_non_matching and non_matched_paths:
                non_match_output = str(Path(output_path).parent / "non_matching")
                if config.output_format == OutputFormat.MBOX:
                    non_match_output += ".mbox"
                elif config.output_format == OutputFormat.PST:
                    non_match_output += ".pst"
                
                non_match_result = self.writer.write(
                    non_matched_paths,
                    non_match_output,
                    config.output_format,
                    folder_name="Non-Matching"
                )
                
                if non_match_result.success:
                    result.non_matched_output_path = non_match_output
                result.warnings.extend(non_match_result.warnings)
            
            result.success = True
            self._report_progress(3, 3, "Filtering complete!")
            
        except Exception as e:
            result.errors.append(f"Filter failed: {e}")
            logger.exception(f"Filter failed: {e}")
        
        finally:
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass
        
        return result
    
    def get_filter_summary(self, result: FilterResult) -> str:
        """Generate human-readable summary."""
        lines = [
            "=" * 50,
            "FILTER SUMMARY",
            "=" * 50,
            "",
            f"Total emails:       {result.total_emails}",
            f"Matched:            {result.matched_emails}",
            f"Not matched:        {result.non_matched_emails}",
            "",
            f"Output: {result.matched_output_path}",
        ]
        
        if result.non_matched_output_path:
            lines.append(f"Non-matching: {result.non_matched_output_path}")
        
        if result.errors:
            lines.extend(["", "ERRORS:"])
            for e in result.errors:
                lines.append(f"  - {e}")
        
        return "\n".join(lines)
