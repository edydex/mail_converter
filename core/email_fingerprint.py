"""
Email Fingerprint Module

Provides unified email fingerprinting for comparison and deduplication.
Used by duplicate_detector, mailbox_comparator, and other tools.
"""

import hashlib
from datetime import datetime, timedelta
from typing import Optional, List
from dataclasses import dataclass
from enum import Enum
import logging
import re

logger = logging.getLogger(__name__)


class MatchCertainty(Enum):
    """Levels of match certainty between emails"""
    EXACT = "exact"           # Same Message-ID (100% certain)
    HIGH = "high"             # Same sender + timestamp + subject + content
    MEDIUM = "medium"         # Same sender + subject (within time window)
    LOW = "low"               # Same subject + similar timestamp


@dataclass
class EmailFingerprint:
    """
    Fingerprint of an email for comparison and deduplication.
    
    Contains all the identifying information needed to determine
    if two emails are the same.
    """
    id: str                       # Unique identifier (e.g., file path)
    message_id: str               # Message-ID header
    sender_email: str             # Sender email address
    subject: str                  # Subject line (original)
    timestamp: Optional[datetime] # Email date/time
    content_hash: str             # Hash of body content
    recipients_hash: str = ""     # Hash of all recipients
    
    # Source tracking
    source_file: str = ""         # Original file path (PST, MBOX, etc.)
    folder_path: str = ""         # Folder within source
    
    def get_message_id_key(self) -> str:
        """Get the Message-ID for exact matching."""
        return self.message_id.strip().lower() if self.message_id else ""
    
    def get_content_key(self) -> str:
        """Get the content hash key."""
        return self.content_hash
    
    def get_sender_subject_key(self) -> str:
        """Get key for sender+subject matching."""
        return f"{self.sender_email.lower()}|{self._normalize_subject()}"
    
    def get_sender_timestamp_subject_key(self, timestamp_tolerance_seconds: int = 0) -> str:
        """
        Get key for sender+timestamp+subject matching.
        
        Args:
            timestamp_tolerance_seconds: Round timestamp to this granularity
        """
        if self.timestamp:
            if timestamp_tolerance_seconds > 0:
                # Round to tolerance bucket
                ts_epoch = int(self.timestamp.timestamp())
                ts_bucket = ts_epoch // timestamp_tolerance_seconds * timestamp_tolerance_seconds
                ts = str(ts_bucket)
            else:
                ts = self.timestamp.strftime("%Y%m%d%H%M")
        else:
            ts = "unknown"
        return f"{self.sender_email.lower()}|{ts}|{self._normalize_subject()}"
    
    def _normalize_subject(self) -> str:
        """Normalize subject for comparison."""
        subject = self.subject.lower().strip()
        
        # Remove common prefixes
        prefixes = ['re:', 'fw:', 'fwd:', 're[', 'fw[', 'aw:', 'antw:']
        changed = True
        while changed:
            changed = False
            for prefix in prefixes:
                if subject.startswith(prefix):
                    subject = subject[len(prefix):].strip()
                    changed = True
                    # Handle numbered prefixes like re[2]:
                    if subject.startswith(']') or (subject and subject[0].isdigit()):
                        idx = subject.find(':')
                        if idx != -1:
                            subject = subject[idx+1:].strip()
        
        return subject
    
    def matches(
        self, 
        other: 'EmailFingerprint',
        use_message_id: bool = True,
        use_content: bool = True,
        timestamp_tolerance_seconds: int = 15
    ) -> Optional[MatchCertainty]:
        """
        Check if this fingerprint matches another.
        
        Args:
            other: Another fingerprint to compare
            use_message_id: Whether to match by Message-ID
            use_content: Whether to match by content hash
            timestamp_tolerance_seconds: Tolerance for timestamp comparison
            
        Returns:
            MatchCertainty if match found, None otherwise
        """
        # EXACT: Same Message-ID
        if use_message_id:
            my_mid = self.get_message_id_key()
            other_mid = other.get_message_id_key()
            if my_mid and other_mid and my_mid == other_mid:
                return MatchCertainty.EXACT
        
        # EXACT: Same content hash
        if use_content:
            if self.content_hash and other.content_hash:
                if self.content_hash == other.content_hash:
                    return MatchCertainty.EXACT
        
        # HIGH: Same sender + subject + similar timestamp + same content
        if self.sender_email.lower() == other.sender_email.lower():
            if self._normalize_subject() == other._normalize_subject():
                if self.timestamp and other.timestamp:
                    time_diff = abs((self.timestamp - other.timestamp).total_seconds())
                    if time_diff <= timestamp_tolerance_seconds:
                        # For HIGH certainty, also check content matches
                        if use_content and self.content_hash == other.content_hash:
                            return MatchCertainty.HIGH
                        elif not use_content:
                            return MatchCertainty.HIGH
        
        return None


@dataclass
class FingerprintMatch:
    """Represents a match between two email fingerprints"""
    fingerprint_a: EmailFingerprint
    fingerprint_b: EmailFingerprint
    certainty: MatchCertainty
    reason: str
    
    def __str__(self):
        return f"{self.certainty.value}: {self.reason}"


class FingerprintIndex:
    """
    Index for fast fingerprint lookups and comparison.
    
    This class maintains various indexes for efficient email comparison
    across large mailboxes.
    """
    
    def __init__(self, timestamp_tolerance_seconds: int = 15):
        """
        Initialize the fingerprint index.
        
        Args:
            timestamp_tolerance_seconds: Tolerance for timestamp matching
        """
        self.timestamp_tolerance = timestamp_tolerance_seconds
        
        # Indexes
        self._message_ids: dict[str, EmailFingerprint] = {}
        self._content_hashes: dict[str, EmailFingerprint] = {}
        self._sender_subject: dict[str, List[EmailFingerprint]] = {}
        
        # All fingerprints by ID
        self._fingerprints: dict[str, EmailFingerprint] = {}
    
    def clear(self):
        """Clear all indexes."""
        self._message_ids.clear()
        self._content_hashes.clear()
        self._sender_subject.clear()
        self._fingerprints.clear()
    
    def add(self, fingerprint: EmailFingerprint):
        """
        Add a fingerprint to the index.
        
        Args:
            fingerprint: Fingerprint to add
        """
        # Store by ID
        self._fingerprints[fingerprint.id] = fingerprint
        
        # Index by Message-ID
        mid_key = fingerprint.get_message_id_key()
        if mid_key:
            self._message_ids[mid_key] = fingerprint
        
        # Index by content hash
        if fingerprint.content_hash:
            self._content_hashes[fingerprint.content_hash] = fingerprint
        
        # Index by sender+subject
        ss_key = fingerprint.get_sender_subject_key()
        if ss_key not in self._sender_subject:
            self._sender_subject[ss_key] = []
        self._sender_subject[ss_key].append(fingerprint)
    
    def find_match(
        self,
        fingerprint: EmailFingerprint,
        use_message_id: bool = True,
        use_content: bool = True
    ) -> Optional[FingerprintMatch]:
        """
        Find a matching fingerprint in the index.
        
        Args:
            fingerprint: Fingerprint to search for
            use_message_id: Whether to match by Message-ID
            use_content: Whether to match by content
            
        Returns:
            FingerprintMatch if found, None otherwise
        """
        # Check Message-ID
        if use_message_id:
            mid_key = fingerprint.get_message_id_key()
            if mid_key and mid_key in self._message_ids:
                matched = self._message_ids[mid_key]
                if matched.id != fingerprint.id:
                    return FingerprintMatch(
                        fingerprint_a=fingerprint,
                        fingerprint_b=matched,
                        certainty=MatchCertainty.EXACT,
                        reason=f"Same Message-ID: {mid_key[:50]}..."
                    )
        
        # Check content hash
        if use_content and fingerprint.content_hash:
            if fingerprint.content_hash in self._content_hashes:
                matched = self._content_hashes[fingerprint.content_hash]
                if matched.id != fingerprint.id:
                    return FingerprintMatch(
                        fingerprint_a=fingerprint,
                        fingerprint_b=matched,
                        certainty=MatchCertainty.EXACT,
                        reason="Identical content hash"
                    )
        
        # Check sender+subject with timestamp tolerance
        ss_key = fingerprint.get_sender_subject_key()
        if ss_key in self._sender_subject:
            for candidate in self._sender_subject[ss_key]:
                if candidate.id == fingerprint.id:
                    continue
                
                certainty = fingerprint.matches(
                    candidate,
                    use_message_id=False,  # Already checked above
                    use_content=use_content,
                    timestamp_tolerance_seconds=self.timestamp_tolerance
                )
                
                if certainty:
                    return FingerprintMatch(
                        fingerprint_a=fingerprint,
                        fingerprint_b=candidate,
                        certainty=certainty,
                        reason=f"Matched by content ({certainty.value})"
                    )
        
        return None
    
    def get_all(self) -> List[EmailFingerprint]:
        """Get all fingerprints in the index."""
        return list(self._fingerprints.values())
    
    def __len__(self) -> int:
        return len(self._fingerprints)
    
    def __contains__(self, fingerprint_id: str) -> bool:
        return fingerprint_id in self._fingerprints


def create_fingerprint(
    email_id: str,
    message_id: str,
    sender_email: str,
    subject: str,
    timestamp: Optional[datetime],
    body_text: str,
    body_html: str = "",
    recipients_to: List[str] = None,
    recipients_cc: List[str] = None,
    source_file: str = "",
    folder_path: str = ""
) -> EmailFingerprint:
    """
    Create an email fingerprint from components.
    
    Args:
        email_id: Unique identifier for this email
        message_id: Message-ID header
        sender_email: Sender's email address
        subject: Email subject
        timestamp: Email timestamp
        body_text: Plain text body
        body_html: HTML body (optional)
        recipients_to: To recipients
        recipients_cc: CC recipients
        source_file: Original source file
        folder_path: Folder within source
        
    Returns:
        EmailFingerprint object
    """
    recipients_to = recipients_to or []
    recipients_cc = recipients_cc or []
    
    # Create content hash from body
    # Use text body preferably, fall back to stripped HTML
    body_for_hash = body_text.strip()
    if not body_for_hash and body_html:
        # Strip HTML tags for hashing
        body_for_hash = re.sub(r'<[^>]+>', '', body_html)[:2000]
    
    content = f"{sender_email.lower()}|{subject}|{body_for_hash[:1000]}"
    content_hash = hashlib.sha256(content.encode(errors='replace')).hexdigest()
    
    # Create recipients hash
    all_recipients = sorted([r.lower() for r in recipients_to + recipients_cc])
    recipients_hash = hashlib.md5("|".join(all_recipients).encode()).hexdigest()
    
    return EmailFingerprint(
        id=email_id,
        message_id=message_id or "",
        sender_email=sender_email or "",
        subject=subject or "",
        timestamp=timestamp,
        content_hash=content_hash,
        recipients_hash=recipients_hash,
        source_file=source_file,
        folder_path=folder_path
    )


def create_fingerprint_from_parsed_email(
    email_data,
    email_id: str,
    source_file: str = "",
    folder_path: str = ""
) -> EmailFingerprint:
    """
    Create a fingerprint from a ParsedEmail object.
    
    Args:
        email_data: ParsedEmail or ParsedMSG object
        email_id: Unique identifier
        source_file: Source file path
        folder_path: Folder path within source
        
    Returns:
        EmailFingerprint
    """
    return create_fingerprint(
        email_id=email_id,
        message_id=getattr(email_data, 'message_id', ''),
        sender_email=getattr(email_data, 'sender_email', ''),
        subject=getattr(email_data, 'subject', ''),
        timestamp=getattr(email_data, 'date', None),
        body_text=getattr(email_data, 'body_text', ''),
        body_html=getattr(email_data, 'body_html', ''),
        recipients_to=getattr(email_data, 'recipients_to', []),
        recipients_cc=getattr(email_data, 'recipients_cc', []),
        source_file=source_file,
        folder_path=folder_path
    )
