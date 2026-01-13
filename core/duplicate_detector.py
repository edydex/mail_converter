"""
Duplicate Detector Module

Detects duplicate emails based on configurable criteria.
"""

import hashlib
from datetime import datetime, timedelta
from typing import List, Dict, Set, Tuple, Optional
from dataclasses import dataclass, field
from enum import Enum
import logging

logger = logging.getLogger(__name__)


class DuplicateCertainty(Enum):
    """Levels of duplicate certainty"""
    EXACT = "exact"           # Same Message-ID (100% certain)
    HIGH = "high"             # Same sender + timestamp + subject
    MEDIUM = "medium"         # Same sender + subject (within time window)
    LOW = "low"               # Same subject + similar timestamp


@dataclass
class DuplicateMatch:
    """Represents a duplicate match between emails"""
    original_id: str
    duplicate_id: str
    certainty: DuplicateCertainty
    reason: str
    
    def __str__(self):
        return f"{self.certainty.value}: {self.reason}"


@dataclass
class EmailFingerprint:
    """Fingerprint of an email for duplicate detection"""
    id: str                    # Unique identifier (e.g., file path)
    message_id: str            # Message-ID header
    sender_email: str          # Sender email address
    subject: str               # Subject line (normalized)
    timestamp: Optional[datetime]
    content_hash: str          # Hash of body content
    
    def get_sender_subject_key(self) -> str:
        """Get key for sender+subject matching."""
        return f"{self.sender_email.lower()}|{self._normalize_subject()}"
    
    def get_sender_timestamp_subject_key(self) -> str:
        """Get key for sender+timestamp+subject matching."""
        ts = self.timestamp.strftime("%Y%m%d%H%M") if self.timestamp else "unknown"
        return f"{self.sender_email.lower()}|{ts}|{self._normalize_subject()}"
    
    def _normalize_subject(self) -> str:
        """Normalize subject for comparison."""
        subject = self.subject.lower().strip()
        
        # Remove common prefixes
        prefixes = ['re:', 'fw:', 'fwd:', 're[', 'fw[', 'aw:']
        for prefix in prefixes:
            while subject.startswith(prefix):
                subject = subject[len(prefix):].strip()
                # Handle numbered prefixes like re[2]:
                if subject.startswith(']'):
                    idx = subject.find(':')
                    if idx != -1:
                        subject = subject[idx+1:].strip()
        
        return subject


class DuplicateDetector:
    """
    Detects duplicate emails with configurable certainty levels.
    """
    
    def __init__(
        self,
        min_certainty: DuplicateCertainty = DuplicateCertainty.HIGH,
        time_window_minutes: int = 5
    ):
        """
        Initialize duplicate detector.
        
        Args:
            min_certainty: Minimum certainty level to consider as duplicate
            time_window_minutes: Time window for MEDIUM certainty matching
        """
        self.min_certainty = min_certainty
        self.time_window = timedelta(minutes=time_window_minutes)
        
        # Indexes for fast lookup
        self._message_ids: Dict[str, str] = {}  # message_id -> email_id
        self._sender_subject: Dict[str, List[str]] = {}  # key -> [email_ids]
        self._sender_ts_subject: Dict[str, str] = {}  # key -> email_id
        self._content_hashes: Dict[str, str] = {}  # content_hash -> email_id
        
        # All fingerprints
        self._fingerprints: Dict[str, EmailFingerprint] = {}
    
    def reset(self):
        """Reset all indexes."""
        self._message_ids.clear()
        self._sender_subject.clear()
        self._sender_ts_subject.clear()
        self._content_hashes.clear()
        self._fingerprints.clear()
    
    def create_fingerprint(
        self,
        email_id: str,
        message_id: str,
        sender_email: str,
        subject: str,
        timestamp: Optional[datetime],
        body_content: str
    ) -> EmailFingerprint:
        """
        Create a fingerprint for an email.
        
        Args:
            email_id: Unique identifier for this email
            message_id: Message-ID header value
            sender_email: Sender's email address
            subject: Email subject
            timestamp: Email timestamp
            body_content: Email body for content hashing
            
        Returns:
            EmailFingerprint object
        """
        # Create content hash
        content = f"{sender_email}|{subject}|{body_content[:1000]}"
        content_hash = hashlib.md5(content.encode()).hexdigest()
        
        return EmailFingerprint(
            id=email_id,
            message_id=message_id,
            sender_email=sender_email,
            subject=subject,
            timestamp=timestamp,
            content_hash=content_hash
        )
    
    def check_duplicate(self, fingerprint: EmailFingerprint) -> Optional[DuplicateMatch]:
        """
        Check if an email is a duplicate of a previously seen email.
        
        Args:
            fingerprint: Email fingerprint to check
            
        Returns:
            DuplicateMatch if duplicate found, None otherwise
        """
        # Check by Message-ID (EXACT certainty)
        if fingerprint.message_id and fingerprint.message_id in self._message_ids:
            original_id = self._message_ids[fingerprint.message_id]
            if original_id != fingerprint.id:
                match = DuplicateMatch(
                    original_id=original_id,
                    duplicate_id=fingerprint.id,
                    certainty=DuplicateCertainty.EXACT,
                    reason=f"Same Message-ID: {fingerprint.message_id}"
                )
                if self._meets_certainty(DuplicateCertainty.EXACT):
                    return match
        
        # Check by content hash (EXACT certainty)
        if fingerprint.content_hash in self._content_hashes:
            original_id = self._content_hashes[fingerprint.content_hash]
            if original_id != fingerprint.id:
                match = DuplicateMatch(
                    original_id=original_id,
                    duplicate_id=fingerprint.id,
                    certainty=DuplicateCertainty.EXACT,
                    reason="Identical content hash"
                )
                if self._meets_certainty(DuplicateCertainty.EXACT):
                    return match
        
        # Check by sender + timestamp + subject (HIGH certainty)
        sts_key = fingerprint.get_sender_timestamp_subject_key()
        if sts_key in self._sender_ts_subject:
            original_id = self._sender_ts_subject[sts_key]
            if original_id != fingerprint.id:
                match = DuplicateMatch(
                    original_id=original_id,
                    duplicate_id=fingerprint.id,
                    certainty=DuplicateCertainty.HIGH,
                    reason="Same sender, timestamp, and subject"
                )
                if self._meets_certainty(DuplicateCertainty.HIGH):
                    return match
        
        # Check by sender + subject within time window (MEDIUM certainty)
        ss_key = fingerprint.get_sender_subject_key()
        if ss_key in self._sender_subject and fingerprint.timestamp:
            for original_id in self._sender_subject[ss_key]:
                if original_id == fingerprint.id:
                    continue
                    
                original = self._fingerprints.get(original_id)
                if original and original.timestamp:
                    time_diff = abs(fingerprint.timestamp - original.timestamp)
                    if time_diff <= self.time_window:
                        match = DuplicateMatch(
                            original_id=original_id,
                            duplicate_id=fingerprint.id,
                            certainty=DuplicateCertainty.MEDIUM,
                            reason=f"Same sender and subject within {self.time_window.seconds // 60} minutes"
                        )
                        if self._meets_certainty(DuplicateCertainty.MEDIUM):
                            return match
        
        # Check by subject only within tight time window (LOW certainty)
        if fingerprint.timestamp:
            normalized_subject = fingerprint._normalize_subject()
            for other_id, other_fp in self._fingerprints.items():
                if other_id == fingerprint.id:
                    continue
                    
                if other_fp._normalize_subject() == normalized_subject and other_fp.timestamp:
                    time_diff = abs(fingerprint.timestamp - other_fp.timestamp)
                    if time_diff <= timedelta(minutes=1):  # Very tight window for LOW
                        match = DuplicateMatch(
                            original_id=other_id,
                            duplicate_id=fingerprint.id,
                            certainty=DuplicateCertainty.LOW,
                            reason="Same subject within 1 minute"
                        )
                        if self._meets_certainty(DuplicateCertainty.LOW):
                            return match
        
        return None
    
    def add_email(self, fingerprint: EmailFingerprint) -> Optional[DuplicateMatch]:
        """
        Add an email to the detector and check for duplicates.
        
        Args:
            fingerprint: Email fingerprint
            
        Returns:
            DuplicateMatch if this is a duplicate, None otherwise
        """
        # Check for duplicates first
        duplicate = self.check_duplicate(fingerprint)
        
        if duplicate:
            logger.info(f"Duplicate found: {duplicate}")
            return duplicate
        
        # Add to indexes
        if fingerprint.message_id:
            self._message_ids[fingerprint.message_id] = fingerprint.id
        
        self._content_hashes[fingerprint.content_hash] = fingerprint.id
        
        sts_key = fingerprint.get_sender_timestamp_subject_key()
        self._sender_ts_subject[sts_key] = fingerprint.id
        
        ss_key = fingerprint.get_sender_subject_key()
        if ss_key not in self._sender_subject:
            self._sender_subject[ss_key] = []
        self._sender_subject[ss_key].append(fingerprint.id)
        
        self._fingerprints[fingerprint.id] = fingerprint
        
        return None
    
    def _meets_certainty(self, certainty: DuplicateCertainty) -> bool:
        """Check if a certainty level meets the minimum threshold."""
        order = [
            DuplicateCertainty.LOW,
            DuplicateCertainty.MEDIUM,
            DuplicateCertainty.HIGH,
            DuplicateCertainty.EXACT
        ]
        return order.index(certainty) >= order.index(self.min_certainty)
    
    def get_statistics(self) -> Dict:
        """Get statistics about processed emails."""
        return {
            'total_emails': len(self._fingerprints),
            'unique_message_ids': len(self._message_ids),
            'unique_content_hashes': len(self._content_hashes),
            'unique_sender_subject_combos': len(self._sender_subject)
        }
    
    def find_all_duplicates(
        self,
        fingerprints: List[EmailFingerprint]
    ) -> Tuple[List[EmailFingerprint], List[DuplicateMatch]]:
        """
        Process a list of fingerprints and return unique emails and duplicate matches.
        
        Args:
            fingerprints: List of email fingerprints
            
        Returns:
            Tuple of (unique_fingerprints, duplicate_matches)
        """
        self.reset()
        
        unique = []
        duplicates = []
        
        for fp in fingerprints:
            match = self.add_email(fp)
            if match:
                duplicates.append(match)
            else:
                unique.append(fp)
        
        return unique, duplicates


def create_fingerprint_from_parsed_email(email_data, email_id: str) -> EmailFingerprint:
    """
    Helper function to create a fingerprint from a ParsedEmail object.
    
    Args:
        email_data: ParsedEmail object from eml_parser
        email_id: Unique identifier for this email
        
    Returns:
        EmailFingerprint
    """
    return EmailFingerprint(
        id=email_id,
        message_id=email_data.message_id,
        sender_email=email_data.sender_email,
        subject=email_data.subject,
        timestamp=email_data.date,
        content_hash=email_data.content_hash
    )
