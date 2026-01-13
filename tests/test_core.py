"""
Tests for the Mail Converter application.
"""

import pytest
import tempfile
import os
from pathlib import Path
from datetime import datetime

# Test EML Parser
class TestEMLParser:
    """Tests for the EML Parser module."""
    
    def test_parse_simple_email(self):
        """Test parsing a simple email."""
        from core.eml_parser import EMLParser
        
        # Create a simple test EML
        eml_content = b"""From: sender@example.com
To: recipient@example.com
Subject: Test Email
Date: Mon, 1 Jan 2024 12:00:00 +0000
Content-Type: text/plain; charset="utf-8"

This is the body of the email.
"""
        
        parser = EMLParser()
        result = parser.parse_bytes(eml_content)
        
        assert result.subject == "Test Email"
        assert result.sender_email == "sender@example.com"
        assert "recipient@example.com" in result.recipients_to
        assert "This is the body" in result.body_plain
    
    def test_get_output_filename(self):
        """Test output filename generation."""
        from core.eml_parser import EMLParser
        
        eml_content = b"""From: sender@example.com
To: recipient@example.com
Subject: Meeting Notes: Q4 Review
Date: Wed, 15 Jan 2025 09:30:45 +0000
Content-Type: text/plain

Test body.
"""
        
        parser = EMLParser()
        result = parser.parse_bytes(eml_content)
        
        filename = result.get_output_filename()
        assert filename.startswith("20250115_093045_")
        assert "Meeting_Notes" in filename


# Test Duplicate Detector
class TestDuplicateDetector:
    """Tests for the Duplicate Detector module."""
    
    def test_exact_duplicate_by_message_id(self):
        """Test detection of exact duplicates by Message-ID."""
        from core.duplicate_detector import DuplicateDetector, EmailFingerprint, DuplicateCertainty
        
        detector = DuplicateDetector(min_certainty=DuplicateCertainty.EXACT)
        
        fp1 = EmailFingerprint(
            id="email1",
            message_id="<test123@example.com>",
            sender_email="sender@example.com",
            subject="Test Subject",
            timestamp=datetime(2024, 1, 15, 10, 0, 0),
            content_hash="abc123"
        )
        
        fp2 = EmailFingerprint(
            id="email2",
            message_id="<test123@example.com>",  # Same message ID
            sender_email="sender@example.com",
            subject="Test Subject",
            timestamp=datetime(2024, 1, 15, 10, 0, 0),
            content_hash="abc123"
        )
        
        # First email should be added
        result1 = detector.add_email(fp1)
        assert result1 is None
        
        # Second email should be detected as duplicate
        result2 = detector.add_email(fp2)
        assert result2 is not None
        assert result2.certainty == DuplicateCertainty.EXACT
    
    def test_high_certainty_duplicate(self):
        """Test detection of HIGH certainty duplicates."""
        from core.duplicate_detector import DuplicateDetector, EmailFingerprint, DuplicateCertainty
        
        detector = DuplicateDetector(min_certainty=DuplicateCertainty.HIGH)
        
        fp1 = EmailFingerprint(
            id="email1",
            message_id="<unique1@example.com>",
            sender_email="sender@example.com",
            subject="Test Subject",
            timestamp=datetime(2024, 1, 15, 10, 0, 0),
            content_hash="hash1"
        )
        
        fp2 = EmailFingerprint(
            id="email2",
            message_id="<unique2@example.com>",  # Different message ID
            sender_email="sender@example.com",    # Same sender
            subject="Test Subject",               # Same subject
            timestamp=datetime(2024, 1, 15, 10, 0, 0),  # Same timestamp
            content_hash="hash2"  # Different content
        )
        
        result1 = detector.add_email(fp1)
        assert result1 is None
        
        result2 = detector.add_email(fp2)
        assert result2 is not None
        assert result2.certainty == DuplicateCertainty.HIGH


# Test File Utilities
class TestFileUtils:
    """Tests for file utilities."""
    
    def test_sanitize_filename(self):
        """Test filename sanitization."""
        from utils.file_utils import sanitize_filename
        
        # Test invalid characters
        assert sanitize_filename('file<>:"/\\|?*.txt') == 'file_.txt'
        
        # Test reserved names
        assert sanitize_filename('CON.txt') == '_CON.txt'
        
        # Test length limit
        long_name = 'a' * 300 + '.pdf'
        result = sanitize_filename(long_name, max_length=50)
        assert len(result) <= 50
        assert result.endswith('.pdf')
    
    def test_ensure_dir(self):
        """Test directory creation."""
        from utils.file_utils import ensure_dir
        
        with tempfile.TemporaryDirectory() as tmpdir:
            new_dir = Path(tmpdir) / 'subdir' / 'nested'
            result = ensure_dir(new_dir)
            
            assert result.exists()
            assert result.is_dir()


# Test Attachment Converter Format Detection
class TestAttachmentConverter:
    """Tests for the Attachment Converter module."""
    
    def test_supported_formats(self):
        """Test that all expected formats are supported."""
        from core.attachment_converter import AttachmentConverter
        
        converter = AttachmentConverter(ocr_enabled=False)
        supported = converter.get_supported_extensions()
        
        # Check documents
        assert '.pdf' in supported
        assert '.docx' in supported
        assert '.doc' in supported
        assert '.xlsx' in supported
        assert '.pptx' in supported
        
        # Check images
        assert '.jpg' in supported
        assert '.png' in supported
        assert '.gif' in supported
        assert '.tiff' in supported
        
        # Check email
        assert '.eml' in supported
        assert '.msg' in supported
        
        # Check text
        assert '.txt' in supported
        assert '.csv' in supported
        assert '.html' in supported
        
        converter.cleanup()


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
