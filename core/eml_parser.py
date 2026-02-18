"""
EML Parser Module

Parses EML files to extract metadata, body content, and attachments.
"""

import os
import email
import email.policy
import hashlib
import re
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime, parseaddr
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple, BinaryIO
from dataclasses import dataclass, field
import logging

from .rtf_converter import convert_rtf_body

logger = logging.getLogger(__name__)


@dataclass
class Attachment:
    """Represents an email attachment"""
    filename: str
    content_type: str
    content: bytes
    size: int
    content_id: Optional[str] = None  # For inline images
    
    def get_extension(self) -> str:
        """Get file extension from filename or content type."""
        if self.filename:
            ext = Path(self.filename).suffix.lower()
            if ext:
                return ext
        
        # Fallback to content type
        type_to_ext = {
            'application/pdf': '.pdf',
            'application/msword': '.doc',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
            'application/vnd.ms-excel': '.xls',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
            'application/vnd.ms-powerpoint': '.ppt',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
            'image/jpeg': '.jpg',
            'image/png': '.png',
            'image/gif': '.gif',
            'image/bmp': '.bmp',
            'image/tiff': '.tiff',
            'text/plain': '.txt',
            'text/html': '.html',
            'text/csv': '.csv',
            'message/rfc822': '.eml',
            'application/vnd.ms-outlook': '.msg',
        }
        
        return type_to_ext.get(self.content_type, '.bin')
    
    def save_to_file(self, directory: str, filename: Optional[str] = None) -> Path:
        """Save attachment to a file."""
        if filename is None:
            filename = self.filename or f"attachment{self.get_extension()}"
        
        # Sanitize filename
        filename = self._sanitize_filename(filename)
        
        filepath = Path(directory) / filename
        filepath.parent.mkdir(parents=True, exist_ok=True)
        
        # Handle duplicate filenames
        counter = 1
        original_stem = filepath.stem
        while filepath.exists():
            filepath = filepath.parent / f"{original_stem}_{counter}{filepath.suffix}"
            counter += 1
        
        with open(filepath, 'wb') as f:
            f.write(self.content)
        
        return filepath
    
    @staticmethod
    def _sanitize_filename(filename: str) -> str:
        """Remove or replace invalid characters from filename."""
        # Remove or replace invalid characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        # Remove leading/trailing spaces and dots
        filename = filename.strip('. ')
        
        # Limit length
        if len(filename) > 200:
            ext = Path(filename).suffix
            filename = filename[:200-len(ext)] + ext
        
        return filename or "unnamed_attachment"


@dataclass
class ParsedEmail:
    """Represents a parsed email with all its components"""
    # Metadata
    message_id: str
    subject: str
    sender: str
    sender_email: str
    recipients_to: List[str]
    recipients_cc: List[str]
    recipients_bcc: List[str]
    date: Optional[datetime]
    
    # Content
    body_plain: str
    body_html: str
    
    # Attachments
    attachments: List[Attachment]
    inline_images: Dict[str, Attachment]  # content_id -> attachment
    
    # Source
    source_path: Path
    raw_headers: Dict[str, str]
    
    # For duplicate detection
    content_hash: str = ""
    
    def __post_init__(self):
        """Calculate content hash after initialization."""
        if not self.content_hash:
            self.content_hash = self._calculate_hash()
    
    def _calculate_hash(self) -> str:
        """Calculate a hash of the email content for duplicate detection."""
        content = f"{self.sender_email}|{self.subject}|{self.body_plain[:1000]}"
        return hashlib.md5(content.encode()).hexdigest()
    
    def get_timestamp_prefix(self) -> str:
        """Get timestamp prefix for filename: YYYYMMDD_HHMMSS"""
        if self.date:
            return self.date.strftime("%Y%m%d_%H%M%S")
        return "00000000_000000"
    
    def get_safe_subject(self, max_length: int = 50) -> str:
        """Get sanitized subject for use in filenames."""
        subject = self.subject or "No_Subject"
        
        # Remove/replace invalid filename characters
        invalid_chars = '<>:"/\\|?*\n\r\t'
        for char in invalid_chars:
            subject = subject.replace(char, '_')
        
        # Replace multiple underscores/spaces with single underscore
        subject = re.sub(r'[_\s]+', '_', subject)
        
        # Trim and limit length
        subject = subject.strip('_. ')[:max_length]
        
        return subject or "No_Subject"
    
    def get_output_filename(self) -> str:
        """Get the standard output filename for this email."""
        return f"{self.get_timestamp_prefix()}_{self.get_safe_subject()}"
    
    def get_display_date(self) -> str:
        """Get formatted date for display."""
        if self.date:
            return self.date.strftime("%Y-%m-%d %H:%M:%S")
        return "Unknown Date"


class EMLParser:
    """
    Parses EML files and extracts all components.
    """
    
    def __init__(self):
        self.policy = email.policy.default
    
    def parse_file(self, eml_path: str) -> ParsedEmail:
        """
        Parse an EML file.
        
        Args:
            eml_path: Path to the EML file
            
        Returns:
            ParsedEmail object with all extracted data
        """
        with open(eml_path, 'rb') as f:
            return self.parse_bytes(f.read(), Path(eml_path))
    
    def parse_bytes(self, eml_bytes: bytes, source_path: Optional[Path] = None) -> ParsedEmail:
        """
        Parse EML content from bytes.
        
        Args:
            eml_bytes: Raw EML content
            source_path: Optional source path for reference
            
        Returns:
            ParsedEmail object with all extracted data
        """
        parser = BytesParser(policy=self.policy)
        msg = parser.parsebytes(eml_bytes)
        
        return self._parse_message(msg, source_path or Path("unknown.eml"))
    
    def _parse_message(self, msg: email.message.Message, source_path: Path) -> ParsedEmail:
        """Parse an email.message.Message object."""
        
        # Extract metadata
        message_id = msg.get('Message-ID', '') or self._generate_message_id(msg)
        subject = msg.get('Subject', '') or ''
        
        # Parse sender
        from_header = msg.get('From', '')
        sender_name, sender_email = parseaddr(from_header)
        sender = sender_name or sender_email or from_header
        
        # Parse recipients
        recipients_to = self._parse_address_list(msg.get('To', ''))
        recipients_cc = self._parse_address_list(msg.get('Cc', ''))
        recipients_bcc = self._parse_address_list(msg.get('Bcc', ''))
        
        # Parse date
        date = self._parse_date(msg.get('Date', ''))
        
        # Extract body and attachments
        body_plain, body_html, attachments, inline_images = self._extract_content(msg)
        
        # Extract raw headers
        raw_headers = {key: str(value) for key, value in msg.items()}
        
        return ParsedEmail(
            message_id=message_id,
            subject=subject,
            sender=sender,
            sender_email=sender_email,
            recipients_to=recipients_to,
            recipients_cc=recipients_cc,
            recipients_bcc=recipients_bcc,
            date=date,
            body_plain=body_plain,
            body_html=body_html,
            attachments=attachments,
            inline_images=inline_images,
            source_path=source_path,
            raw_headers=raw_headers
        )
    
    def _parse_address_list(self, header: str) -> List[str]:
        """Parse a comma-separated list of email addresses."""
        if not header:
            return []
        
        addresses = []
        for addr in header.split(','):
            addr = addr.strip()
            if addr:
                name, email_addr = parseaddr(addr)
                addresses.append(name or email_addr or addr)
        
        return addresses
    
    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse email date string to datetime."""
        if not date_str:
            return None
        
        try:
            return parsedate_to_datetime(date_str)
        except (ValueError, TypeError) as e:
            logger.warning(f"Failed to parse date '{date_str}': {e}")
            
            # Try some common alternative formats
            alternative_formats = [
                "%a, %d %b %Y %H:%M:%S",
                "%d %b %Y %H:%M:%S",
                "%Y-%m-%d %H:%M:%S",
            ]
            
            for fmt in alternative_formats:
                try:
                    # Strip timezone info for simple parsing
                    clean_date = re.sub(r'\s*[+-]\d{4}.*$', '', date_str)
                    return datetime.strptime(clean_date.strip(), fmt)
                except ValueError:
                    continue
            
            return None
    
    def _generate_message_id(self, msg: email.message.Message) -> str:
        """Generate a message ID if none exists."""
        # Create hash from available headers
        content = f"{msg.get('From', '')}|{msg.get('Subject', '')}|{msg.get('Date', '')}"
        hash_val = hashlib.md5(content.encode()).hexdigest()[:16]
        return f"<generated-{hash_val}@local>"
    
    def _extract_content(
        self, 
        msg: email.message.Message
    ) -> Tuple[str, str, List[Attachment], Dict[str, Attachment]]:
        """
        Extract body text and attachments from message.
        
        Returns:
            Tuple of (plain_text_body, html_body, attachments, inline_images)
        """
        body_plain = ""
        body_html = ""
        attachments = []
        inline_images = {}
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))
                content_id = part.get("Content-ID", "")
                
                # Clean content ID (remove < >)
                if content_id:
                    content_id = content_id.strip('<>')
                
                try:
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        payload = part.get_payload(decode=True)
                        if payload:
                            charset = part.get_content_charset() or 'utf-8'
                            body_plain += self._decode_payload(payload, charset)
                    
                    elif content_type == "text/html" and "attachment" not in content_disposition:
                        payload = part.get_payload(decode=True)
                        if payload:
                            charset = part.get_content_charset() or 'utf-8'
                            body_html += self._decode_payload(payload, charset)
                    
                    elif "attachment" in content_disposition or part.get_filename():
                        # This is an attachment
                        attachment = self._extract_attachment(part)
                        if attachment:
                            attachments.append(attachment)
                            # If this attachment has a Content-ID, also register
                            # it as an inline image so cid: references in the
                            # HTML body can be resolved.
                            if content_id and content_type.startswith("image/"):
                                attachment.content_id = content_id
                                inline_images[content_id] = attachment
                    
                    elif content_type.startswith("image/") and content_id:
                        # Inline image (without filename / disposition)
                        attachment = self._extract_attachment(part)
                        if attachment:
                            attachment.content_id = content_id
                            inline_images[content_id] = attachment
                    
                    elif content_type == "message/rfc822":
                        # Nested email - treat as attachment
                        attachment = self._extract_attachment(part)
                        if attachment:
                            attachments.append(attachment)
                
                except Exception as e:
                    logger.warning(f"Error processing part {content_type}: {e}")
        
        else:
            # Not multipart
            content_type = msg.get_content_type()
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    charset = msg.get_content_charset() or 'utf-8'
                    decoded = self._decode_payload(payload, charset)
                    
                    if content_type == "text/html":
                        body_html = decoded
                    else:
                        body_plain = decoded
            except Exception as e:
                logger.warning(f"Error extracting body: {e}")
        
        # If both body_plain and body_html are empty, check for RTF body
        # readpst stores RTF-only bodies as an attachment named "rtf-body.rtf"
        if not body_plain.strip() and not body_html.strip():
            body_plain, body_html, attachments = self._try_extract_rtf_body(
                attachments
            )
            # After RTF extraction we may have HTML with cid: image refs.
            # Scan attachments for images with content_ids and register
            # them as inline images so the cid: URLs can be resolved.
            if body_html:
                for att in attachments:
                    cid = getattr(att, 'content_id', None) or ''
                    if cid and att.content_type.startswith('image/'):
                        if cid not in inline_images:
                            inline_images[cid] = att
        
        return body_plain, body_html, attachments, inline_images
    
    def _try_extract_rtf_body(
        self,
        attachments: List[Attachment]
    ) -> Tuple[str, str, List[Attachment]]:
        """
        Check if an RTF body attachment exists and extract content from it.
        
        When readpst extracts emails from PST files where the body is stored
        in RTF format only, it saves the RTF body as an attachment called
        "rtf-body.rtf". This method detects that attachment, extracts
        text/HTML from it, and removes it from the attachment list.
        
        Returns:
            Tuple of (plain_text, html, filtered_attachments)
        """
        body_plain = ""
        body_html = ""
        remaining_attachments = []
        rtf_found = False
        
        for att in attachments:
            # Check for RTF body attachment (created by readpst)
            is_rtf_body = (
                att.filename and 
                att.filename.lower() == 'rtf-body.rtf' and
                att.content
            )
            # Also check for application/rtf content type without a real filename
            if not is_rtf_body:
                is_rtf_body = (
                    att.content_type == 'application/rtf' and
                    att.filename and
                    'rtf-body' in att.filename.lower() and
                    att.content
                )
            
            if is_rtf_body and not rtf_found:
                rtf_found = True
                logger.info(f"Found RTF body attachment: {att.filename} ({att.size} bytes)")
                try:
                    plain, html = convert_rtf_body(att.content)
                    if html:
                        body_html = html
                        logger.info("Extracted HTML from RTF body")
                    if plain:
                        body_plain = plain
                        logger.info("Extracted plain text from RTF body")
                    
                    if not plain and not html:
                        # Couldn't extract content - keep as attachment
                        logger.warning("Could not extract content from RTF body, keeping as attachment")
                        remaining_attachments.append(att)
                except Exception as e:
                    logger.warning(f"Failed to extract RTF body: {e}")
                    remaining_attachments.append(att)
            else:
                remaining_attachments.append(att)
        
        if not rtf_found:
            return body_plain, body_html, attachments
        
        return body_plain, body_html, remaining_attachments
    
    def _decode_payload(self, payload: bytes, charset: str) -> str:
        """Decode payload bytes to string with fallback encodings."""
        # Try declared charset first, then common fallbacks
        encodings = [charset, 'utf-8', 'cp1252', 'latin-1', 'ascii']
        
        for encoding in encodings:
            if not encoding:
                continue
            try:
                decoded = payload.decode(encoding)
                # If we decoded as UTF-8 but got replacement chars, try cp1252
                # This handles emails that claim UTF-8 but are actually Windows-1252
                if encoding == 'utf-8' and '\ufffd' in decoded:
                    try:
                        return payload.decode('cp1252')
                    except (UnicodeDecodeError, LookupError):
                        pass
                return decoded
            except (UnicodeDecodeError, LookupError):
                continue
        
        # Last resort: decode with replacement
        return payload.decode('utf-8', errors='replace')
    
    def _extract_attachment(self, part: email.message.Message) -> Optional[Attachment]:
        """Extract an attachment from a message part."""
        try:
            payload = part.get_payload(decode=True)
            if not payload:
                return None
            
            filename = part.get_filename()
            content_type = part.get_content_type()
            
            if filename:
                # Decode filename if needed
                if isinstance(filename, bytes):
                    filename = filename.decode('utf-8', errors='replace')
                
                # Ensure filename has an extension - if missing, add from content type
                if not Path(filename).suffix:
                    ext = self._get_extension_from_content_type(content_type)
                    filename = f"{filename}{ext}"
            else:
                # Generate filename from content type
                ext = content_type.split('/')[-1] if '/' in content_type else 'bin'
                filename = f"attachment.{ext}"
            
            return Attachment(
                filename=filename,
                content_type=content_type,
                content=payload,
                size=len(payload),
                content_id=part.get("Content-ID", "").strip('<>')
            )
        
        except Exception as e:
            logger.warning(f"Error extracting attachment: {e}")
            return None
    
    def _get_extension_from_content_type(self, content_type: str) -> str:
        """Get file extension from MIME content type."""
        type_to_ext = {
            'application/pdf': '.pdf',
            'application/msword': '.doc',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
            'application/vnd.ms-excel': '.xls',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
            'application/vnd.ms-powerpoint': '.ppt',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
            'image/jpeg': '.jpg',
            'image/png': '.png',
            'image/gif': '.gif',
            'image/bmp': '.bmp',
            'image/tiff': '.tiff',
            'text/plain': '.txt',
            'text/html': '.html',
            'text/csv': '.csv',
            'text/calendar': '.ics',
            'message/rfc822': '.eml',
            'application/vnd.ms-outlook': '.msg',
            'application/rtf': '.rtf',
            'application/zip': '.zip',
            'application/x-zip-compressed': '.zip',
            'audio/mpeg': '.mp3',
            'audio/wav': '.wav',
            'video/mp4': '.mp4',
            'application/octet-stream': '.bin',
        }
        return type_to_ext.get(content_type.lower(), '.bin')
