"""
MSG Parser Module

Parses Microsoft Outlook .msg files and converts them to a standard format.
MSG files are proprietary Microsoft format used by Outlook for individual emails.
"""

import os
import logging
from pathlib import Path
from typing import Optional, Callable, List, Dict, Any
from dataclasses import dataclass, field
from datetime import datetime
import re

logger = logging.getLogger(__name__)

# Try to import RTF converter for RTF-only emails
try:
    from .rtf_converter import convert_rtf_body
    RTF_CONVERTER_AVAILABLE = True
except ImportError:
    RTF_CONVERTER_AVAILABLE = False

# Try to import extract_msg
MSG_AVAILABLE = False
try:
    import extract_msg
    from extract_msg import Message
    MSG_AVAILABLE = True
except ImportError:
    logger.warning("extract-msg not available - MSG support disabled")


@dataclass
class MSGAttachment:
    """Represents an attachment from an MSG file."""
    filename: str
    content: bytes
    content_type: str
    content_id: Optional[str] = None
    is_inline: bool = False


@dataclass
class ParsedMSG:
    """Parsed MSG file data."""
    subject: str
    sender: str
    sender_email: str
    recipients_to: List[str]
    recipients_cc: List[str]
    recipients_bcc: List[str]
    date: Optional[datetime]
    body_text: str
    body_html: str
    attachments: List[MSGAttachment]
    headers: Dict[str, str]
    
    # For compatibility with ParsedEmail
    inline_images: Dict[str, MSGAttachment] = field(default_factory=dict)
    
    def get_display_date(self) -> str:
        """Get formatted date string."""
        if self.date:
            return self.date.strftime("%Y-%m-%d %H:%M:%S")
        return "Unknown Date"


class MSGParser:
    """
    Parser for Microsoft Outlook .msg files.
    
    Uses the extract-msg library to parse MSG files.
    """
    
    def __init__(self):
        """Initialize the MSG parser."""
        if not MSG_AVAILABLE:
            logger.warning("MSG support not available - install extract-msg")
    
    @staticmethod
    def is_available() -> bool:
        """Check if MSG parsing is available."""
        return MSG_AVAILABLE
    
    def parse(self, msg_path: str) -> Optional[ParsedMSG]:
        """
        Parse an MSG file.
        
        Args:
            msg_path: Path to the MSG file
            
        Returns:
            ParsedMSG object or None if parsing failed
        """
        if not MSG_AVAILABLE:
            logger.error("MSG parsing not available - extract-msg not installed")
            return None
        
        if not os.path.isfile(msg_path):
            logger.error(f"MSG file not found: {msg_path}")
            return None
        
        try:
            msg = Message(msg_path)
            
            # Extract basic fields
            subject = msg.subject or "(No Subject)"
            sender = msg.sender or ""
            sender_email = msg.senderEmail or sender
            
            # Handle sender display name
            if sender_email and sender and sender != sender_email:
                sender = sender  # Keep display name
            elif sender_email:
                sender = sender_email
            
            # Get date
            date = None
            if msg.date:
                try:
                    if isinstance(msg.date, datetime):
                        date = msg.date
                    else:
                        # Try to parse date string
                        date = datetime.fromisoformat(str(msg.date))
                except Exception:
                    pass
            
            # Get recipients
            recipients_to = []
            recipients_cc = []
            recipients_bcc = []
            
            if msg.to:
                recipients_to = self._parse_recipients(msg.to)
            if msg.cc:
                recipients_cc = self._parse_recipients(msg.cc)
            if msg.bcc:
                recipients_bcc = self._parse_recipients(msg.bcc)
            
            # Get body
            body_text = msg.body or ""
            body_html = ""
            
            # Try to get HTML body
            if hasattr(msg, 'htmlBody') and msg.htmlBody:
                body_html = msg.htmlBody
            elif hasattr(msg, 'html') and msg.html:
                body_html = msg.html
            
            # If body_html is bytes, decode it
            if isinstance(body_html, bytes):
                try:
                    body_html = body_html.decode('utf-8', errors='replace')
                except Exception:
                    body_html = ""
            
            # If both text and HTML bodies are empty, try RTF body
            # Many Outlook emails store the body only in RTF format
            if not body_text.strip() and not body_html.strip():
                if RTF_CONVERTER_AVAILABLE:
                    rtf_data = None
                    try:
                        if hasattr(msg, 'rtfBody') and msg.rtfBody:
                            rtf_data = msg.rtfBody
                        elif hasattr(msg, 'compressedRtf') and msg.compressedRtf:
                            # Some versions expose compressed RTF
                            rtf_data = msg.compressedRtf
                    except Exception as e:
                        logger.debug(f"Could not access RTF body: {e}")
                    
                    if rtf_data:
                        if isinstance(rtf_data, str):
                            rtf_data = rtf_data.encode('utf-8')
                        try:
                            rtf_plain, rtf_html = convert_rtf_body(rtf_data)
                            if rtf_html:
                                body_html = rtf_html
                                logger.info("Extracted HTML from MSG RTF body")
                            if rtf_plain:
                                body_text = rtf_plain
                                logger.info("Extracted text from MSG RTF body")
                        except Exception as e:
                            logger.warning(f"Failed to extract RTF body: {e}")
            
            # Get attachments
            attachments = []
            inline_images = {}
            
            if msg.attachments:
                for att in msg.attachments:
                    try:
                        attachment = self._parse_attachment(att)
                        if attachment:
                            attachments.append(attachment)
                            
                            # Check if inline
                            if attachment.content_id:
                                inline_images[attachment.content_id] = attachment
                    except Exception as e:
                        logger.warning(f"Failed to parse attachment: {e}")
            
            # Build headers dict
            headers = {
                'Subject': subject,
                'From': f"{sender} <{sender_email}>" if sender != sender_email else sender_email,
                'To': ', '.join(recipients_to),
                'Cc': ', '.join(recipients_cc) if recipients_cc else '',
                'Date': date.isoformat() if date else '',
            }
            
            msg.close()
            
            return ParsedMSG(
                subject=subject,
                sender=sender,
                sender_email=sender_email,
                recipients_to=recipients_to,
                recipients_cc=recipients_cc,
                recipients_bcc=recipients_bcc,
                date=date,
                body_text=body_text,
                body_html=body_html,
                attachments=attachments,
                headers=headers,
                inline_images=inline_images
            )
            
        except Exception as e:
            logger.error(f"Failed to parse MSG file {msg_path}: {e}")
            return None
    
    def _parse_recipients(self, recipients) -> List[str]:
        """Parse recipient field to list of addresses."""
        if isinstance(recipients, str):
            # Split by comma or semicolon
            return [r.strip() for r in re.split(r'[;,]', recipients) if r.strip()]
        elif isinstance(recipients, list):
            return [str(r).strip() for r in recipients if r]
        return []
    
    def _parse_attachment(self, att) -> Optional[MSGAttachment]:
        """Parse an attachment from MSG."""
        try:
            filename = att.longFilename or att.shortFilename or "attachment"
            
            # Get content
            content = None
            if hasattr(att, 'data') and att.data:
                content = att.data
            elif hasattr(att, 'getFilename'):
                # Might need to read from temp file
                content = att.data
            
            if content is None:
                return None
            
            # Determine content type
            content_type = "application/octet-stream"
            if hasattr(att, 'mimetype') and att.mimetype:
                content_type = att.mimetype
            else:
                # Guess from filename
                ext = Path(filename).suffix.lower()
                content_type = self._guess_content_type(ext)
            
            # Get content ID for inline images
            content_id = None
            if hasattr(att, 'contentId') and att.contentId:
                content_id = att.contentId
            
            return MSGAttachment(
                filename=filename,
                content=content,
                content_type=content_type,
                content_id=content_id,
                is_inline=bool(content_id)
            )
            
        except Exception as e:
            logger.warning(f"Failed to parse attachment: {e}")
            return None
    
    def _guess_content_type(self, ext: str) -> str:
        """Guess MIME type from file extension."""
        mime_types = {
            '.pdf': 'application/pdf',
            '.doc': 'application/msword',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.xls': 'application/vnd.ms-excel',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.ppt': 'application/vnd.ms-powerpoint',
            '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.png': 'image/png',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.tif': 'image/tiff',
            '.tiff': 'image/tiff',
            '.txt': 'text/plain',
            '.html': 'text/html',
            '.htm': 'text/html',
            '.csv': 'text/csv',
            '.zip': 'application/zip',
            '.rar': 'application/x-rar-compressed',
            '.7z': 'application/x-7z-compressed',
        }
        return mime_types.get(ext.lower(), 'application/octet-stream')
    
    def convert_to_eml(self, msg_path: str, eml_path: str) -> bool:
        """
        Convert MSG file to EML format.
        
        Args:
            msg_path: Path to MSG file
            eml_path: Output path for EML file
            
        Returns:
            True if conversion successful
        """
        if not MSG_AVAILABLE:
            return False
        
        try:
            msg = Message(msg_path)
            
            # Build EML content
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            from email.mime.base import MIMEBase
            from email import encoders
            from email.utils import formatdate
            
            # Create message
            if msg.htmlBody or (msg.attachments and len(msg.attachments) > 0):
                email_msg = MIMEMultipart('mixed')
            else:
                email_msg = MIMEMultipart()
            
            # Set headers
            email_msg['Subject'] = msg.subject or ""
            email_msg['From'] = msg.sender or ""
            email_msg['To'] = msg.to or ""
            if msg.cc:
                email_msg['Cc'] = msg.cc
            if msg.date:
                try:
                    if isinstance(msg.date, datetime):
                        email_msg['Date'] = formatdate(msg.date.timestamp(), localtime=True)
                    else:
                        email_msg['Date'] = str(msg.date)
                except Exception:
                    email_msg['Date'] = formatdate(localtime=True)
            
            # Add body
            html_body_str = ""
            plain_body_str = msg.body or ""
            
            if msg.htmlBody:
                html_body_str = msg.htmlBody
                if isinstance(html_body_str, bytes):
                    html_body_str = html_body_str.decode('utf-8', errors='replace')
            
            # If both bodies are empty, try extracting from RTF
            if not plain_body_str.strip() and not html_body_str.strip():
                if RTF_CONVERTER_AVAILABLE:
                    rtf_data = None
                    try:
                        if hasattr(msg, 'rtfBody') and msg.rtfBody:
                            rtf_data = msg.rtfBody
                    except Exception:
                        pass
                    
                    if rtf_data:
                        if isinstance(rtf_data, str):
                            rtf_data = rtf_data.encode('utf-8')
                        try:
                            rtf_plain, rtf_html = convert_rtf_body(rtf_data)
                            if rtf_html:
                                html_body_str = rtf_html
                            if rtf_plain:
                                plain_body_str = rtf_plain
                        except Exception as e:
                            logger.warning(f"RTF body extraction failed in convert_to_eml: {e}")
            
            if html_body_str:
                # Create alternative part for text and HTML
                alt_part = MIMEMultipart('alternative')
                if plain_body_str:
                    alt_part.attach(MIMEText(plain_body_str, 'plain', 'utf-8'))
                alt_part.attach(MIMEText(html_body_str, 'html', 'utf-8'))
                email_msg.attach(alt_part)
            elif plain_body_str:
                email_msg.attach(MIMEText(plain_body_str, 'plain', 'utf-8'))
            
            # Add attachments
            if msg.attachments:
                for att in msg.attachments:
                    try:
                        filename = att.longFilename or att.shortFilename or "attachment"
                        if att.data:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(att.data)
                            encoders.encode_base64(part)
                            part.add_header(
                                'Content-Disposition',
                                f'attachment; filename="{filename}"'
                            )
                            email_msg.attach(part)
                    except Exception as e:
                        logger.warning(f"Failed to add attachment: {e}")
            
            msg.close()
            
            # Write to file
            Path(eml_path).parent.mkdir(parents=True, exist_ok=True)
            with open(eml_path, 'wb') as f:
                f.write(email_msg.as_bytes())
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to convert MSG to EML: {e}")
            return False


def msg_to_eml(msg_path: str, eml_path: str) -> bool:
    """
    Convenience function to convert MSG to EML.
    
    Args:
        msg_path: Path to MSG file
        eml_path: Output path for EML file
        
    Returns:
        True if conversion successful
    """
    parser = MSGParser()
    return parser.convert_to_eml(msg_path, eml_path)
