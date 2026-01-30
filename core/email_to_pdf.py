"""
Email to PDF Converter Module

Converts parsed email content to PDF format with proper formatting.
Uses WeasyPrint for high-fidelity HTML rendering with tables and CSS support.
Falls back to reportlab for plain text emails.
"""

import io
import os
import re
import sys
import base64
import tempfile
from pathlib import Path
from typing import Optional, List, Dict, Tuple
from datetime import datetime
import logging
from bs4 import BeautifulSoup

# Note: DPI awareness is now set in main.py BEFORE any imports
# This ensures it happens before tkinter creates any windows.
# We still set GTK/Cairo environment variables here as a safety measure.
if sys.platform == 'win32':
    # Set environment variables that GTK/Cairo/Pango use for rendering
    # These affect WeasyPrint's PDF output
    os.environ['GDK_SCALE'] = '1'
    os.environ['GDK_DPI_SCALE'] = '1'

# Check for WeasyPrint availability
# WeasyPrint requires GTK/GLib native libraries which may not be available on Windows
WEASYPRINT_AVAILABLE = False
WEASYPRINT_DEFAULT_URL_FETCHER = None
try:
    from weasyprint import HTML, CSS, default_url_fetcher
    from weasyprint.text.fonts import FontConfiguration
    WEASYPRINT_AVAILABLE = True
    WEASYPRINT_DEFAULT_URL_FETCHER = default_url_fetcher
except (ImportError, OSError, Exception):
    # ImportError: package not installed
    # OSError: native libraries (libgobject, etc.) not found on Windows
    pass

from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
    Table, TableStyle, PageBreak, HRFlowable
)
from reportlab.lib.utils import ImageReader
from PIL import Image

logger = logging.getLogger(__name__)


class EmailToPDFConverter:
    """
    Converts parsed email data to PDF format.
    Uses WeasyPrint for HTML emails with proper table/CSS support.
    Falls back to reportlab for plain text.
    """
    
    def __init__(self, page_size=letter, page_margin: float = 0.5, load_remote_images: bool = False):
        """
        Initialize the email to PDF converter.
        
        Args:
            page_size: Page size (default: letter)
            page_margin: Page margin in inches (default: 0.5)
            load_remote_images: Whether to load images from the web (default: False for security)
        """
        self.page_size = page_size
        self.page_size_name = "Letter" if page_size == letter else "A4"
        self.page_margin = page_margin
        self.load_remote_images = load_remote_images
        self._setup_styles()
        
        if WEASYPRINT_AVAILABLE:
            logger.info("WeasyPrint available - HTML emails will render with full formatting")
            if not self.load_remote_images:
                logger.info("Remote image loading disabled for security")
        else:
            logger.warning("WeasyPrint not available - HTML emails will use simplified rendering")
    
    def _url_fetcher(self, url: str):
        """
        Custom URL fetcher for WeasyPrint that blocks remote URLs for security.
        Only allows data: URLs (embedded images) and local file: URLs.
        """
        if url.startswith('data:'):
            # Allow data: URLs (base64 embedded images)
            return WEASYPRINT_DEFAULT_URL_FETCHER(url)
        elif url.startswith('file://'):
            # Allow local file URLs
            return WEASYPRINT_DEFAULT_URL_FETCHER(url)
        elif url.startswith(('http://', 'https://')):
            # Block remote URLs - return empty result
            logger.debug(f"Blocked remote image: {url[:100]}")
            return {'string': b'', 'mime_type': 'image/png'}
        else:
            # For other URLs (relative paths, etc.), try the default fetcher
            try:
                return WEASYPRINT_DEFAULT_URL_FETCHER(url)
            except Exception:
                return {'string': b'', 'mime_type': 'image/png'}
    
    def _setup_styles(self):
        """Set up paragraph styles for the PDF (used for plain text fallback)."""
        self.styles = getSampleStyleSheet()
        
        # Header style for metadata
        self.styles.add(ParagraphStyle(
            name='EmailHeader',
            parent=self.styles['Normal'],
            fontSize=10,
            leading=14,
            textColor=colors.Color(0.3, 0.3, 0.3)
        ))
        
        # Subject style
        self.styles.add(ParagraphStyle(
            name='EmailSubject',
            parent=self.styles['Heading1'],
            fontSize=14,
            leading=18,
            spaceAfter=12,
            textColor=colors.Color(0.1, 0.1, 0.1)
        ))
        
        # Body style
        self.styles.add(ParagraphStyle(
            name='EmailBody',
            parent=self.styles['Normal'],
            fontSize=10,
            leading=14,
            spaceBefore=6,
            spaceAfter=6
        ))
        
        # Attachment header style
        self.styles.add(ParagraphStyle(
            name='AttachmentHeader',
            parent=self.styles['Heading2'],
            fontSize=12,
            leading=16,
            spaceBefore=20,
            spaceAfter=10,
            textColor=colors.Color(0.2, 0.2, 0.6)
        ))
    
    def convert_email_to_pdf(
        self,
        email_data,  # ParsedEmail from eml_parser
        output_path: Path,
        include_headers: bool = True
    ) -> Path:
        """
        Convert a parsed email to PDF.
        
        Args:
            email_data: ParsedEmail object from eml_parser
            output_path: Path for the output PDF
            include_headers: Whether to include email headers
            
        Returns:
            Path to the created PDF
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Use WeasyPrint for HTML content if available
        if WEASYPRINT_AVAILABLE and email_data.body_html:
            return self._convert_with_weasyprint(email_data, output_path, include_headers)
        else:
            return self._convert_with_reportlab(email_data, output_path, include_headers)
    
    def _convert_with_weasyprint(
        self,
        email_data,
        output_path: Path,
        include_headers: bool
    ) -> Path:
        """
        Convert email to PDF using WeasyPrint for full HTML/CSS support.
        Falls back to reportlab if WeasyPrint fails.
        """
        try:
            # Build complete HTML document
            html_content = self._build_html_document(email_data, include_headers)
            
            # Create PDF with WeasyPrint
            font_config = FontConfiguration()
            
            # Page size CSS with configurable margin
            margin_str = f"{self.page_margin}in"
            # Use explicit physical dimensions instead of keywords like "letter" or "A4"
            # This ensures consistent output regardless of screen DPI or system settings
            if self.page_size_name == "Letter":
                # Letter size: 8.5 x 11 inches
                page_size_css = f"@page {{ size: 8.5in 11in; margin: {margin_str}; }}"
            else:
                # A4 size: 210 x 297 mm
                page_size_css = f"@page {{ size: 210mm 297mm; margin: {margin_str}; }}"
            
            css = CSS(string=page_size_css, font_config=font_config)
            
            # Use custom url_fetcher to block remote images if setting is disabled
            if self.load_remote_images:
                html = HTML(string=html_content)
                html.write_pdf(str(output_path), stylesheets=[css], font_config=font_config)
            else:
                html = HTML(string=html_content, url_fetcher=self._url_fetcher)
                html.write_pdf(str(output_path), stylesheets=[css], font_config=font_config)
            
            logger.info(f"Created PDF with WeasyPrint: {output_path}")
            return output_path
            
        except Exception as e:
            logger.warning(f"WeasyPrint conversion failed, falling back to reportlab: {e}")
            # Fall back to reportlab
            return self._convert_with_reportlab(email_data, output_path, include_headers)
    
    def _build_html_document(self, email_data, include_headers: bool) -> str:
        """
        Build a complete HTML document for the email with embedded images.
        """
        # Process HTML body - embed inline images as base64
        body_html = email_data.body_html or ""
        
        # Sanitize the email body HTML to remove conflicting styles
        body_html = self._sanitize_email_html(body_html)
        
        body_html = self._embed_inline_images(body_html, email_data.inline_images)
        
        # Build header section
        header_html = ""
        if include_headers:
            header_html = self._build_header_html(email_data)
        
        # Build attachment list
        attachment_html = ""
        if email_data.attachments:
            attachment_html = self._build_attachment_list_html(email_data)
        
        # Complete HTML document with CSS
        # IMPORTANT: Font consistency is critical for identical rendering.
        # Different machines may have different fonts installed, and Fontconfig
        # can substitute fonts with different metrics, causing text wrapping differences.
        # 
        # Solution: Use Segoe UI (Windows built-in since Vista) as primary font,
        # with Arial as fallback. These are present on ALL Windows systems.
        # For non-Windows, we fall back to system defaults.
        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        /* Reset any browser defaults and ensure consistent rendering */
        html {{
            font-size: 11pt;
        }}
        
        body {{
            /* Segoe UI is built into Windows and has consistent metrics.
               Arial is the universal fallback available everywhere. */
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 11pt;
            line-height: 1.4;
            color: #333;
            max-width: 100%;
            width: 100%;
            margin: 0;
            padding: 0;
        }}
        
        .email-header {{
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 4px;
            padding: 15px;
            margin-bottom: 20px;
        }}
        
        .email-subject {{
            font-size: 16pt;
            font-weight: bold;
            color: #1a1a1a;
            margin-bottom: 12px;
            word-wrap: break-word;
        }}
        
        .email-meta {{
            font-size: 10pt;
            color: #555;
            margin: 4px 0;
        }}
        
        .email-meta strong {{
            color: #333;
        }}
        
        .email-body {{
            margin-top: 20px;
            overflow-wrap: break-word;
            word-wrap: break-word;
        }}
        
        /* Force all content to fit within page width */
        * {{
            max-width: 100% !important;
            box-sizing: border-box;
        }}
        
        /* Tables - constrain to page width but preserve original styling */
        table {{
            border-collapse: collapse;
            max-width: 100% !important;
            width: auto !important;
            table-layout: auto;
            overflow: hidden;
        }}
        
        /* Only add padding to td/th if the table has explicit borders */
        /* Tables with border="0" or no border are layout tables - don't add padding */
        td, th {{
            vertical-align: top;
            overflow-wrap: break-word;
            word-wrap: break-word;
            word-break: break-word;
            overflow: hidden;
            width: auto !important;
        }}
        
        /* Style tables that explicitly request borders (border > 0) */
        /* Note: border="0" tables should remain invisible */
        table[border]:not([border="0"]) {{
            border: 1px solid #ddd;
        }}
        
        table[border]:not([border="0"]) td, 
        table[border]:not([border="0"]) th {{
            border: 1px solid #ddd;
            padding: 4px 8px;
        }}
        
        /* Images - respect explicit width/height from email, but cap at page width */
        img {{
            max-width: 100%;
            height: auto;
        }}
        
        /* Images without explicit sizing get a reasonable max */
        img:not([width]):not([style*="width"]) {{
            max-width: 300px;
        }}
        
        /* Links */
        a {{
            color: #0066cc;
            text-decoration: none;
        }}
        
        a:hover {{
            text-decoration: underline;
        }}
        
        /* Attachment list */
        .attachment-section {{
            margin-top: 30px;
            padding-top: 15px;
            border-top: 1px solid #ccc;
        }}
        
        .attachment-header {{
            font-size: 12pt;
            font-weight: bold;
            color: #333;
            margin-bottom: 10px;
        }}
        
        .attachment-list {{
            font-size: 10pt;
            color: #555;
        }}
        
        .attachment-item {{
            margin: 4px 0;
        }}
        
        /* Handle pre-formatted text */
        pre {{
            white-space: pre-wrap;
            word-wrap: break-word;
            font-family: DejaVu Sans Mono, Courier New, monospace;
            font-size: 10pt;
            background-color: #f5f5f5;
            padding: 10px;
            border-radius: 4px;
        }}
        
        /* Blockquotes - MINIMAL padding to prevent squeeze on long chains */
        blockquote {{
            margin: 4px 0 4px 0 !important;
            padding: 0 0 0 4px !important;
            border-left: 2px solid #ccc !important;
            color: #555;
        }}
        
        /* Nested blockquotes - no additional indent */
        blockquote blockquote {{
            margin-left: 0 !important;
            padding-left: 4px !important;
        }}
        
        blockquote blockquote blockquote,
        blockquote blockquote blockquote blockquote,
        blockquote blockquote blockquote blockquote blockquote {{
            margin-left: 0 !important;
            padding-left: 2px !important;
            border-left-width: 1px !important;
        }}
        
        /* Gmail-style quote divs - minimal indent */
        .gmail_quote, .gmail_extra, [class*="quote"], div[type="cite"] {{
            margin: 4px 0 4px 0 !important;
            padding: 0 0 0 4px !important;
            border-left: 2px solid #ccc !important;
            color: #555;
        }}
        
        /* Nested gmail quotes */
        .gmail_quote .gmail_quote,
        div[type="cite"] div[type="cite"] {{
            margin-left: 0 !important;
            padding-left: 4px !important;
        }}
        
        /* Outlook-style quoted text */
        .OutlookMessageHeader, .MsoNormal {{
            margin: 2px 0 !important;
        }}
        
        /* Prevent excessive nesting from squeezing content */
        .gmail_quote .gmail_quote,
        div[type="cite"] div[type="cite"] {{
            margin-left: 0;
            padding-left: 6px;
        }}
    </style>
</head>
<body>
    {header_html}
    <div class="email-body">
        {body_html}
    </div>
    {attachment_html}
</body>
</html>"""
        
        return html
    
    def _build_header_html(self, email_data) -> str:
        """Build HTML for email headers."""
        subject = self._escape_html(email_data.subject or "(No Subject)")
        sender = self._escape_html(f"{email_data.sender} <{email_data.sender_email}>")
        date_str = email_data.get_display_date()
        
        to_html = ""
        if email_data.recipients_to:
            to_list = self._escape_html(", ".join(email_data.recipients_to))
            to_html = f'<div class="email-meta"><strong>To:</strong> {to_list}</div>'
        
        cc_html = ""
        if email_data.recipients_cc:
            cc_list = self._escape_html(", ".join(email_data.recipients_cc))
            cc_html = f'<div class="email-meta"><strong>Cc:</strong> {cc_list}</div>'
        
        attachment_count_html = ""
        if email_data.attachments:
            count = len(email_data.attachments)
            attachment_count_html = f'<div class="email-meta"><strong>Attachments:</strong> {count} file(s)</div>'
        
        return f"""
<div class="email-header">
    <div class="email-subject">{subject}</div>
    <div class="email-meta"><strong>From:</strong> {sender}</div>
    {to_html}
    {cc_html}
    <div class="email-meta"><strong>Date:</strong> {date_str}</div>
    {attachment_count_html}
</div>
"""
    
    def _build_attachment_list_html(self, email_data) -> str:
        """Build HTML for attachment list."""
        items = []
        for i, attachment in enumerate(email_data.attachments, 1):
            size_kb = attachment.size / 1024
            size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
            filename = self._escape_html(attachment.filename)
            items.append(f'<div class="attachment-item">{i}. {filename} ({size_str})</div>')
        
        return f"""
<div class="attachment-section">
    <div class="attachment-header">Attachments:</div>
    <div class="attachment-list">
        {''.join(items)}
    </div>
</div>
"""

    def _fix_encoding_issues(self, text: str) -> str:
        """
        Fix common encoding issues in email text.
        
        Handles:
        - Windows-1252 "smart quotes" and special characters
        - Unicode replacement characters (U+FFFD / �)
        - Mojibake from incorrect encoding/decoding
        - Control characters that render as boxes
        """
        if not text:
            return ""
        
        # Map of problematic characters to their ASCII equivalents
        # This handles both the raw Windows-1252 bytes AND their Unicode equivalents
        char_replacements = {
            # Unicode replacement character - just remove it
            '\ufffd': '',
            '�': '',
            
            # Windows-1252 "smart" punctuation -> ASCII equivalents
            # Left/Right single quotes -> apostrophe
            '\u2018': "'",  # '
            '\u2019': "'",  # '
            '\u201a': "'",  # ‚ (single low-9 quote)
            
            # Left/Right double quotes -> straight quote
            '\u201c': '"',  # "
            '\u201d': '"',  # "
            '\u201e': '"',  # „ (double low-9 quote)
            
            # Dashes
            '\u2013': '-',  # en dash –
            '\u2014': '--', # em dash —
            '\u2015': '--', # horizontal bar
            
            # Ellipsis
            '\u2026': '...',  # …
            
            # Bullets and symbols
            '\u2022': '*',  # bullet •
            '\u2023': '>',  # triangular bullet
            '\u25aa': '*',  # small black square
            '\u25cf': '*',  # black circle
            '\u2219': '*',  # bullet operator
            
            # Trademark/copyright symbols (keep these as-is but ensure they work)
            '\u2122': '(TM)',  # ™
            '\u00a9': '(c)',   # ©
            '\u00ae': '(R)',   # ®
            
            # Non-breaking spaces and other whitespace
            '\u00a0': ' ',   # NBSP
            '\u00ad': '',    # soft hyphen - remove
            '\u200b': '',    # zero-width space
            '\u200c': '',    # zero-width non-joiner
            '\u200d': '',    # zero-width joiner
            '\ufeff': '',    # BOM / zero-width no-break space
            
            # Windows-1252 control characters (0x80-0x9F range)
            # These appear when text is incorrectly decoded
            '\x85': '...',  # ellipsis
            '\x91': "'",    # left single quote
            '\x92': "'",    # right single quote  
            '\x93': '"',    # left double quote
            '\x94': '"',    # right double quote
            '\x95': '*',    # bullet
            '\x96': '-',    # en dash
            '\x97': '--',   # em dash
            '\x99': '(TM)', # trademark
            
            # Fraction characters that may not render
            '\u00bc': '1/4',
            '\u00bd': '1/2',
            '\u00be': '3/4',
        }
        
        for bad_char, replacement in char_replacements.items():
            text = text.replace(bad_char, replacement)
        
        # Also handle the case where we have mangled multi-byte sequences
        # Common pattern: Â followed by special char (UTF-8 decoded as Latin-1)
        text = re.sub(r'Â\s*([''""•–—…])', r'\1', text)
        
        # Remove any remaining control characters (except newlines and tabs)
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
        
        return text
    
    def _sanitize_email_html(self, html_content: str) -> str:
        """
        Sanitize email HTML to remove conflicting CSS and problematic elements.
        
        This removes:
        - <style> tags (their CSS can conflict with our layout)
        - <link> tags (external stylesheets)
        - CSS pseudo-element content that renders as dots/icons
        - <head> section content
        - Windows-1252 control characters that render as dots
        - Decorative timeline/bullet div elements that don't contain text
        """
        if not html_content:
            return ""
        
        # First, normalize the text encoding
        html_content = self._fix_encoding_issues(html_content)
        
        # Remove decorative timeline divs (empty divs with border-radius that render as dots)
        # These are typically: <div style="...border-radius:10px..."></div>
        html_content = re.sub(
            r'<div[^>]*style="[^"]*border-radius[^"]*"[^>]*>\s*</div>',
            '',
            html_content,
            flags=re.IGNORECASE
        )
        
        # Remove decorative line divs (empty divs with just background-color for timeline lines)
        html_content = re.sub(
            r'<div[^>]*style="[^"]*(?:width:\s*[12]px|height:\s*\d+px)[^"]*background-color[^"]*"[^>]*>\s*</div>',
            '',
            html_content,
            flags=re.IGNORECASE
        )
        
        # Remove entire <head> section if present
        html_content = re.sub(r'<head[^>]*>.*?</head>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # Remove all <style> tags and their content
        html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # Remove <link> tags (external CSS)
        html_content = re.sub(r'<link[^>]*>', '', html_content, flags=re.IGNORECASE)
        
        # Remove <script> tags
        html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # Remove <meta> tags
        html_content = re.sub(r'<meta[^>]*>', '', html_content, flags=re.IGNORECASE)
        
        # Remove <title> tags
        html_content = re.sub(r'<title[^>]*>.*?</title>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # Strip out <html> and <body> opening/closing tags (we provide our own)
        html_content = re.sub(r'</?html[^>]*>', '', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</?body[^>]*>', '', html_content, flags=re.IGNORECASE)
        
        # Remove DOCTYPE
        html_content = re.sub(r'<!DOCTYPE[^>]*>', '', html_content, flags=re.IGNORECASE)
        
        # Remove XML declarations
        html_content = re.sub(r'<\?xml[^>]*\?>', '', html_content, flags=re.IGNORECASE)
        
        return html_content.strip()
    
    def _embed_inline_images(self, html_content: str, inline_images: Dict) -> str:
        """
        Replace cid: references with base64 embedded images.
        Also processes images and strips fixed widths to fit page.
        """
        # Always strip fixed widths from tables to prevent overflow
        html_content = self._strip_fixed_table_widths(html_content)
        
        if not inline_images:
            return self._constrain_images_without_dimensions(html_content)
        
        def replace_cid(match):
            cid = match.group(1)
            # Try with and without angle brackets
            attachment = inline_images.get(cid) or inline_images.get(f"<{cid}>")
            
            if attachment and attachment.content:
                try:
                    # Detect image type
                    content_type = attachment.content_type or 'image/png'
                    if not content_type.startswith('image/'):
                        content_type = 'image/png'
                    
                    # Convert to base64
                    b64_data = base64.b64encode(attachment.content).decode('utf-8')
                    return f'data:{content_type};base64,{b64_data}'
                except Exception as e:
                    logger.warning(f"Failed to embed image {cid}: {e}")
                    return match.group(0)
            
            return match.group(0)
        
        # Replace cid: references
        html_content = re.sub(
            r'cid:([^"\'>\s]+)',
            replace_cid,
            html_content,
            flags=re.IGNORECASE
        )
        
        # Constrain images that don't have explicit dimensions
        html_content = self._constrain_images_without_dimensions(html_content)
        
        return html_content
    
    def _strip_fixed_table_widths(self, html_content: str) -> str:
        """
        Use BeautifulSoup to remove all fixed pixel widths from tables and cells.
        This ensures text wraps properly on all platforms (especially Windows).
        """
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            removed_count = 0
            # Remove width attribute from tables, td, th, tr, tbody elements
            for tag in soup.find_all(['table', 'td', 'th', 'tr', 'tbody', 'thead', 'tfoot']):
                # Remove width attribute entirely
                if tag.has_attr('width'):
                    del tag['width']
                    removed_count += 1
                
                # Also strip width from inline styles
                if tag.has_attr('style'):
                    style = tag['style']
                    # Remove width declarations from style
                    new_style = re.sub(r'width\s*:\s*\d+(?:px)?\s*;?', '', style, flags=re.IGNORECASE)
                    if new_style != style:
                        removed_count += 1
                    if new_style.strip():
                        tag['style'] = new_style.strip()
                    else:
                        del tag['style']
            
            # Also handle divs with fixed widths that are used as containers
            for div in soup.find_all('div'):
                if div.has_attr('style'):
                    style = div['style']
                    # Only remove large fixed widths (>400px)
                    width_match = re.search(r'width\s*:\s*(\d+)px', style, re.IGNORECASE)
                    if width_match and int(width_match.group(1)) > 400:
                        style = re.sub(r'width\s*:\s*\d+px\s*;?', '', style, flags=re.IGNORECASE)
                        removed_count += 1
                        if style.strip():
                            div['style'] = style.strip()
                        else:
                            del div['style']
            
            logger.info(f"Stripped {removed_count} fixed width attributes from HTML")
            return str(soup)
        except Exception as e:
            logger.warning(f"Failed to strip table widths with BeautifulSoup: {e}")
            # Fallback to regex approach
            return self._strip_fixed_table_widths_regex(html_content)
    
    def _strip_fixed_table_widths_regex(self, html_content: str) -> str:
        """
        Fallback regex-based method to remove fixed widths.
        """
        # Remove width attributes from table-related tags
        html_content = re.sub(
            r'(<(?:table|td|th|tr|tbody|thead|tfoot)[^>]*)\s+width\s*=\s*["\']?\d+["\']?',
            r'\1',
            html_content,
            flags=re.IGNORECASE
        )
        return html_content
    
    def _constrain_images_without_dimensions(self, html_content: str) -> str:
        """
        Process image tags to ensure proper sizing in PDF output.
        
        - Images WITH width/height attributes: Convert to inline CSS so WeasyPrint honors them
        - Images WITHOUT dimensions: Add a reasonable max-width constraint
        
        This preserves the email author's intended image sizing (like scaled logos).
        """
        def process_image_tag(match):
            img_tag = match.group(0)
            
            # Extract width and height attributes if present (before modifying tag)
            width_match = re.search(r'\bwidth\s*=\s*["\']?(\d+)(?:px)?["\']?', img_tag, re.IGNORECASE)
            height_match = re.search(r'\bheight\s*=\s*["\']?(\d+)(?:px)?["\']?', img_tag, re.IGNORECASE)
            
            # Extract existing style attribute
            style_match = re.search(r'style\s*=\s*["\']([^"\']*)["\']', img_tag, re.IGNORECASE)
            existing_style = style_match.group(1) if style_match else ""
            
            # Build new styles based on what we found
            new_styles = []
            
            if width_match:
                width_val = int(width_match.group(1))
                # Cap at reasonable max for page width, but respect smaller values
                if width_val > 600:
                    width_val = 600
                new_styles.append(f"width: {width_val}px")
            
            if height_match:
                height_val = int(height_match.group(1))
                new_styles.append(f"height: {height_val}px")
            
            if not width_match and not height_match:
                # No explicit dimensions - constrain to reasonable size
                if 'max-width' not in existing_style.lower() and 'width' not in existing_style.lower():
                    new_styles.append("max-width: 200px")
                    new_styles.append("height: auto")
            
            # Also ensure images don't exceed page width
            if 'max-width' not in existing_style.lower():
                new_styles.append("max-width: 100%")
            
            if not new_styles:
                return img_tag
            
            # Remove width and height attributes from tag (we'll use CSS instead)
            img_tag = re.sub(r'\s*\bwidth\s*=\s*["\']?\d+(?:px)?["\']?', '', img_tag, flags=re.IGNORECASE)
            img_tag = re.sub(r'\s*\bheight\s*=\s*["\']?\d+(?:px)?["\']?', '', img_tag, flags=re.IGNORECASE)
            
            # Remove existing style attribute (we'll add a new combined one)
            img_tag = re.sub(r'\s*style\s*=\s*["\'][^"\']*["\']', '', img_tag, flags=re.IGNORECASE)
            
            # Build combined style
            combined_style = existing_style.rstrip(';')
            if combined_style:
                combined_style += "; "
            combined_style += "; ".join(new_styles)
            
            # Add the style attribute
            if img_tag.rstrip().endswith('/>'):
                img_tag = img_tag.rstrip()[:-2].rstrip() + f' style="{combined_style}" />'
            else:
                img_tag = img_tag.rstrip()[:-1].rstrip() + f' style="{combined_style}">'
            
            return img_tag
        
        # Match <img ...> tags
        html_content = re.sub(
            r'<img\s+[^>]*>',
            process_image_tag,
            html_content,
            flags=re.IGNORECASE | re.DOTALL
        )
        
        return html_content
    
    def _escape_html(self, text: str) -> str:
        """Escape text for HTML."""
        if not text:
            return ""
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        text = text.replace('"', '&quot;')
        return text
    
    def _convert_with_reportlab(
        self,
        email_data,
        output_path: Path,
        include_headers: bool
    ) -> Path:
        """
        Convert email to PDF using reportlab (fallback for plain text or when WeasyPrint unavailable).
        """
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=self.page_size,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=0.75*inch,
            bottomMargin=0.75*inch
        )
        
        story = []
        
        # Add email headers
        if include_headers:
            story.extend(self._create_header_section(email_data))
        
        # Add separator
        story.append(HRFlowable(
            width="100%",
            thickness=1,
            color=colors.Color(0.7, 0.7, 0.7),
            spaceBefore=10,
            spaceAfter=20
        ))
        
        # Add body content
        story.extend(self._create_body_section(email_data))
        
        # Add attachment list (not content, just list)
        if email_data.attachments:
            story.extend(self._create_attachment_list(email_data))
        
        # Build PDF
        doc.build(story)
        
        return output_path
    
    def _create_header_section(self, email_data) -> List:
        """Create the email header section."""
        elements = []
        
        # Subject
        subject = self._escape_text(email_data.subject or "(No Subject)")
        elements.append(Paragraph(f"<b>{subject}</b>", self.styles['EmailSubject']))
        
        # From
        sender = self._escape_text(f"{email_data.sender} <{email_data.sender_email}>")
        elements.append(Paragraph(f"<b>From:</b> {sender}", self.styles['EmailHeader']))
        
        # To
        if email_data.recipients_to:
            to_list = self._escape_text(", ".join(email_data.recipients_to))
            elements.append(Paragraph(f"<b>To:</b> {to_list}", self.styles['EmailHeader']))
        
        # CC
        if email_data.recipients_cc:
            cc_list = self._escape_text(", ".join(email_data.recipients_cc))
            elements.append(Paragraph(f"<b>Cc:</b> {cc_list}", self.styles['EmailHeader']))
        
        # Date
        date_str = email_data.get_display_date()
        elements.append(Paragraph(f"<b>Date:</b> {date_str}", self.styles['EmailHeader']))
        
        # Attachment count
        if email_data.attachments:
            count = len(email_data.attachments)
            elements.append(Paragraph(
                f"<b>Attachments:</b> {count} file(s)",
                self.styles['EmailHeader']
            ))
        
        elements.append(Spacer(1, 12))
        
        return elements
    
    def _create_body_section(self, email_data) -> List:
        """Create the email body section."""
        elements = []
        
        # Prefer HTML body if available (simplified conversion), otherwise use plain text
        if email_data.body_html:
            body_content = self._html_to_paragraphs(
                email_data.body_html,
                email_data.inline_images
            )
            elements.extend(body_content)
        elif email_data.body_plain:
            # Process plain text
            lines = email_data.body_plain.split('\n')
            for line in lines:
                escaped = self._escape_text(line)
                if escaped.strip():
                    elements.append(Paragraph(escaped, self.styles['EmailBody']))
                else:
                    elements.append(Spacer(1, 6))
        else:
            elements.append(Paragraph(
                "<i>(No message content)</i>",
                self.styles['EmailBody']
            ))
        
        return elements
    
    def _create_attachment_list(self, email_data) -> List:
        """Create a list of attachments section."""
        elements = []
        
        elements.append(Spacer(1, 20))
        elements.append(HRFlowable(
            width="100%",
            thickness=0.5,
            color=colors.Color(0.8, 0.8, 0.8),
            spaceBefore=10,
            spaceAfter=10
        ))
        
        elements.append(Paragraph("Attachments:", self.styles['AttachmentHeader']))
        
        for i, attachment in enumerate(email_data.attachments, 1):
            size_kb = attachment.size / 1024
            size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
            
            filename = self._escape_text(attachment.filename)
            elements.append(Paragraph(
                f"{i}. {filename} ({size_str})",
                self.styles['EmailHeader']
            ))
        
        return elements
    
    def _html_to_paragraphs(
        self,
        html_content: str,
        inline_images: Dict = None
    ) -> List:
        """
        Convert HTML content to reportlab paragraphs.
        
        This is a simplified HTML to PDF conversion that handles basic formatting.
        For complex HTML with tables, WeasyPrint is recommended.
        """
        elements = []
        inline_images = inline_images or {}
        
        # Strip HTML tags but preserve some structure
        # Remove style and script tags completely
        html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # Replace common HTML elements
        html_content = re.sub(r'<br\s*/?>', '\n', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</p>', '\n\n', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</div>', '\n', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</tr>', '\n', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</td>', '  |  ', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</th>', '  |  ', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'</li>', '\n', html_content, flags=re.IGNORECASE)
        
        # Handle inline images - collect them for later embedding
        images_to_embed = []
        img_pattern = r'<img[^>]+src=["\'](?:cid:)?([^"\']+)["\'][^>]*>'
        
        def replace_image(match):
            cid = match.group(1)
            if cid in inline_images:
                images_to_embed.append(inline_images[cid])
                return f"\n[IMAGE: {inline_images[cid].filename}]\n"
            return "[IMAGE]"
        
        html_content = re.sub(img_pattern, replace_image, html_content, flags=re.IGNORECASE)
        
        # Remove remaining HTML tags
        html_content = re.sub(r'<[^>]+>', '', html_content)
        
        # Decode HTML entities
        html_content = self._decode_html_entities(html_content)
        
        # Split into paragraphs and create elements
        paragraphs = html_content.split('\n')
        
        for para in paragraphs:
            para = para.strip()
            if para:
                escaped = self._escape_text(para)
                elements.append(Paragraph(escaped, self.styles['EmailBody']))
            else:
                elements.append(Spacer(1, 6))
        
        # Add inline images if available
        for attachment in images_to_embed:
            try:
                img = self._create_image_flowable(attachment.content)
                if img:
                    elements.append(Spacer(1, 12))
                    elements.append(img)
                    elements.append(Spacer(1, 12))
            except Exception as e:
                logger.warning(f"Could not embed inline image: {e}")
        
        return elements
    
    def _create_image_flowable(self, image_bytes: bytes, max_width: float = 6*inch, max_height: float = 4*inch):
        """Create a reportlab Image flowable from bytes."""
        try:
            img = Image.open(io.BytesIO(image_bytes))
            
            # Convert to RGB if necessary
            if img.mode in ('RGBA', 'P', 'LA'):
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode in ('RGBA', 'LA'):
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Calculate size to fit within bounds
            orig_width, orig_height = img.size
            
            # Scale to fit
            width_ratio = max_width / orig_width
            height_ratio = max_height / orig_height
            ratio = min(width_ratio, height_ratio, 1.0)  # Don't upscale
            
            new_width = orig_width * ratio
            new_height = orig_height * ratio
            
            # Save to buffer
            buffer = io.BytesIO()
            img.save(buffer, format='PNG')
            buffer.seek(0)
            
            return RLImage(buffer, width=new_width, height=new_height)
        
        except Exception as e:
            logger.warning(f"Error creating image flowable: {e}")
            return None
    
    def _escape_text(self, text: str) -> str:
        """Escape text for use in reportlab Paragraph."""
        if not text:
            return ""
        
        # Replace special characters
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        
        # Replace non-breaking spaces
        text = text.replace('\xa0', ' ')
        
        # Remove other control characters
        text = ''.join(char if ord(char) >= 32 or char in '\n\t' else ' ' for char in text)
        
        return text
    
    def _decode_html_entities(self, text: str) -> str:
        """Decode common HTML entities."""
        import html
        try:
            return html.unescape(text)
        except:
            # Manual fallback
            entities = {
                '&nbsp;': ' ',
                '&amp;': '&',
                '&lt;': '<',
                '&gt;': '>',
                '&quot;': '"',
                '&#39;': "'",
                '&apos;': "'",
                '&mdash;': '—',
                '&ndash;': '–',
                '&hellip;': '...',
            }
            for entity, char in entities.items():
                text = text.replace(entity, char)
            return text


def create_email_pdf(email_data, output_path: Path) -> Path:
    """
    Convenience function to create a PDF from parsed email.
    
    Args:
        email_data: ParsedEmail object
        output_path: Output PDF path
        
    Returns:
        Path to created PDF
    """
    converter = EmailToPDFConverter()
    return converter.convert_email_to_pdf(email_data, output_path)
