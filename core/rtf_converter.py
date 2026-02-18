"""
RTF Converter Module

Extracts text and HTML content from RTF data.
Handles Outlook's RTF-encapsulated HTML format where the original HTML
body is wrapped inside RTF tags.
"""

import re
import logging
from typing import Tuple, Optional

logger = logging.getLogger(__name__)

# Try to import striprtf for plain text extraction
STRIPRTF_AVAILABLE = False
try:
    from striprtf.striprtf import rtf_to_text
    STRIPRTF_AVAILABLE = True
except ImportError:
    logger.warning("striprtf not available - RTF body extraction will be limited")


def extract_html_from_rtf(rtf_data: bytes) -> Optional[str]:
    """
    Extract HTML from Outlook's RTF-encapsulated HTML format.
    
    Outlook often stores email bodies as RTF with the original HTML embedded.
    The RTF header contains '\\fromhtml1' to indicate encapsulated HTML.
    The HTML content is stored in \\htmltag groups.
    
    Args:
        rtf_data: Raw RTF content as bytes
        
    Returns:
        Extracted HTML string, or None if not RTF-encapsulated HTML
    """
    try:
        # Decode RTF to string
        rtf_text = _decode_rtf_bytes(rtf_data)
        if not rtf_text:
            return None
        
        # Check if this is Outlook's RTF-encapsulated HTML
        if '\\fromhtml' not in rtf_text:
            return None
        
        logger.debug("Detected Outlook RTF-encapsulated HTML")
        
        # Method: Extract HTML from \htmltag groups
        html = _extract_encapsulated_html(rtf_text)
        if html and len(html.strip()) > 10:
            return html
        
        return None
        
    except Exception as e:
        logger.warning(f"Failed to extract HTML from RTF: {e}")
        return None


def extract_text_from_rtf(rtf_data: bytes) -> Optional[str]:
    """
    Extract plain text from RTF data.
    
    Args:
        rtf_data: Raw RTF content as bytes
        
    Returns:
        Extracted plain text, or None if extraction failed
    """
    if not STRIPRTF_AVAILABLE:
        # Fallback: basic RTF text extraction without striprtf
        return _basic_rtf_to_text(rtf_data)
    
    try:
        rtf_text = _decode_rtf_bytes(rtf_data)
        if not rtf_text:
            return None
        
        text = rtf_to_text(rtf_text)
        if text and text.strip():
            return text.strip()
        return None
        
    except Exception as e:
        logger.warning(f"striprtf extraction failed: {e}")
        # Fall back to basic extraction
        return _basic_rtf_to_text(rtf_data)


def convert_rtf_body(rtf_data: bytes) -> Tuple[str, str]:
    """
    Convert RTF body content to usable text/HTML.
    
    First tries to extract encapsulated HTML (best quality).
    Falls back to plain text extraction.
    
    Args:
        rtf_data: Raw RTF content as bytes
        
    Returns:
        Tuple of (plain_text, html) - one or both may be populated
    """
    plain_text = ""
    html = ""
    
    # First, try to extract encapsulated HTML (preserves formatting)
    html = extract_html_from_rtf(rtf_data) or ""
    
    # Also extract plain text as fallback
    plain_text = extract_text_from_rtf(rtf_data) or ""
    
    if not html and not plain_text:
        logger.warning("Could not extract any content from RTF body")
    
    return plain_text, html


def _decode_rtf_bytes(rtf_data: bytes) -> Optional[str]:
    """Decode RTF bytes to string, trying multiple encodings."""
    if not rtf_data:
        return None
    
    # RTF is typically ASCII with escape sequences for other chars
    for encoding in ['ascii', 'cp1252', 'utf-8', 'latin-1']:
        try:
            return rtf_data.decode(encoding, errors='replace')
        except (UnicodeDecodeError, LookupError):
            continue
    
    return rtf_data.decode('ascii', errors='replace')


def _extract_encapsulated_html(rtf_text: str) -> Optional[str]:
    """
    Extract HTML from Outlook's RTF-encapsulated format.
    
    In this format, the RTF contains special control words:
    - \\*\\htmltag<N> introduces HTML tags
    - Regular text between htmltag groups is HTML content
    - \\htmlrtf ... \\htmlrtf0 marks RTF-only content (to be skipped)
    
    This implements a simplified version of the de-encapsulation algorithm
    described in MS-OXRTFEX.
    """
    try:
        html_parts = []
        i = 0
        length = len(rtf_text)
        in_htmlrtf = False  # Track if we're in an \htmlrtf block (RTF-only, skip)
        brace_depth = 0
        htmlrtf_depth = 0
        
        while i < length:
            ch = rtf_text[i]
            
            if ch == '{':
                brace_depth += 1
                i += 1
                continue
            elif ch == '}':
                if in_htmlrtf and brace_depth <= htmlrtf_depth:
                    in_htmlrtf = False
                brace_depth -= 1
                i += 1
                continue
            elif ch == '\\':
                # Control word
                ctrl, param, i = _parse_rtf_control(rtf_text, i)
                
                if ctrl == 'htmlrtf':
                    if param != '0':
                        in_htmlrtf = True
                        htmlrtf_depth = brace_depth
                    else:
                        in_htmlrtf = False
                elif ctrl == '*':
                    # Possible \*\htmltag
                    pass
                elif ctrl == 'htmltag':
                    # HTML tag content follows
                    pass
                elif ctrl == 'par' and not in_htmlrtf:
                    html_parts.append('\n')
                elif ctrl == 'tab' and not in_htmlrtf:
                    html_parts.append('\t')
                elif ctrl == 'line' and not in_htmlrtf:
                    html_parts.append('<br>')
                elif ctrl == 'lquote' and not in_htmlrtf:
                    html_parts.append('\u2018')
                elif ctrl == 'rquote' and not in_htmlrtf:
                    html_parts.append('\u2019')
                elif ctrl == 'ldblquote' and not in_htmlrtf:
                    html_parts.append('\u201c')
                elif ctrl == 'rdblquote' and not in_htmlrtf:
                    html_parts.append('\u201d')
                elif ctrl == 'emdash' and not in_htmlrtf:
                    html_parts.append('\u2014')
                elif ctrl == 'endash' and not in_htmlrtf:
                    html_parts.append('\u2013')
                elif ctrl == "'" and not in_htmlrtf:
                    # Hex character: \'XX
                    if len(param) >= 2:
                        try:
                            char_code = int(param[:2], 16)
                            html_parts.append(chr(char_code))
                        except ValueError:
                            pass
                elif ctrl == 'u' and not in_htmlrtf:
                    # Unicode character: \uN
                    try:
                        code_point = int(param)
                        if code_point < 0:
                            code_point += 65536
                        html_parts.append(chr(code_point))
                    except (ValueError, OverflowError):
                        pass
                continue
            elif not in_htmlrtf:
                # Regular character - part of HTML content
                if ch == '\r' or ch == '\n':
                    i += 1
                    continue
                html_parts.append(ch)
                i += 1
            else:
                i += 1
        
        result = ''.join(html_parts).strip()
        
        # Validate that we got something that looks like HTML
        if result and ('<' in result or '&' in result):
            # If it doesn't have html/body tags, wrap it
            if '<html' not in result.lower() and '<body' not in result.lower():
                result = f"<html><body>{result}</body></html>"
            return result
        
        return None
        
    except Exception as e:
        logger.warning(f"Failed to de-encapsulate HTML from RTF: {e}")
        return None


def _parse_rtf_control(rtf_text: str, pos: int) -> Tuple[str, str, int]:
    """
    Parse an RTF control word starting at pos (which points to the backslash).
    
    Returns:
        Tuple of (control_word, parameter, new_position)
    """
    i = pos + 1  # Skip the backslash
    length = len(rtf_text)
    
    if i >= length:
        return ('', '', i)
    
    ch = rtf_text[i]
    
    # Special single-character controls
    if ch == "'":
        # Hex escape: \'XX
        hex_str = rtf_text[i+1:i+3] if i + 2 < length else ''
        return ("'", hex_str, i + 3)
    
    if ch == '*':
        return ('*', '', i + 1)
    
    if ch == '\\':
        return ('\\', '', i + 1)
    
    if ch == '{':
        return ('{', '', i + 1)
    
    if ch == '}':
        return ('}', '', i + 1)
    
    if ch == '\r' or ch == '\n':
        return ('par', '', i + 1)
    
    # Read control word (alphabetic characters)
    ctrl_start = i
    while i < length and rtf_text[i].isalpha():
        i += 1
    
    ctrl_word = rtf_text[ctrl_start:i]
    
    if not ctrl_word:
        return ('', '', i)
    
    # Read optional numeric parameter (possibly negative)
    param_start = i
    if i < length and (rtf_text[i] == '-' or rtf_text[i].isdigit()):
        if rtf_text[i] == '-':
            i += 1
        while i < length and rtf_text[i].isdigit():
            i += 1
    
    param = rtf_text[param_start:i]
    
    # Skip single trailing space (delimiter)
    if i < length and rtf_text[i] == ' ':
        i += 1
    
    return (ctrl_word, param, i)


def _basic_rtf_to_text(rtf_data: bytes) -> Optional[str]:
    """
    Basic RTF to text extraction without external libraries.
    Strips RTF control words and extracts visible text.
    """
    try:
        rtf_text = _decode_rtf_bytes(rtf_data)
        if not rtf_text:
            return None
        
        # Remove RTF groups we don't care about
        # Remove {\*\...} destination groups
        text = re.sub(r'\{\\\*\\[^}]+\}', '', rtf_text)
        
        # Remove \' hex escapes - convert to characters
        def hex_to_char(match):
            try:
                return chr(int(match.group(1), 16))
            except (ValueError, OverflowError):
                return ''
        text = re.sub(r"\\'([0-9a-fA-F]{2})", hex_to_char, text)
        
        # Remove RTF control words
        text = re.sub(r'\\[a-z]+[-]?\d*\s?', '', text)
        
        # Remove remaining braces
        text = re.sub(r'[{}]', '', text)
        
        # Remove backslashes
        text = text.replace('\\', '')
        
        # Clean up whitespace
        text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)
        text = text.strip()
        
        return text if text else None
        
    except Exception as e:
        logger.warning(f"Basic RTF extraction failed: {e}")
        return None
