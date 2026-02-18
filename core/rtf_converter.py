"""
RTF Converter Module

Extracts HTML content from Outlook's RTF-encapsulated HTML format.
Implements the de-encapsulation algorithm described in [MS-OXRTFEX].

When Outlook stores an email whose original body was HTML, it wraps the
HTML inside RTF using special control words (``\\fromhtml1``, ``\\*\\htmltag``,
``\\htmlrtf`` … ``\\htmlrtf0``).  This module faithfully recovers the original
HTML so that it can be rendered with full formatting.

Hex escapes (\'XX) are decoded via cp1252, which is the default ANSI
code page for RTF produced by Outlook on Western-locale Windows.
"""

import re
import logging
from typing import Tuple, Optional, List

logger = logging.getLogger(__name__)

# Try to import striprtf for plain text extraction from non-encapsulated RTF
STRIPRTF_AVAILABLE = False
try:
    from striprtf.striprtf import rtf_to_text
    STRIPRTF_AVAILABLE = True
except ImportError:
    logger.warning("striprtf not available - RTF plain-text fallback will be limited")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def convert_rtf_body(rtf_data: bytes) -> Tuple[str, str]:
    """
    Convert RTF body content to usable text and/or HTML.

    1. If the RTF is Outlook-encapsulated HTML (\\fromhtml1), the original
       HTML is recovered with full styling.
    2. Otherwise plain text is extracted via *striprtf* (or a basic fallback).

    Returns:
        (plain_text, html) – one or both may be non-empty.
    """
    html = extract_html_from_rtf(rtf_data) or ""
    plain_text = ""

    if not html:
        # Not encapsulated HTML – fall back to plain text extraction
        plain_text = extract_text_from_rtf(rtf_data) or ""

    if not html and not plain_text:
        logger.warning("Could not extract any content from RTF body")

    return plain_text, html


def extract_html_from_rtf(rtf_data: bytes) -> Optional[str]:
    """Extract original HTML from Outlook RTF-encapsulated HTML."""
    try:
        rtf_text = _decode_rtf_bytes(rtf_data)
        if not rtf_text or '\\fromhtml' not in rtf_text:
            return None

        logger.debug("Detected Outlook RTF-encapsulated HTML")
        html = _deencapsulate_html(rtf_data)
        if html and len(html.strip()) > 10:
            return html
        return None
    except Exception as e:
        logger.warning(f"Failed to extract HTML from RTF: {e}")
        return None


def extract_text_from_rtf(rtf_data: bytes) -> Optional[str]:
    """Extract plain text from (non-encapsulated) RTF data."""
    if STRIPRTF_AVAILABLE:
        try:
            rtf_text = _decode_rtf_bytes(rtf_data)
            if rtf_text:
                text = rtf_to_text(rtf_text, errors="replace")
                if text and text.strip():
                    return text.strip()
        except Exception as e:
            logger.warning(f"striprtf failed: {e}")

    return _basic_rtf_to_text(rtf_data)


# ---------------------------------------------------------------------------
# Outlook RTF-encapsulated-HTML de-encapsulation
# ---------------------------------------------------------------------------

# Destination groups that should be skipped entirely – their textual content
# is RTF bookkeeping, not part of the email body.
_SKIP_DESTINATIONS = frozenset([
    'fonttbl', 'colortbl', 'stylesheet', 'listtable', 'listoverridetable',
    'info', 'xmlnstbl', 'latentstyles', 'datastore', 'themedata',
    'colorschememapping', 'fldinst', 'revtbl', 'bkmkstart', 'bkmkend',
    'pntext', 'pntxta', 'pntxtb',
    'pgdsctbl',      # page description table
    'generator',     # {\*\generator …}
    'mmathPr',       # Office math properties
    'passwordhash',
])

# Non-starred destination groups that also must be skipped.
# These appear as {\fonttbl ...} rather than {\*\fonttbl ...}.
_SKIP_NONSTAR_DESTINATIONS = frozenset([
    'fonttbl', 'colortbl', 'stylesheet', 'listtable', 'listoverridetable',
    'pgdsctbl', 'revtbl', 'info',
])


def _deencapsulate_html(rtf_data: bytes) -> Optional[str]:
    """
    Full RTF-encapsulated-HTML de-encapsulation.

    Walks the raw *bytes* so that \\' hex escapes are decoded through
    the correct code page (cp1252 by default, or as declared by \\ansicpgN).
    """
    try:
        data = rtf_data
        length = len(data)
        i = 0

        # State ----------------------------------------------------------
        codepage = 'cp1252'           # default ANSI code page
        html_parts: List[str] = []
        # Stack of group states: each entry = (in_htmlrtf, skip_group)
        group_stack: List[Tuple[bool, bool]] = []
        in_htmlrtf = False            # inside \htmlrtf … \htmlrtf0
        skip_group = False            # inside a destination to skip entirely
        uc_skip = 1                   # \ucN – chars to skip after \uN
        pending_skip = 0              # remaining chars to skip after a \uN

        def _emit(s: str):
            """Append *s* to the output if we're not in an RTF-only region."""
            if not in_htmlrtf and not skip_group:
                html_parts.append(s)

        while i < length:
            b = data[i]

            # ---- braces ------------------------------------------------
            if b == 0x7B:  # '{'
                # Push current state
                group_stack.append((in_htmlrtf, skip_group))
                i += 1

                # Peek ahead: check for \* destinations first
                dest = _peek_destination(data, i)
                if dest is not None:
                    dest_lower = dest.lower()
                    if dest_lower in _SKIP_DESTINATIONS:
                        skip_group = True
                    elif dest_lower == 'htmltag':
                        # \*\htmltag – the text inside IS HTML content
                        # Skip past the \*\htmltag control word itself
                        i = _skip_past_control(data, i)  # skip \*
                        i = _skip_past_control(data, i)  # skip \htmltag[N]
                        # remaining content in this group is HTML → keep
                    elif dest_lower == 'mhtmltag':
                        # Same idea as htmltag
                        i = _skip_past_control(data, i)
                        i = _skip_past_control(data, i)
                    else:
                        # Unknown \* destination – skip it to be safe
                        skip_group = True
                else:
                    # Not a \* group — check for non-starred destinations
                    # like {\fonttbl ...} or {\colortbl ...}
                    nonstar = _peek_nonstar_destination(data, i)
                    if nonstar and nonstar.lower() in _SKIP_NONSTAR_DESTINATIONS:
                        skip_group = True
                continue

            if b == 0x7D:  # '}'
                if group_stack:
                    in_htmlrtf, skip_group = group_stack.pop()
                i += 1
                continue

            # ---- control word / symbol ---------------------------------
            if b == 0x5C:  # '\\'
                ctrl, param_str, i = _parse_control_word(data, i)

                if ctrl == "'":
                    # ---- hex escape: \'XX → decode via codepage --------
                    try:
                        byte_val = int(param_str, 16)
                        char = bytes([byte_val]).decode(codepage, errors='replace')
                        _emit(char)
                    except (ValueError, OverflowError):
                        pass
                    continue

                if ctrl == 'u':
                    # ---- Unicode escape: \uN followed by uc_skip chars -
                    try:
                        cp = int(param_str)
                        if cp < 0:
                            cp += 65536
                        _emit(chr(cp))
                    except (ValueError, OverflowError):
                        pass
                    pending_skip = uc_skip
                    continue

                if ctrl == 'uc':
                    try:
                        uc_skip = int(param_str)
                    except ValueError:
                        pass
                    continue

                if ctrl == 'ansicpg':
                    try:
                        cpnum = int(param_str)
                        codepage = f'cp{cpnum}'
                        # Validate it
                        b'x'.decode(codepage)
                    except Exception:
                        codepage = 'cp1252'
                    continue

                if ctrl == 'htmlrtf':
                    # \htmlrtf  → start of RTF-only region
                    # \htmlrtf0 → end
                    in_htmlrtf = (param_str != '0')
                    continue

                if ctrl == 'par' and not in_htmlrtf and not skip_group:
                    _emit('\r\n')
                    continue
                if ctrl == 'line' and not in_htmlrtf and not skip_group:
                    _emit('<br>')
                    continue
                if ctrl == 'tab' and not in_htmlrtf and not skip_group:
                    _emit('\t')
                    continue
                if ctrl == 'lquote':
                    _emit('\u2018')
                    continue
                if ctrl == 'rquote':
                    _emit('\u2019')
                    continue
                if ctrl == 'ldblquote':
                    _emit('\u201c')
                    continue
                if ctrl == 'rdblquote':
                    _emit('\u201d')
                    continue
                if ctrl == 'emdash':
                    _emit('\u2014')
                    continue
                if ctrl == 'endash':
                    _emit('\u2013')
                    continue
                if ctrl == 'bullet':
                    _emit('\u2022')
                    continue
                if ctrl == '{':
                    _emit('{')
                    continue
                if ctrl == '}':
                    _emit('}')
                    continue
                if ctrl == '\\':
                    _emit('\\')
                    continue

                # All other control words are ignored (font switches,
                # paragraph formatting, etc.) – they are RTF-only.
                continue

            # ---- CR / LF → ignore (not meaningful in RTF) ---------------
            if b in (0x0D, 0x0A):
                i += 1
                continue

            # ---- regular byte -------------------------------------------
            if pending_skip > 0:
                pending_skip -= 1
                i += 1
                continue

            if not in_htmlrtf and not skip_group:
                # Decode this byte through the codepage
                try:
                    html_parts.append(bytes([b]).decode(codepage, errors='replace'))
                except Exception:
                    html_parts.append(chr(b) if b < 128 else '?')
            i += 1

        result = ''.join(html_parts).strip()
        if not result:
            return None

        # Validate we got something resembling HTML
        if '<' not in result and '&' not in result:
            return None

        return result

    except Exception as e:
        logger.warning(f"De-encapsulation failed: {e}")
        return None


def _peek_destination(data: bytes, pos: int) -> Optional[str]:
    """
    If the bytes starting at *pos* look like ``\\*\\word``, return *word*.
    Otherwise return ``None``.  Does NOT advance *pos*.
    """
    i = pos
    length = len(data)

    # skip whitespace
    while i < length and data[i] in (0x20, 0x0D, 0x0A):
        i += 1
    if i >= length or data[i] != 0x5C:  # '\\'
        return None
    i += 1
    if i >= length or data[i] != 0x2A:  # '*'
        return None
    i += 1
    # skip optional space
    while i < length and data[i] == 0x20:
        i += 1
    if i >= length or data[i] != 0x5C:  # next '\'
        return None
    i += 1

    # Read alphabetic control word name
    start = i
    while i < length and 0x61 <= (data[i] | 0x20) <= 0x7A:  # a-z / A-Z
        i += 1
    if i == start:
        return None
    return data[start:i].decode('ascii')


def _peek_nonstar_destination(data: bytes, pos: int) -> Optional[str]:
    """
    If the bytes at *pos* look like ``\\word`` (a control word immediately
    after an opening brace), return *word*.  Used to detect non-starred
    destinations like ``{\\fonttbl ...}`` or ``{\\colortbl ...}``.
    Does NOT advance *pos*.
    """
    i = pos
    length = len(data)

    # skip whitespace
    while i < length and data[i] in (0x20, 0x0D, 0x0A):
        i += 1
    if i >= length or data[i] != 0x5C:  # '\\'
        return None
    i += 1
    if i >= length:
        return None
    # Must be an alphabetic control word (not \' or \* etc.)
    if not (0x61 <= (data[i] | 0x20) <= 0x7A):
        return None

    start = i
    while i < length and 0x61 <= (data[i] | 0x20) <= 0x7A:
        i += 1
    if i == start:
        return None
    return data[start:i].decode('ascii')


def _skip_past_control(data: bytes, pos: int) -> int:
    """Skip past a single control word (including its numeric parameter and
    the trailing delimiter space) starting at *pos*."""
    i = pos
    length = len(data)

    # skip optional whitespace
    while i < length and data[i] in (0x20, 0x0D, 0x0A):
        i += 1
    if i >= length or data[i] != 0x5C:
        return i
    i += 1  # skip backslash

    if i >= length:
        return i

    # \* is a special two-char symbol
    if data[i] == 0x2A:
        i += 1
        # skip trailing space
        if i < length and data[i] == 0x20:
            i += 1
        return i

    # Read alpha control word
    while i < length and 0x61 <= (data[i] | 0x20) <= 0x7A:
        i += 1

    # Read optional numeric parameter (possibly negative)
    if i < length and (data[i] == 0x2D or 0x30 <= data[i] <= 0x39):
        if data[i] == 0x2D:
            i += 1
        while i < length and 0x30 <= data[i] <= 0x39:
            i += 1

    # Skip trailing delimiter space
    if i < length and data[i] == 0x20:
        i += 1

    return i


def _parse_control_word(data: bytes, pos: int) -> Tuple[str, str, int]:
    """
    Parse a control word starting at the backslash at *pos*.

    Returns (name, param_string, new_pos).  For ``\\'XX`` returns
    (``"'"``, ``"XX"``, pos).
    """
    i = pos + 1  # skip '\\'
    length = len(data)

    if i >= length:
        return ('', '', i)

    b = data[i]

    # \'XX hex escape
    if b == 0x27:  # "'"
        hex_str = ''
        if i + 2 < length:
            hex_str = chr(data[i + 1]) + chr(data[i + 2])
        return ("'", hex_str, i + 3)

    # Single-character symbols
    if b == 0x2A:  return ('*',  '', i + 1)
    if b == 0x5C:  return ('\\', '', i + 1)
    if b == 0x7B:  return ('{',  '', i + 1)
    if b == 0x7D:  return ('}',  '', i + 1)
    if b == 0x7E:  return ('~',  '', i + 1)   # non-breaking space
    if b == 0x5F:  return ('_',  '', i + 1)   # non-breaking hyphen
    if b in (0x0D, 0x0A):
        return ('par', '', i + 1)

    # Alphabetic control word
    start = i
    while i < length and 0x61 <= (data[i] | 0x20) <= 0x7A:
        i += 1
    ctrl = data[start:i].decode('ascii') if i > start else ''

    if not ctrl:
        return ('', '', i)

    # Optional numeric parameter (possibly negative)
    param_start = i
    if i < length and (data[i] == 0x2D or 0x30 <= data[i] <= 0x39):
        if data[i] == 0x2D:
            i += 1
        while i < length and 0x30 <= data[i] <= 0x39:
            i += 1
    param = data[param_start:i].decode('ascii')

    # Trailing delimiter space
    if i < length and data[i] == 0x20:
        i += 1

    return (ctrl, param, i)


# ---------------------------------------------------------------------------
# Plain-text fallback (non-encapsulated RTF)
# ---------------------------------------------------------------------------

def _decode_rtf_bytes(rtf_data: bytes) -> Optional[str]:
    """Decode RTF bytes to a Python string for text-level inspection."""
    if not rtf_data:
        return None
    for enc in ('ascii', 'cp1252', 'utf-8', 'latin-1'):
        try:
            return rtf_data.decode(enc, errors='replace')
        except (UnicodeDecodeError, LookupError):
            continue
    return rtf_data.decode('ascii', errors='replace')


def _basic_rtf_to_text(rtf_data: bytes) -> Optional[str]:
    """Very basic RTF-to-text fallback when *striprtf* is not installed."""
    try:
        rtf_text = _decode_rtf_bytes(rtf_data)
        if not rtf_text:
            return None

        # Strip destination groups ({\*\...})
        text = re.sub(r'\{\\\*\\[^{}]*\}', '', rtf_text)

        # Decode \'XX hex escapes through cp1252
        def _hex_char(m):
            try:
                return bytes([int(m.group(1), 16)]).decode('cp1252', errors='replace')
            except (ValueError, OverflowError):
                return ''
        text = re.sub(r"\\'([0-9a-fA-F]{2})", _hex_char, text)

        # Remove control words
        text = re.sub(r'\\[a-z]{1,32}(?:-?\d+)?\s?', '', text)

        # Remove braces
        text = re.sub(r'[{}]', '', text)

        # Remove stray backslashes
        text = text.replace('\\', '')

        # Collapse blank lines
        text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)
        text = text.strip()

        return text if text else None
    except Exception as e:
        logger.warning(f"Basic RTF extraction failed: {e}")
        return None
