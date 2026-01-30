"""
Mailbox Writer Module

Writes emails to various formats: MBOX, EML folder, PST (Windows only).
"""

import os
import sys
import mailbox
import logging
from pathlib import Path
from typing import List, Optional, Callable, Tuple
from dataclasses import dataclass, field
from enum import Enum
from email import policy
from email.generator import BytesGenerator
import io

logger = logging.getLogger(__name__)


class OutputFormat(Enum):
    """Supported output formats"""
    MBOX = "mbox"
    EML_FOLDER = "eml_folder"
    PST = "pst"  # Windows only


@dataclass
class WriteResult:
    """Result of a mailbox write operation"""
    success: bool
    output_path: str
    emails_written: int = 0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


def is_pst_write_available() -> bool:
    """
    Check if PST writing is available (Windows with Outlook).
    
    Returns:
        True if PST writing is supported, False otherwise
    """
    if sys.platform != 'win32':
        return False
    
    try:
        import win32com.client
        # Try to create Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        return True
    except Exception as e:
        logger.debug(f"PST writing not available: {e}")
        return False


def is_mapi_available() -> bool:
    """
    Check if Extended MAPI is available via pywin32.
    
    Extended MAPI allows setting message dates properly.
    Requires Windows with Outlook installed.
    
    Returns:
        True if Extended MAPI is available, False otherwise
    """
    if sys.platform != 'win32':
        return False
    
    try:
        from win32com.mapi import mapi, mapitags
        import win32com.client
        return True
    except Exception as e:
        logger.debug(f"Extended MAPI not available: {e}")
        return False


class MailboxWriter:
    """
    Writes emails to various mailbox formats.
    
    Supports:
    - MBOX: Standard format, works everywhere
    - EML Folder: Individual .eml files in a folder
    - PST: Windows only, requires Outlook (Extended MAPI for date preservation)
    """
    
    def __init__(
        self,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ):
        """
        Initialize the mailbox writer.
        
        Args:
            progress_callback: Optional callback(current, total, message)
        """
        self.progress_callback = progress_callback
        self._pst_available = None  # Lazy check
        self._mapi_available = None  # Lazy check
    
    def _report_progress(self, current: int, total: int, message: str):
        """Report progress to callback."""
        if self.progress_callback:
            self.progress_callback(current, total, message)
    
    @property
    def pst_available(self) -> bool:
        """Check if PST writing is available."""
        if self._pst_available is None:
            self._pst_available = is_pst_write_available()
        return self._pst_available
    
    @property
    def mapi_available(self) -> bool:
        """Check if Extended MAPI is available for PST writing with date preservation."""
        if self._mapi_available is None:
            self._mapi_available = is_mapi_available()
        return self._mapi_available
    
    def get_available_formats(self) -> List[OutputFormat]:
        """Get list of available output formats."""
        formats = [OutputFormat.MBOX, OutputFormat.EML_FOLDER]
        if self.pst_available:
            formats.append(OutputFormat.PST)
        return formats
    
    def write(
        self,
        eml_paths: List[str],
        output_path: str,
        output_format: OutputFormat,
        folder_name: str = "Emails"
    ) -> WriteResult:
        """
        Write emails to the specified format.
        
        Args:
            eml_paths: List of paths to EML files to write
            output_path: Output file/folder path
            output_format: Desired output format
            folder_name: Folder name within mailbox (for PST)
            
        Returns:
            WriteResult with details
        """
        if output_format == OutputFormat.MBOX:
            return self._write_mbox(eml_paths, output_path)
        elif output_format == OutputFormat.EML_FOLDER:
            return self._write_eml_folder(eml_paths, output_path)
        elif output_format == OutputFormat.PST:
            if not self.pst_available:
                return WriteResult(
                    success=False,
                    output_path=output_path,
                    errors=["PST writing requires Windows with Outlook installed"]
                )
            return self._write_pst(eml_paths, output_path, folder_name)
        else:
            return WriteResult(
                success=False,
                output_path=output_path,
                errors=[f"Unknown output format: {output_format}"]
            )
    
    def _write_mbox(self, eml_paths: List[str], output_path: str) -> WriteResult:
        """Write emails to MBOX format."""
        result = WriteResult(success=False, output_path=output_path)
        
        try:
            # Ensure output directory exists
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            
            # Create MBOX file
            mbox = mailbox.mbox(output_path)
            mbox.lock()
            
            try:
                total = len(eml_paths)
                for i, eml_path in enumerate(eml_paths):
                    try:
                        self._report_progress(i + 1, total, f"Writing {Path(eml_path).name}")
                        
                        # Read EML file as raw bytes and use compat32 policy
                        # to avoid MIME structure changes that confuse Outlook
                        with open(eml_path, 'rb') as f:
                            eml_content = f.read()
                        
                        # Use compat32 policy for maximum compatibility with email clients
                        from email import message_from_bytes
                        from email.policy import compat32
                        msg = message_from_bytes(eml_content, policy=compat32)
                        
                        # Fix common MIME issues that cause "body" attachment problem
                        msg = self._fix_mime_structure(msg)
                        
                        mbox.add(msg)
                        result.emails_written += 1
                        
                    except Exception as e:
                        result.warnings.append(f"Failed to add {eml_path}: {e}")
                        logger.warning(f"Failed to add {eml_path} to MBOX: {e}")
                
                mbox.flush()
                result.success = True
                
            finally:
                mbox.unlock()
                mbox.close()
        
        except Exception as e:
            result.errors.append(f"MBOX write failed: {e}")
            logger.error(f"MBOX write failed: {e}")
        
        return result
    
    def _fix_mime_structure(self, msg):
        """
        Fix MIME structure issues that cause body to appear as attachment.
        
        Common issues from readpst output:
        1. Body parts with Content-Disposition: attachment instead of inline
        2. Missing Content-Type on text parts
        3. Filename on body parts that shouldn't have one
        """
        if msg.is_multipart():
            # Check each part
            for part in msg.walk():
                content_type = part.get_content_type()
                disposition = part.get('Content-Disposition', '')
                
                # If it's a text/plain or text/html part that's marked as attachment
                # with a generic name like "body" or no name, fix it
                if content_type in ('text/plain', 'text/html'):
                    filename = part.get_filename()
                    
                    # Check if this looks like a body that got marked as attachment
                    if filename and filename.lower() in ('body', 'body.txt', 'body.html', 'body.htm'):
                        # Remove Content-Disposition to make it inline
                        if 'Content-Disposition' in part:
                            del part['Content-Disposition']
                        # Also remove the filename param if present
                        
                    elif 'attachment' in disposition.lower() and not filename:
                        # Body text marked as attachment with no filename - fix it
                        if 'Content-Disposition' in part:
                            del part['Content-Disposition']
        else:
            # Single-part message - ensure no weird disposition
            content_type = msg.get_content_type()
            if content_type in ('text/plain', 'text/html'):
                disposition = msg.get('Content-Disposition', '')
                filename = msg.get_filename()
                
                if filename and filename.lower() in ('body', 'body.txt', 'body.html', 'body.htm'):
                    if 'Content-Disposition' in msg:
                        del msg['Content-Disposition']
        
        return msg
    
    def _write_eml_folder(self, eml_paths: List[str], output_path: str) -> WriteResult:
        """Write emails to EML folder with proper naming (YYYYMMDD_HHMMSS_Subject.eml)."""
        result = WriteResult(success=False, output_path=output_path)
        
        try:
            # Create output directory
            output_dir = Path(output_path)
            output_dir.mkdir(parents=True, exist_ok=True)
            
            total = len(eml_paths)
            used_names = set()  # Track names to avoid collisions
            
            for i, eml_path in enumerate(eml_paths):
                try:
                    self._report_progress(i + 1, total, f"Processing {i+1}/{total}")
                    
                    # Read and parse the email to get date and subject
                    with open(eml_path, 'rb') as f:
                        eml_content = f.read()
                    
                    from email import message_from_bytes
                    from email.policy import compat32
                    from email.utils import parsedate_to_datetime
                    import re
                    
                    msg = message_from_bytes(eml_content, policy=compat32)
                    
                    # Get date
                    date_str = msg.get('Date', '')
                    try:
                        dt = parsedate_to_datetime(date_str)
                        date_prefix = dt.strftime('%Y%m%d_%H%M%S')
                    except:
                        # Fallback to index if date parsing fails
                        date_prefix = f"00000000_{i:06d}"
                    
                    # Get subject and sanitize for filename
                    subject = msg.get('Subject', '') or 'No Subject'
                    # Decode if needed
                    if hasattr(subject, 'encode'):
                        subject = str(subject)
                    
                    # Sanitize subject for filename
                    # Remove/replace invalid characters
                    subject = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', subject)
                    subject = re.sub(r'\s+', ' ', subject).strip()
                    # Truncate if too long (keep room for date prefix and extension)
                    max_subject_len = 100
                    if len(subject) > max_subject_len:
                        subject = subject[:max_subject_len].rsplit(' ', 1)[0] + '...'
                    
                    # Build filename
                    base_name = f"{date_prefix}_{subject}"
                    filename = f"{base_name}.eml"
                    
                    # Handle collisions
                    counter = 1
                    while filename.lower() in used_names:
                        filename = f"{base_name}_{counter}.eml"
                        counter += 1
                    
                    used_names.add(filename.lower())
                    dst = output_dir / filename
                    
                    # Write the file
                    with open(dst, 'wb') as f:
                        f.write(eml_content)
                    result.emails_written += 1
                    
                except Exception as e:
                    result.warnings.append(f"Failed to copy {eml_path}: {e}")
                    logger.warning(f"Failed to copy {eml_path}: {e}")
            
            result.success = True
            
        except Exception as e:
            result.errors.append(f"EML folder write failed: {e}")
            logger.error(f"EML folder write failed: {e}")
        
        return result
    
    def _write_pst(
        self, 
        eml_paths: List[str], 
        output_path: str,
        folder_name: str = "Emails"
    ) -> WriteResult:
        """
        Write emails to PST format (Windows only).
        
        Uses Extended MAPI for date preservation, falls back to Outlook COM.
        """
        result = WriteResult(success=False, output_path=output_path)
        
        if sys.platform != 'win32':
            result.errors.append("PST writing is only available on Windows")
            return result
        
        # Use Extended MAPI (preserves dates, requires Outlook running)
        if is_mapi_available():
            logger.info("Using Extended MAPI for PST writing (preserves dates)")
            return self._write_pst_mapi(eml_paths, output_path, folder_name)
        
        # Fall back to Outlook COM (dates not preserved)
        logger.info("MAPI not available, using standard Outlook COM (dates may not be preserved)")
        result.warnings.append(
            "Extended MAPI not available - email dates may show as today's date. "
            "Ensure Outlook is running for proper date preservation."
        )
        return self._write_pst_outlook(eml_paths, output_path, folder_name)
    
    def _write_pst_mapi(
        self, 
        eml_paths: List[str], 
        output_path: str,
        folder_name: str = "Emails"
    ) -> WriteResult:
        """
        Write emails to PST using Extended MAPI via pywin32.
        
        This method properly preserves sent/received dates using Extended MAPI.
        Requires Outlook to be running.
        """
        result = WriteResult(success=False, output_path=output_path)
        
        try:
            from win32com.mapi import mapi, mapitags
            import win32com.client
            import pythoncom
            import pywintypes
            from email import message_from_bytes, policy as email_policy
            from email.utils import parsedate_to_datetime, getaddresses
            import time
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            try:
                # Initialize MAPI
                mapi.MAPIInitialize((mapi.MAPI_INIT_VERSION, mapi.MAPI_MULTITHREAD_NOTIFICATIONS))
                
                # Connect to Outlook
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                
                # Get MAPI session
                session = mapi.MAPILogonEx(0, "", None, mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT)
                
                # Create PST file path
                pst_path = Path(output_path).resolve()
                pst_path.parent.mkdir(parents=True, exist_ok=True)
                
                # Remove existing file if present
                if pst_path.exists():
                    os.remove(str(pst_path))
                
                # Add PST to Outlook profile (creates the file)
                namespace.AddStore(str(pst_path))
                time.sleep(1)  # Let it initialize
                
                # Find the PST store we just created
                outlook_store = None
                for store in namespace.Stores:
                    if store.FilePath:
                        if Path(store.FilePath).resolve() == pst_path:
                            outlook_store = store
                            break
                
                if not outlook_store:
                    result.errors.append(f"Could not find PST store after creation: {pst_path}")
                    return result
                
                # Open store via MAPI using Outlook's StoreID
                pst_eid = bytes.fromhex(outlook_store.StoreID)
                mapi_store = session.OpenMsgStore(0, pst_eid, None, mapi.MDB_WRITE | mapi.MAPI_BEST_ACCESS)
                
                # Get IPM Subtree (root folder for mail)
                PR_IPM_SUBTREE_ENTRYID = 0x35E00102
                props = mapi_store.GetProps([PR_IPM_SUBTREE_ENTRYID], 0)
                
                if isinstance(props[0], tuple):
                    root_eid = props[0][1]
                elif isinstance(props, tuple) and len(props) >= 2:
                    root_eid = props[1][0][1]
                else:
                    root_eid = props[0]
                
                root_folder = mapi_store.OpenEntry(root_eid, None, mapi.MAPI_MODIFY | mapi.MAPI_BEST_ACCESS)
                
                # Create target folder
                try:
                    target_folder = root_folder.CreateFolder(1, folder_name, "Imported emails", None, 0)
                except Exception as e:
                    # Folder might exist, try to find it
                    table = root_folder.GetHierarchyTable(0)
                    table.SetColumns([0x0FFF0102, 0x3001001E], 0)  # PR_ENTRYID, PR_DISPLAY_NAME_A
                    rows = table.QueryRows(100, 0)
                    
                    target_folder = None
                    for row in rows:
                        fname = row[1][1]
                        if isinstance(fname, bytes):
                            fname = fname.decode('utf-8', errors='replace')
                        if fname == folder_name:
                            feid = row[0][1]
                            target_folder = mapi_store.OpenEntry(feid, None, mapi.MAPI_MODIFY | mapi.MAPI_BEST_ACCESS)
                            break
                    
                    if not target_folder:
                        raise RuntimeError(f"Could not create or find folder: {folder_name}")
                
                # Property tags for messages
                PR_MESSAGE_CLASS_A = mapitags.PR_MESSAGE_CLASS_A
                PR_SUBJECT_A = mapitags.PR_SUBJECT_A
                PR_BODY_A = mapitags.PR_BODY_A
                PR_HTML = 0x10130102
                PR_MESSAGE_FLAGS = mapitags.PR_MESSAGE_FLAGS
                PR_MESSAGE_DELIVERY_TIME = mapitags.PR_MESSAGE_DELIVERY_TIME
                PR_CLIENT_SUBMIT_TIME = mapitags.PR_CLIENT_SUBMIT_TIME
                PR_SENDER_NAME_A = mapitags.PR_SENDER_NAME_A
                PR_SENDER_EMAIL_A = mapitags.PR_SENDER_EMAIL_ADDRESS_A
                PR_SENT_REP_NAME_A = 0x0042001E
                PR_SENT_REP_EMAIL_A = 0x0065001E
                MSGFLAG_READ = 0x0001
                
                # Helper for ANSI string encoding
                def safe_ansi(s, max_len=None):
                    if not s:
                        return ""
                    result = s.encode('latin-1', errors='replace').decode('latin-1')
                    if max_len:
                        result = result[:max_len]
                    return result
                
                # Import each email
                total = len(eml_paths)
                imported_count = 0
                
                for i, eml_path in enumerate(eml_paths):
                    try:
                        self._report_progress(i + 1, total, f"Importing {i+1}/{total}")
                        
                        # Parse the email
                        with open(eml_path, 'rb') as f:
                            msg = message_from_bytes(f.read(), policy=email_policy.default)
                        
                        # Get email properties
                        subject = msg.get('Subject', '(No Subject)') or '(No Subject)'
                        from_header = msg.get('From', '')
                        date_str = msg.get('Date', '')
                        
                        # Parse sender
                        from_name = ''
                        from_email = from_header
                        addrs = getaddresses([from_header])
                        if addrs:
                            from_name, from_email = addrs[0]
                        
                        # Parse date - MUST be naive datetime for pywintypes
                        email_date = None
                        if date_str:
                            try:
                                email_date = parsedate_to_datetime(date_str)
                                if email_date.tzinfo is not None:
                                    email_date = email_date.replace(tzinfo=None)
                            except:
                                pass
                        
                        if not email_date:
                            from datetime import datetime
                            email_date = datetime.now()
                        
                        # Get body
                        body_plain = ''
                        body_html = ''
                        if msg.is_multipart():
                            for part in msg.walk():
                                ct = part.get_content_type()
                                if ct == 'text/plain' and not body_plain:
                                    try:
                                        body_plain = part.get_content()
                                    except:
                                        pass
                                elif ct == 'text/html' and not body_html:
                                    try:
                                        body_html = part.get_content()
                                    except:
                                        pass
                        else:
                            ct = msg.get_content_type()
                            try:
                                content = msg.get_content()
                                if ct == 'text/html':
                                    body_html = content
                                else:
                                    body_plain = content
                            except:
                                pass
                        
                        # Create message
                        mail = target_folder.CreateMessage(None, 0)
                        
                        # Convert date to PyTime
                        pytime = pywintypes.Time(email_date)
                        
                        # Set core properties
                        props = [
                            (PR_MESSAGE_CLASS_A, "IPM.Note"),
                            (PR_SUBJECT_A, safe_ansi(subject, 255)),
                            (PR_MESSAGE_FLAGS, MSGFLAG_READ),
                            (PR_MESSAGE_DELIVERY_TIME, pytime),
                            (PR_CLIENT_SUBMIT_TIME, pytime),
                        ]
                        
                        if from_email:
                            props.extend([
                                (PR_SENDER_NAME_A, safe_ansi(from_name or from_email)),
                                (PR_SENDER_EMAIL_A, safe_ansi(from_email)),
                                (PR_SENT_REP_NAME_A, safe_ansi(from_name or from_email)),
                                (PR_SENT_REP_EMAIL_A, safe_ansi(from_email)),
                            ])
                        
                        mail.SetProps(props)
                        
                        # Set body separately
                        body_props = []
                        if body_html:
                            body_props.append((PR_HTML, body_html.encode('utf-8')))
                        if body_plain:
                            body_props.append((PR_BODY_A, safe_ansi(body_plain)))
                        
                        if body_props:
                            mail.SetProps(body_props)
                        
                        # Save the message
                        mail.SaveChanges(0)
                        
                        result.emails_written += 1
                        imported_count += 1
                        
                    except Exception as e:
                        if len(result.warnings) < 10:
                            result.warnings.append(f"Failed to import email {i+1}: {str(e)[:100]}")
                        logger.warning(f"Failed to import {eml_path} to PST via MAPI: {e}")
                
                logger.info(f"Successfully imported {imported_count}/{total} emails to PST using Extended MAPI")
                
                if imported_count == 0 and total > 0:
                    result.errors.append(
                        f"Failed to import any emails. Try using MBOX or EML Folder output format instead."
                    )
                
                result.success = imported_count > 0
                
            finally:
                try:
                    mapi.MAPIUninitialize()
                except:
                    pass
                pythoncom.CoUninitialize()
        
        except Exception as e:
            result.errors.append(f"PST write with Extended MAPI failed: {e}")
            logger.error(f"PST write with Extended MAPI failed: {e}")
        
        return result
    
    def _write_pst_outlook(
        self, 
        eml_paths: List[str], 
        output_path: str,
        folder_name: str = "Emails"
    ) -> WriteResult:
        """
        Write emails to PST using standard Outlook COM.
        
        Note: This method cannot preserve original sent/received dates.
        Dates will show as the import time.
        """
        result = WriteResult(success=False, output_path=output_path)
        
        if sys.platform != 'win32':
            result.errors.append("PST writing is only available on Windows")
            return result
        
        try:
            import win32com.client
            import pythoncom
            from email import message_from_bytes, policy as email_policy
            from email.utils import parsedate_to_datetime
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                
                # Create new PST file
                pst_path = os.path.abspath(output_path)
                
                # Ensure directory exists
                Path(pst_path).parent.mkdir(parents=True, exist_ok=True)
                
                # Remove existing file if present
                if os.path.exists(pst_path):
                    os.remove(pst_path)
                
                # Add the PST to Outlook
                namespace.AddStore(pst_path)
                
                # Find the new store
                pst_store = None
                for store in namespace.Stores:
                    if store.FilePath.lower() == pst_path.lower():
                        pst_store = store
                        break
                
                if not pst_store:
                    result.errors.append("Failed to create PST store")
                    return result
                
                # Get root folder
                root_folder = pst_store.GetRootFolder()
                
                # Create subfolder for emails
                try:
                    target_folder = root_folder.Folders.Add(folder_name)
                except:
                    target_folder = root_folder.Folders[folder_name]
                
                # Import each email
                total = len(eml_paths)
                imported_count = 0
                
                for i, eml_path in enumerate(eml_paths):
                    try:
                        self._report_progress(i + 1, total, f"Importing {i+1}/{total}")
                        
                        # Parse the email file with Python's email module
                        with open(eml_path, 'rb') as f:
                            msg = message_from_bytes(f.read(), policy=email_policy.default)
                        
                        # Create new MailItem in Outlook
                        mail_item = outlook.CreateItem(0)  # 0 = olMailItem
                        
                        # Set basic properties
                        mail_item.Subject = msg.get('Subject', '(No Subject)') or '(No Subject)'
                        
                        # Set sender (display only - can't set actual sender on sent items)
                        sender = msg.get('From', '')
                        if sender:
                            mail_item.SentOnBehalfOfName = sender
                        
                        # Set recipients (To, CC, BCC)
                        to_addrs = msg.get('To', '')
                        cc_addrs = msg.get('Cc', '')
                        bcc_addrs = msg.get('Bcc', '')
                        
                        if to_addrs:
                            mail_item.To = to_addrs
                        if cc_addrs:
                            mail_item.CC = cc_addrs
                        if bcc_addrs:
                            mail_item.BCC = bcc_addrs
                        
                        # Set body
                        body = msg.get_body(preferencelist=('html', 'plain'))
                        if body:
                            content = body.get_content()
                            if body.get_content_type() == 'text/html':
                                mail_item.HTMLBody = content
                            else:
                                mail_item.Body = content
                        
                        # Set date
                        date_str = msg.get('Date', '')
                        if date_str:
                            try:
                                dt = parsedate_to_datetime(date_str)
                                # Note: SentOn is read-only, but we can set it via PropertyAccessor
                                # For now, the email will have current date - this is a limitation
                            except:
                                pass
                        
                        # Save and move to target folder
                        mail_item.Save()
                        mail_item.Move(target_folder)
                        
                        result.emails_written += 1
                        imported_count += 1
                        
                    except Exception as e:
                        error_msg = str(e)
                        # Only log first few errors to avoid spam
                        if len(result.warnings) < 10:
                            result.warnings.append(f"Failed to import email {i+1}: {error_msg[:100]}")
                        logger.warning(f"Failed to import {eml_path} to PST: {e}")
                
                logger.info(f"Successfully imported {imported_count}/{total} emails to PST")
                
                if imported_count == 0 and total > 0:
                    result.errors.append(
                        f"Failed to import any emails. Try using MBOX or EML Folder output format instead."
                    )
                
                # Remove the PST from Outlook (keeps the file)
                namespace.RemoveStore(root_folder)
                
                # Consider success if we imported at least some emails
                result.success = imported_count > 0
                
            finally:
                pythoncom.CoUninitialize()
        
        except ImportError as e:
            result.errors.append(f"win32com not available: {e}")
            logger.error(f"win32com not available: {e}")
        except Exception as e:
            result.errors.append(f"PST write failed: {e}")
            logger.error(f"PST write failed: {e}")
        
        return result
    
    def write_categorized(
        self,
        categories: dict[str, List[str]],
        output_dir: str,
        output_format: OutputFormat
    ) -> dict[str, WriteResult]:
        """
        Write multiple categories of emails to separate files/folders.
        
        Args:
            categories: Dict mapping category name to list of EML paths
            output_dir: Base output directory
            output_format: Desired output format
            
        Returns:
            Dict mapping category name to WriteResult
        """
        results = {}
        output_base = Path(output_dir)
        output_base.mkdir(parents=True, exist_ok=True)
        
        for category_name, eml_paths in categories.items():
            if not eml_paths:
                continue
            
            # Determine output path based on format
            if output_format == OutputFormat.MBOX:
                output_path = str(output_base / f"{category_name}.mbox")
            elif output_format == OutputFormat.EML_FOLDER:
                output_path = str(output_base / category_name)
            elif output_format == OutputFormat.PST:
                output_path = str(output_base / f"{category_name}.pst")
            else:
                continue
            
            self._report_progress(0, len(eml_paths), f"Writing {category_name}...")
            result = self.write(eml_paths, output_path, output_format, category_name)
            results[category_name] = result
        
        return results
