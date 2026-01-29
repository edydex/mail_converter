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


def is_redemption_available() -> bool:
    """
    Check if Redemption library is available for full PST write support.
    
    Redemption allows setting SentOn/ReceivedTime which standard Outlook COM cannot do.
    
    Returns:
        True if Redemption is available, False otherwise
    """
    if sys.platform != 'win32':
        return False
    
    try:
        import win32com.client
        # Try to create Redemption session
        session = win32com.client.Dispatch("Redemption.RDOSession")
        return True
    except Exception as e:
        logger.debug(f"Redemption not available: {e}")
        return False


class MailboxWriter:
    """
    Writes emails to various mailbox formats.
    
    Supports:
    - MBOX: Standard format, works everywhere
    - EML Folder: Individual .eml files in a folder
    - PST: Windows only, requires Outlook (with Redemption for date preservation)
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
        self._redemption_available = None  # Lazy check
    
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
    def redemption_available(self) -> bool:
        """Check if Redemption library is available for full PST date support."""
        if self._redemption_available is None:
            self._redemption_available = is_redemption_available()
        return self._redemption_available
    
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
                        
                        # Read EML file
                        with open(eml_path, 'rb') as f:
                            eml_content = f.read()
                        
                        # Parse and add to mbox
                        from email import message_from_bytes
                        msg = message_from_bytes(eml_content, policy=policy.default)
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
    
    def _write_eml_folder(self, eml_paths: List[str], output_path: str) -> WriteResult:
        """Write emails to EML folder."""
        result = WriteResult(success=False, output_path=output_path)
        
        try:
            # Create output directory
            output_dir = Path(output_path)
            output_dir.mkdir(parents=True, exist_ok=True)
            
            total = len(eml_paths)
            for i, eml_path in enumerate(eml_paths):
                try:
                    self._report_progress(i + 1, total, f"Copying {Path(eml_path).name}")
                    
                    src = Path(eml_path)
                    dst = output_dir / src.name
                    
                    # Handle name collision
                    counter = 1
                    while dst.exists():
                        dst = output_dir / f"{src.stem}_{counter}{src.suffix}"
                        counter += 1
                    
                    import shutil
                    shutil.copy2(src, dst)
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
        
        Uses Redemption if available (preserves dates), 
        otherwise falls back to Outlook COM (dates may not be preserved).
        """
        result = WriteResult(success=False, output_path=output_path)
        
        if sys.platform != 'win32':
            result.errors.append("PST writing is only available on Windows")
            return result
        
        # Try Redemption first (it preserves dates properly)
        if is_redemption_available():
            return self._write_pst_redemption(eml_paths, output_path, folder_name)
        else:
            logger.info("Redemption not available, using standard Outlook COM (dates may not be preserved)")
            result.warnings.append(
                "Redemption library not installed - email dates may show as today's date. "
                "Install Redemption from dimastr.com/redemption for proper date preservation."
            )
            return self._write_pst_outlook(eml_paths, output_path, folder_name)
    
    def _write_pst_redemption(
        self, 
        eml_paths: List[str], 
        output_path: str,
        folder_name: str = "Emails"
    ) -> WriteResult:
        """
        Write emails to PST using Redemption library.
        
        Redemption allows setting SentOn and ReceivedTime properly,
        which is not possible with standard Outlook COM.
        """
        result = WriteResult(success=False, output_path=output_path)
        
        try:
            import win32com.client
            import pythoncom
            from email import message_from_bytes, policy as email_policy
            from email.utils import parsedate_to_datetime
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            try:
                # Create Redemption session
                session = win32com.client.Dispatch("Redemption.RDOSession")
                
                # Connect to Outlook's MAPI session
                outlook = win32com.client.Dispatch("Outlook.Application")
                session.MAPIOBJECT = outlook.Session.MAPIOBJECT
                
                # Create new PST file
                pst_path = os.path.abspath(output_path)
                
                # Ensure directory exists
                Path(pst_path).parent.mkdir(parents=True, exist_ok=True)
                
                # Remove existing file if present
                if os.path.exists(pst_path):
                    os.remove(pst_path)
                
                # Create PST store
                pst_store = session.Stores.AddPSTStore(pst_path)
                
                # Get root folder and create target folder
                root_folder = pst_store.RootFolder
                
                try:
                    target_folder = root_folder.Folders.Add(folder_name)
                except:
                    # Folder might already exist
                    try:
                        target_folder = root_folder.Folders(folder_name)
                    except:
                        target_folder = root_folder
                
                # Import each email
                total = len(eml_paths)
                imported_count = 0
                
                for i, eml_path in enumerate(eml_paths):
                    try:
                        self._report_progress(i + 1, total, f"Importing {i+1}/{total}")
                        
                        # Parse the email file with Python's email module
                        with open(eml_path, 'rb') as f:
                            msg = message_from_bytes(f.read(), policy=email_policy.default)
                        
                        # Create new message in "sent" state
                        # CRITICAL: Set Sent=True BEFORE first save to allow date setting
                        mail_item = target_folder.Items.Add("IPM.Note")
                        mail_item.Sent = True
                        
                        # Set subject
                        mail_item.Subject = msg.get('Subject', '(No Subject)') or '(No Subject)'
                        
                        # Set sender
                        sender = msg.get('From', '')
                        if sender:
                            try:
                                mail_item.SentOnBehalfOfName = sender
                            except:
                                pass
                        
                        # Set recipients
                        to_addrs = msg.get('To', '')
                        cc_addrs = msg.get('Cc', '')
                        
                        if to_addrs:
                            try:
                                mail_item.To = to_addrs
                            except:
                                pass
                        if cc_addrs:
                            try:
                                mail_item.CC = cc_addrs
                            except:
                                pass
                        
                        # Set body
                        body = msg.get_body(preferencelist=('html', 'plain'))
                        if body:
                            try:
                                content = body.get_content()
                                if body.get_content_type() == 'text/html':
                                    mail_item.HTMLBody = content
                                else:
                                    mail_item.Body = content
                            except Exception as e:
                                logger.debug(f"Could not set body: {e}")
                        
                        # SET DATES - This is why we use Redemption!
                        date_str = msg.get('Date', '')
                        if date_str:
                            try:
                                dt = parsedate_to_datetime(date_str)
                                mail_item.SentOn = dt
                                mail_item.ReceivedTime = dt
                            except Exception as e:
                                logger.debug(f"Could not set date: {e}")
                        
                        # Save the message
                        mail_item.Save()
                        
                        result.emails_written += 1
                        imported_count += 1
                        
                    except Exception as e:
                        if len(result.warnings) < 10:
                            result.warnings.append(f"Failed to import email {i+1}: {str(e)[:100]}")
                        logger.warning(f"Failed to import {eml_path} to PST: {e}")
                
                logger.info(f"Successfully imported {imported_count}/{total} emails to PST using Redemption")
                
                if imported_count == 0 and total > 0:
                    result.errors.append(
                        f"Failed to import any emails. Try using MBOX or EML Folder output format instead."
                    )
                
                # Detach the PST store (keeps the file)
                try:
                    session.Stores.RemoveStore(pst_store)
                except:
                    pass
                
                result.success = imported_count > 0
                
            finally:
                pythoncom.CoUninitialize()
        
        except Exception as e:
            result.errors.append(f"PST write with Redemption failed: {e}")
            logger.error(f"PST write with Redemption failed: {e}")
        
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
