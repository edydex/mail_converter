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


class MailboxWriter:
    """
    Writes emails to various mailbox formats.
    
    Supports:
    - MBOX: Standard format, works everywhere
    - EML Folder: Individual .eml files in a folder
    - PST: Windows only, requires Outlook
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
        
        Uses Outlook COM automation to create PST and import emails.
        """
        result = WriteResult(success=False, output_path=output_path)
        
        if sys.platform != 'win32':
            result.errors.append("PST writing is only available on Windows")
            return result
        
        try:
            import win32com.client
            import pythoncom
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                
                # Create new PST file
                # Outlook will create the file when we add the store
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
                
                # Get or create the target folder
                root_folder = pst_store.GetRootFolder()
                
                # Create subfolder for emails
                try:
                    target_folder = root_folder.Folders.Add(folder_name)
                except:
                    # Folder might already exist
                    target_folder = root_folder.Folders[folder_name]
                
                # Import each EML
                total = len(eml_paths)
                for i, eml_path in enumerate(eml_paths):
                    try:
                        self._report_progress(i + 1, total, f"Importing {Path(eml_path).name}")
                        
                        # Read EML content
                        with open(eml_path, 'rb') as f:
                            eml_content = f.read()
                        
                        # Create mail item from EML
                        # We need to save as temp .msg then import, OR use OpenSharedItem
                        # OpenSharedItem can open EML files directly
                        mail_item = namespace.OpenSharedItem(os.path.abspath(eml_path))
                        
                        # Move/copy to target folder
                        mail_item.Move(target_folder)
                        
                        result.emails_written += 1
                        
                    except Exception as e:
                        result.warnings.append(f"Failed to import {eml_path}: {e}")
                        logger.warning(f"Failed to import {eml_path} to PST: {e}")
                
                # Remove the PST from Outlook (keeps the file)
                namespace.RemoveStore(root_folder)
                
                result.success = True
                
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
