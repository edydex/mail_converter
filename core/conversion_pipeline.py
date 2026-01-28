"""
Conversion Pipeline Module

Orchestrates the complete PST/MBOX/MSG/EML to PDF conversion process.
Supports multiple input types and combining folders by name.
"""

import os
import shutil
import logging
from pathlib import Path
from typing import Optional, Callable, List, Dict, Any, Tuple
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
import json
import glob

from .pst_extractor import PSTExtractor, ExtractionResult
from .mbox_extractor import MBOXExtractor, MboxExtractionResult
from .msg_parser import MSGParser
from .eml_parser import EMLParser, ParsedEmail
from .email_to_pdf import EmailToPDFConverter
from .attachment_converter import AttachmentConverter, ConversionStatus
from .pdf_merger import PDFMerger, MergeResult
from .duplicate_detector import (
    DuplicateDetector, 
    DuplicateCertainty, 
    create_fingerprint_from_parsed_email
)

logger = logging.getLogger(__name__)


class PipelineStage(Enum):
    """Stages of the conversion pipeline"""
    INITIALIZING = "initializing"
    EXTRACTING_PST = "extracting_pst"
    PARSING_EMAILS = "parsing_emails"
    CONVERTING_EMAILS = "converting_emails"
    CONVERTING_ATTACHMENTS = "converting_attachments"
    MERGING_INDIVIDUAL = "merging_individual"
    MERGING_FINAL = "merging_final"
    COMPLETE = "complete"
    FAILED = "failed"


@dataclass
class PipelineProgress:
    """Progress information for the pipeline"""
    stage: PipelineStage
    current_item: int
    total_items: int
    current_item_name: str
    percentage: float
    message: str


class InputType(Enum):
    """Types of input files/folders"""
    PST = "pst"
    MBOX = "mbox"
    EML = "eml"
    MSG = "msg"
    EML_FOLDER = "eml_folder"
    PST_FOLDER = "pst_folder"
    MBOX_FOLDER = "mbox_folder"
    MIXED_FOLDER = "mixed_folder"


@dataclass
class PipelineConfig:
    """Configuration for the conversion pipeline"""
    # Input/Output - now supports multiple input paths
    pst_path: str  # Kept for backwards compatibility
    output_dir: str
    
    # New: List of input paths (files or folders)
    input_paths: List[str] = field(default_factory=list)
    input_type: Optional[InputType] = None  # Auto-detected if not specified
    
    # Duplicate detection
    detect_duplicates: bool = True
    duplicate_certainty: DuplicateCertainty = DuplicateCertainty.HIGH
    
    # Date filtering (for future use)
    date_from: Optional[datetime] = None
    date_to: Optional[datetime] = None
    
    # OCR
    ocr_enabled: bool = True
    
    # Output options
    keep_individual_pdfs: bool = True
    create_combined_pdf: bool = True
    add_toc: bool = True
    add_separators: bool = False  # Separator pages between emails
    add_att_separators: bool = False  # Separator pages before attachments
    page_size: str = "Letter"  # "Letter" or "A4"
    page_margin: float = 0.5  # Page margin in inches
    load_remote_images: bool = False  # Load images from the web (security concern)
    merge_folders: bool = False  # True = one combined PDF, False = separate PDF per folder
    
    # New: Combine folders by name across multiple PST/MBOX files
    combine_folders_by_name: bool = False
    
    # Processing
    preserve_pst_structure: bool = True
    rename_emls: bool = True  # Rename EMLs to YYYYMMDD_HHMMSS_subject.eml
    skip_deleted_items: bool = True  # Skip emails from "Deleted Items" folder
    
    def __post_init__(self):
        """Initialize input_paths from pst_path if not provided."""
        if not self.input_paths and self.pst_path:
            self.input_paths = [self.pst_path]


@dataclass
class PipelineResult:
    """Result of the conversion pipeline"""
    success: bool
    stage_reached: PipelineStage
    
    # Statistics
    emails_found: int = 0
    emails_processed: int = 0
    duplicates_skipped: int = 0
    attachments_converted: int = 0
    
    # Outputs
    individual_pdfs_dir: Optional[Path] = None
    combined_pdf_path: Optional[Path] = None
    log_path: Optional[Path] = None
    
    # Errors and warnings
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    # Timing
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    
    @property
    def duration_seconds(self) -> float:
        if self.start_time and self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return 0


class ConversionPipeline:
    """
    Main orchestrator for PST/MBOX/MSG/EML to PDF conversion.
    
    Handles the complete workflow:
    1. Extract/collect EMLs from various sources (PST, MBOX, MSG, EML files/folders)
    2. Parse EMLs and detect duplicates
    3. Convert emails to PDF
    4. Convert attachments to PDF
    5. Merge email + attachments into individual PDFs
    6. Merge all into chronological combined PDF
    
    Supports:
    - Multiple input files/folders
    - PST, MBOX, MSG, EML formats
    - Combining folders by name across multiple sources
    """
    
    def __init__(
        self,
        config: PipelineConfig,
        progress_callback: Optional[Callable[[PipelineProgress], None]] = None
    ):
        """
        Initialize the conversion pipeline.
        
        Args:
            config: Pipeline configuration
            progress_callback: Optional callback for progress updates
        """
        self.config = config
        self.progress_callback = progress_callback
        
        # Initialize components
        self.pst_extractor = PSTExtractor(
            progress_callback=self._pst_progress
        )
        self.mbox_extractor = MBOXExtractor()
        self.msg_parser = MSGParser()
        self.eml_parser = EMLParser()
        self.email_converter = EmailToPDFConverter(
            page_margin=config.page_margin,
            load_remote_images=config.load_remote_images
        )
        self.attachment_converter = AttachmentConverter(
            ocr_enabled=config.ocr_enabled,
            progress_callback=self._attachment_progress
        )
        self.pdf_merger = PDFMerger(
            progress_callback=self._merger_progress,
            page_size=config.page_size
        )
        
        # Duplicate detection - when merge_folders=False, use per-folder detectors
        self.detect_duplicates = config.detect_duplicates
        self.duplicate_certainty = config.duplicate_certainty
        
        if config.detect_duplicates and config.merge_folders:
            # Single global detector when merging all folders
            self.duplicate_detector = DuplicateDetector(
                min_certainty=config.duplicate_certainty
            )
            self.per_folder_detectors: Dict[str, DuplicateDetector] = {}
        elif config.detect_duplicates and not config.merge_folders:
            # Per-folder detectors when not merging folders
            self.duplicate_detector = None
            self.per_folder_detectors: Dict[str, DuplicateDetector] = {}
        else:
            self.duplicate_detector = None
            self.per_folder_detectors: Dict[str, DuplicateDetector] = {}
        
        # State
        self._current_stage = PipelineStage.INITIALIZING
        self._cancelled = False
        
        # Setup output directories
        self._setup_output_dirs()
    
    def _setup_output_dirs(self):
        """Set up output directory structure."""
        self.output_base = Path(self.config.output_dir)
        self.output_base.mkdir(parents=True, exist_ok=True)
        
        self.emls_dir = self.output_base / "1_extracted_emls"
        self.individual_pdfs_dir = self.output_base / "2_individual_pdfs"
        self.combined_dir = self.output_base / "3_combined_output"
        self.temp_dir = self.output_base / "_temp"
        
        # Create directories
        for d in [self.emls_dir, self.individual_pdfs_dir, self.combined_dir, self.temp_dir]:
            d.mkdir(exist_ok=True)
    
    def _report_progress(
        self,
        stage: PipelineStage,
        current: int,
        total: int,
        item_name: str,
        message: str
    ):
        """Report progress to callback."""
        self._current_stage = stage
        
        if not self.progress_callback:
            return
        
        # Calculate overall percentage
        stage_weights = {
            PipelineStage.INITIALIZING: (0, 5),
            PipelineStage.EXTRACTING_PST: (5, 20),
            PipelineStage.PARSING_EMAILS: (20, 30),
            PipelineStage.CONVERTING_EMAILS: (30, 50),
            PipelineStage.CONVERTING_ATTACHMENTS: (50, 70),
            PipelineStage.MERGING_INDIVIDUAL: (70, 85),
            PipelineStage.MERGING_FINAL: (85, 100),
            PipelineStage.COMPLETE: (100, 100),
            PipelineStage.FAILED: (0, 0)
        }
        
        start_pct, end_pct = stage_weights.get(stage, (0, 100))
        
        if total > 0:
            stage_progress = current / total
        else:
            stage_progress = 0
        
        overall_pct = start_pct + (end_pct - start_pct) * stage_progress
        
        progress = PipelineProgress(
            stage=stage,
            current_item=current,
            total_items=total,
            current_item_name=item_name,
            percentage=overall_pct,
            message=message
        )
        
        self.progress_callback(progress)
    
    def _pst_progress(self, current: int, total: int, message: str):
        """Progress callback for PST extraction."""
        self._report_progress(
            PipelineStage.EXTRACTING_PST,
            current, total, "", message
        )
    
    def _attachment_progress(self, message: str):
        """Progress callback for attachment conversion."""
        logger.info(message)
    
    def _merger_progress(self, current: int, total: int, message: str):
        """Progress callback for PDF merging."""
        self._report_progress(
            self._current_stage,
            current, total, "", message
        )
    
    def _detect_input_type(self, path: str) -> InputType:
        """
        Auto-detect the type of input from a file path.
        
        Args:
            path: Path to file or folder
            
        Returns:
            Detected InputType
        """
        p = Path(path)
        
        if p.is_dir():
            # Check what files are in the folder
            has_pst = any(p.glob("*.pst")) or any(p.glob("*.PST"))
            has_mbox = any(p.glob("*.mbox")) or any(p.glob("*.MBOX"))
            has_eml = any(p.glob("*.eml")) or any(p.glob("*.EML"))
            has_msg = any(p.glob("*.msg")) or any(p.glob("*.MSG"))
            
            if has_pst and not (has_mbox or has_eml or has_msg):
                return InputType.PST_FOLDER
            elif has_mbox and not (has_pst or has_eml or has_msg):
                return InputType.MBOX_FOLDER
            elif has_eml and not (has_pst or has_mbox):
                # EML folder may also have MSG files
                return InputType.EML_FOLDER
            elif has_msg and not (has_pst or has_mbox or has_eml):
                return InputType.EML_FOLDER  # Treat MSG-only folders as EML folders
            else:
                return InputType.MIXED_FOLDER
        else:
            # Single file
            ext = p.suffix.lower()
            if ext == ".pst":
                return InputType.PST
            elif ext == ".mbox":
                return InputType.MBOX
            elif ext == ".eml":
                return InputType.EML
            elif ext == ".msg":
                return InputType.MSG
            else:
                # Try to detect by content
                logger.warning(f"Unknown file type: {path}")
                return InputType.EML  # Default to EML
    
    def _extract_emails_from_inputs(self, result: PipelineResult) -> List[Tuple[Path, str]]:
        """
        Extract/collect EML files from all input sources.
        
        This method handles:
        - PST files: Extract using pst_extractor
        - MBOX files: Extract using mbox_extractor
        - MSG files: Convert to EML using msg_parser
        - EML files: Copy directly
        - Folders: Process contents based on detected type
        
        Args:
            result: PipelineResult to update with stats
            
        Returns:
            List of (eml_path, folder_name) tuples where folder_name is used
            for organizing output (used for combine_folders_by_name feature)
        """
        all_emls: List[Tuple[Path, str]] = []  # (eml_path, source_folder_name)
        
        for input_path in self.config.input_paths:
            if self._cancelled:
                break
            
            input_type = self._detect_input_type(input_path)
            logger.info(f"Processing input: {input_path} (detected type: {input_type.value})")
            
            if input_type == InputType.PST:
                emls = self._extract_from_pst(input_path, result)
                all_emls.extend(emls)
            
            elif input_type == InputType.MBOX:
                emls = self._extract_from_mbox(input_path, result)
                all_emls.extend(emls)
            
            elif input_type == InputType.MSG:
                eml = self._convert_msg_to_eml(input_path, result)
                if eml:
                    all_emls.append(eml)
            
            elif input_type == InputType.EML:
                eml = self._copy_eml(input_path, result)
                if eml:
                    all_emls.append(eml)
            
            elif input_type in (InputType.PST_FOLDER, InputType.MBOX_FOLDER, 
                               InputType.EML_FOLDER, InputType.MIXED_FOLDER):
                emls = self._process_folder(input_path, input_type, result)
                all_emls.extend(emls)
        
        return all_emls
    
    def _extract_from_pst(self, pst_path: str, result: PipelineResult) -> List[Tuple[Path, str]]:
        """Extract EMLs from a PST file."""
        pst_name = Path(pst_path).stem
        output_subdir = self.emls_dir / pst_name
        output_subdir.mkdir(parents=True, exist_ok=True)
        
        self._report_progress(
            PipelineStage.EXTRACTING_PST, 0, 1,
            pst_path, f"Extracting from PST: {pst_name}..."
        )
        
        extraction = self.pst_extractor.extract(
            pst_path,
            str(output_subdir),
            preserve_structure=self.config.preserve_pst_structure
        )
        
        if not extraction.success:
            result.errors.extend(extraction.errors)
            return []
        
        result.emails_found += extraction.eml_count
        result.warnings.extend(extraction.warnings)
        
        # Collect EML paths with folder info
        emls: List[Tuple[Path, str]] = []
        for eml_path in self.pst_extractor.get_extracted_emls(str(output_subdir)):
            # Get relative folder within PST
            try:
                rel_path = eml_path.relative_to(output_subdir)
                folder_name = str(rel_path.parent)
                if folder_name == ".":
                    folder_name = "Inbox"  # Default folder name
            except ValueError:
                folder_name = "Inbox"
            
            emls.append((eml_path, folder_name))
        
        return emls
    
    def _extract_from_mbox(self, mbox_path: str, result: PipelineResult) -> List[Tuple[Path, str]]:
        """Extract EMLs from an MBOX file."""
        mbox_name = Path(mbox_path).stem
        output_subdir = self.emls_dir / mbox_name
        output_subdir.mkdir(parents=True, exist_ok=True)
        
        self._report_progress(
            PipelineStage.EXTRACTING_PST, 0, 1,
            mbox_path, f"Extracting from MBOX: {mbox_name}..."
        )
        
        extraction_result = self.mbox_extractor.extract(
            mbox_path,
            str(output_subdir)
        )
        
        if not extraction_result.success:
            result.errors.extend(extraction_result.errors)
            return []
        
        result.emails_found += extraction_result.email_count
        result.warnings.extend(extraction_result.warnings)
        
        # For MBOX, use the filename as the folder name
        emls: List[Tuple[Path, str]] = []
        for eml_path in extraction_result.extracted_files:
            emls.append((Path(eml_path), mbox_name))
        
        return emls
    
    def _convert_msg_to_eml(self, msg_path: str, result: PipelineResult) -> Optional[Tuple[Path, str]]:
        """Convert a MSG file to EML format."""
        msg_name = Path(msg_path).stem
        
        self._report_progress(
            PipelineStage.EXTRACTING_PST, 0, 1,
            msg_path, f"Converting MSG: {msg_name}..."
        )
        
        try:
            # Convert MSG to EML
            eml_content = self.msg_parser.convert_to_eml(msg_path)
            if eml_content:
                eml_path = self.emls_dir / f"{msg_name}.eml"
                with open(eml_path, 'wb') as f:
                    f.write(eml_content)
                result.emails_found += 1
                return (eml_path, "")  # No folder for single MSG
        except Exception as e:
            result.warnings.append(f"Error converting MSG {msg_path}: {e}")
            logger.warning(f"Error converting MSG {msg_path}: {e}")
        
        return None
    
    def _copy_eml(self, eml_path: str, result: PipelineResult) -> Optional[Tuple[Path, str]]:
        """Copy an EML file to the working directory."""
        src = Path(eml_path)
        dest = self.emls_dir / src.name
        
        # Handle name collision
        counter = 1
        while dest.exists():
            dest = self.emls_dir / f"{src.stem}_{counter}{src.suffix}"
            counter += 1
        
        try:
            shutil.copy(src, dest)
            result.emails_found += 1
            return (dest, "")  # No folder for single EML
        except Exception as e:
            result.warnings.append(f"Error copying EML {eml_path}: {e}")
            logger.warning(f"Error copying EML {eml_path}: {e}")
            return None
    
    def _process_folder(
        self, 
        folder_path: str, 
        input_type: InputType, 
        result: PipelineResult
    ) -> List[Tuple[Path, str]]:
        """Process a folder containing email files."""
        folder = Path(folder_path)
        emls: List[Tuple[Path, str]] = []
        
        self._report_progress(
            PipelineStage.EXTRACTING_PST, 0, 1,
            folder_path, f"Processing folder: {folder.name}..."
        )
        
        # Process PST files
        for pst_file in list(folder.glob("*.pst")) + list(folder.glob("*.PST")):
            pst_emls = self._extract_from_pst(str(pst_file), result)
            emls.extend(pst_emls)
        
        # Process MBOX files
        for mbox_file in list(folder.glob("*.mbox")) + list(folder.glob("*.MBOX")):
            mbox_emls = self._extract_from_mbox(str(mbox_file), result)
            emls.extend(mbox_emls)
        
        # Process MSG files
        for msg_file in list(folder.glob("*.msg")) + list(folder.glob("*.MSG")):
            msg_result = self._convert_msg_to_eml(str(msg_file), result)
            if msg_result:
                emls.append(msg_result)
        
        # Process EML files - preserve folder structure
        for eml_file in list(folder.glob("**/*.eml")) + list(folder.glob("**/*.EML")):
            try:
                # Get relative path within the input folder
                rel_path = eml_file.relative_to(folder)
                folder_name = str(rel_path.parent)
                if folder_name == ".":
                    folder_name = folder.name
                
                # Copy to working directory preserving structure
                dest_folder = self.emls_dir / folder_name
                dest_folder.mkdir(parents=True, exist_ok=True)
                dest = dest_folder / eml_file.name
                
                # Handle name collision
                counter = 1
                while dest.exists():
                    dest = dest_folder / f"{eml_file.stem}_{counter}{eml_file.suffix}"
                    counter += 1
                
                shutil.copy(eml_file, dest)
                result.emails_found += 1
                emls.append((dest, folder_name))
            except Exception as e:
                result.warnings.append(f"Error copying EML {eml_file}: {e}")
                logger.warning(f"Error copying EML {eml_file}: {e}")
        
        return emls
    
    def cancel(self):
        """Cancel the pipeline execution."""
        self._cancelled = True
        logger.info("Pipeline cancellation requested")
    
    def run(self) -> PipelineResult:
        """
        Run the complete conversion pipeline.
        
        Returns:
            PipelineResult with all details
        """
        result = PipelineResult(
            success=False,
            stage_reached=PipelineStage.INITIALIZING,
            start_time=datetime.now()
        )
        
        try:
            # Stage 1: Extract/collect emails from all inputs
            self._report_progress(
                PipelineStage.EXTRACTING_PST, 0, 1, 
                "", "Extracting emails from input sources..."
            )
            
            # Get all EMLs with their folder information
            all_emls = self._extract_emails_from_inputs(result)
            
            if not all_emls:
                result.errors.append("No emails found in any input source")
                result.stage_reached = PipelineStage.EXTRACTING_PST
                return result
            
            if self._cancelled:
                return self._handle_cancellation(result)
            
            # Stage 2: Parse emails
            self._report_progress(
                PipelineStage.PARSING_EMAILS, 0, len(all_emls),
                "", "Parsing extracted emails..."
            )
            
            parsed_emails: List[Tuple[ParsedEmail, Path, str]] = []  # (parsed, path, folder_name)
            
            for i, (eml_path, source_folder) in enumerate(all_emls):
                if self._cancelled:
                    return self._handle_cancellation(result)
                
                self._report_progress(
                    PipelineStage.PARSING_EMAILS, i + 1, len(all_emls),
                    eml_path.name, f"Parsing {eml_path.name}..."
                )
                
                try:
                    email_data = self.eml_parser.parse_file(str(eml_path))
                    
                    # Rename EML file if enabled (for diagnostics)
                    if self.config.rename_emls:
                        new_name = f"{email_data.get_output_filename()}.eml"
                        new_path = eml_path.parent / new_name
                        counter = 1
                        while new_path.exists() and new_path != eml_path:
                            new_name = f"{email_data.get_output_filename()}_{counter}.eml"
                            new_path = eml_path.parent / new_name
                            counter += 1
                        if new_path != eml_path:
                            eml_path.rename(new_path)
                            eml_path = new_path
                    
                    # Determine effective folder path
                    # If combine_folders_by_name is enabled, use just the folder name (e.g., "Inbox")
                    # Otherwise, use the full source path (e.g., "mailbox1/Inbox")
                    if self.config.combine_folders_by_name and source_folder:
                        # Use only the last folder name (e.g., "Inbox" from "MailMeter/Inbox")
                        folder_path = Path(source_folder).name
                    else:
                        folder_path = source_folder
                    
                    # Check for duplicates
                    if self.detect_duplicates:
                        fp = create_fingerprint_from_parsed_email(email_data, str(eml_path))
                        
                        if self.duplicate_detector:
                            duplicate = self.duplicate_detector.add_email(fp)
                        else:
                            if folder_path not in self.per_folder_detectors:
                                self.per_folder_detectors[folder_path] = DuplicateDetector(
                                    min_certainty=self.duplicate_certainty
                                )
                            duplicate = self.per_folder_detectors[folder_path].add_email(fp)
                        
                        if duplicate:
                            result.duplicates_skipped += 1
                            logger.info(f"Skipping duplicate: {eml_path.name} ({duplicate.reason})")
                            continue
                    
                    # Skip "Deleted Items" folder if configured
                    if self.config.skip_deleted_items:
                        folder_lower = source_folder.lower()
                        folder_parts = folder_lower.replace("\\", "/").split("/")
                        is_deleted = (
                            "deleted items" in folder_lower or
                            "deleted" in folder_parts or
                            any(part.startswith("deleted") for part in folder_parts)
                        )
                        if is_deleted:
                            logger.info(f"Skipping email from Deleted Items: {eml_path.name}")
                            continue
                    
                    # Apply date filter
                    if self.config.date_from and email_data.date:
                        if email_data.date < self.config.date_from:
                            continue
                    if self.config.date_to and email_data.date:
                        if email_data.date > self.config.date_to:
                            continue
                    
                    parsed_emails.append((email_data, eml_path, folder_path))
                
                except Exception as e:
                    result.warnings.append(f"Error parsing {eml_path.name}: {e}")
                    logger.warning(f"Error parsing {eml_path}: {e}")
            
            result.stage_reached = PipelineStage.PARSING_EMAILS
            
            if not parsed_emails:
                result.warnings.append("No valid emails to process after filtering")
                return result
            
            # Stage 3-5: Convert emails and attachments, merge individual
            self._report_progress(
                PipelineStage.CONVERTING_EMAILS, 0, len(parsed_emails),
                "", "Converting emails to PDF..."
            )
            
            # Track PDFs by folder for separate combined PDFs option
            pdfs_by_folder: Dict[str, List[Tuple[Path, str]]] = {}
            individual_pdfs: List[Tuple[Path, str]] = []
            
            for i, (email_data, eml_path, folder_path) in enumerate(parsed_emails):
                if self._cancelled:
                    return self._handle_cancellation(result)
                
                output_name = email_data.get_output_filename()
                timestamp = email_data.get_timestamp_prefix()
                
                self._report_progress(
                    PipelineStage.CONVERTING_EMAILS, i + 1, len(parsed_emails),
                    output_name, f"Converting: {output_name}"
                )
                
                try:
                    # Convert email to PDF
                    email_pdf_path = self.temp_dir / f"{output_name}_email.pdf"
                    self.email_converter.convert_email_to_pdf(email_data, email_pdf_path)
                    
                    # Convert attachments
                    attachment_pdfs = []
                    
                    for j, attachment in enumerate(email_data.attachments):
                        self._report_progress(
                            PipelineStage.CONVERTING_ATTACHMENTS, 
                            i * 100 + j + 1, 
                            len(parsed_emails) * 100,
                            attachment.filename,
                            f"Converting attachment: {attachment.filename}"
                        )
                        
                        att_output_name = f"{output_name}_att{j+1:02d}_{attachment.filename}"
                        
                        conv_result = self.attachment_converter.convert_bytes(
                            attachment.content,
                            attachment.content_type,
                            attachment.filename,
                            str(self.temp_dir),
                            att_output_name
                        )
                        
                        if conv_result.status in (ConversionStatus.SUCCESS, ConversionStatus.PARTIAL):
                            is_placeholder = conv_result.status == ConversionStatus.PARTIAL
                            attachment_pdfs.append((conv_result.output_path, is_placeholder))
                            result.attachments_converted += 1
                        else:
                            result.warnings.append(
                                f"Attachment conversion failed: {attachment.filename} - {conv_result.message}"
                            )
                    
                    # Merge email + attachments
                    self._report_progress(
                        PipelineStage.MERGING_INDIVIDUAL, i + 1, len(parsed_emails),
                        output_name, f"Creating individual PDF: {output_name}"
                    )
                    
                    # Preserve folder structure in individual PDFs output
                    if folder_path:
                        pdf_output_folder = self.individual_pdfs_dir / folder_path
                        pdf_output_folder.mkdir(parents=True, exist_ok=True)
                    else:
                        pdf_output_folder = self.individual_pdfs_dir
                    
                    final_pdf_path = pdf_output_folder / f"{output_name}.pdf"
                    
                    if attachment_pdfs:
                        merge_result = self.pdf_merger.merge_email_with_attachments(
                            email_pdf_path,
                            attachment_pdfs,
                            final_pdf_path,
                            add_separators=self.config.add_att_separators
                        )
                        
                        if not merge_result.success:
                            result.warnings.extend(merge_result.errors)
                    else:
                        shutil.copy(email_pdf_path, final_pdf_path)
                    
                    individual_pdfs.append((final_pdf_path, timestamp))
                    
                    # Track by folder for per-folder combined PDFs
                    if folder_path not in pdfs_by_folder:
                        pdfs_by_folder[folder_path] = []
                    pdfs_by_folder[folder_path].append((final_pdf_path, timestamp))
                    
                    result.emails_processed += 1
                
                except Exception as e:
                    import traceback
                    tb = traceback.format_exc()
                    result.errors.append(f"Error processing {output_name}: {e}")
                    logger.error(f"Error processing email {output_name}:\n{tb}")
            
            result.stage_reached = PipelineStage.MERGING_INDIVIDUAL
            result.individual_pdfs_dir = self.individual_pdfs_dir
            
            # Stage 6: Merge into final PDF(s)
            if self.config.create_combined_pdf and individual_pdfs:
                if self.config.merge_folders:
                    # Merge ALL emails into one combined PDF
                    self._report_progress(
                        PipelineStage.MERGING_FINAL, 0, len(individual_pdfs),
                        "", "Creating combined PDF..."
                    )
                    
                    combined_path = self.combined_dir / "combined_chronological.pdf"
                    
                    merge_result = self.pdf_merger.merge_chronologically(
                        individual_pdfs,
                        combined_path,
                        add_toc=self.config.add_toc,
                        add_separators=self.config.add_separators
                    )
                    
                    if merge_result.success:
                        result.combined_pdf_path = combined_path
                    else:
                        result.warnings.extend(merge_result.errors)
                else:
                    # Create separate combined PDFs for each folder
                    total_folders = len(pdfs_by_folder)
                    for folder_idx, (folder_path, folder_pdfs) in enumerate(pdfs_by_folder.items()):
                        if not folder_pdfs:
                            continue
                        
                        # Create folder name for output
                        if folder_path:
                            safe_folder_name = folder_path.replace("/", "_").replace("\\", "_")
                            combined_filename = f"combined_{safe_folder_name}.pdf"
                        else:
                            combined_filename = "combined_root.pdf"
                        
                        self._report_progress(
                            PipelineStage.MERGING_FINAL, folder_idx + 1, total_folders,
                            folder_path or "root", f"Creating combined PDF for: {folder_path or 'root'}"
                        )
                        
                        combined_path = self.combined_dir / combined_filename
                        
                        merge_result = self.pdf_merger.merge_chronologically(
                            folder_pdfs,
                            combined_path,
                            add_toc=self.config.add_toc,
                            add_separators=self.config.add_separators
                        )
                        
                        if merge_result.success:
                            if result.combined_pdf_path is None:
                                result.combined_pdf_path = combined_path
                        else:
                            result.warnings.extend(merge_result.errors)
            
            result.stage_reached = PipelineStage.COMPLETE
            
            # Cleanup temp directory
            if not self.config.keep_individual_pdfs:
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            
            result.success = len(result.errors) == 0
            
            self._report_progress(
                PipelineStage.COMPLETE, 1, 1,
                "", f"Conversion complete! Processed {result.emails_processed} emails."
            )
        
        except Exception as e:
            result.errors.append(f"Pipeline error: {e}")
            logger.exception(f"Pipeline failed: {e}")
            result.stage_reached = PipelineStage.FAILED
        
        finally:
            result.end_time = datetime.now()
            self.attachment_converter.cleanup()
            result.log_path = self._write_log(result)
        
        return result
    
    def _handle_cancellation(self, result: PipelineResult) -> PipelineResult:
        """Handle pipeline cancellation."""
        result.errors.append("Pipeline was cancelled")
        result.end_time = datetime.now()
        return result
    
    def _write_log(self, result: PipelineResult) -> Path:
        """Write conversion log to file."""
        log_path = self.output_base / "conversion_log.txt"
        
        with open(log_path, 'w') as f:
            f.write("Mayo's Mail Converter - Conversion Log\n")
            f.write("=" * 50 + "\n\n")
            
            f.write("Input(s):\n")
            for input_path in self.config.input_paths:
                f.write(f"  - {input_path}\n")
            f.write(f"\nOutput: {self.config.output_dir}\n")
            f.write(f"Date: {result.start_time}\n")
            f.write(f"Duration: {result.duration_seconds:.1f} seconds\n")
            f.write(f"Combine folders by name: {self.config.combine_folders_by_name}\n\n")
            
            f.write("Statistics:\n")
            f.write(f"  - Emails found: {result.emails_found}\n")
            f.write(f"  - Emails processed: {result.emails_processed}\n")
            f.write(f"  - Duplicates skipped: {result.duplicates_skipped}\n")
            f.write(f"  - Attachments converted: {result.attachments_converted}\n\n")
            
            if result.errors:
                f.write("Errors:\n")
                for error in result.errors:
                    f.write(f"  - {error}\n")
                f.write("\n")
            
            if result.warnings:
                f.write("Warnings:\n")
                for warning in result.warnings:
                    f.write(f"  - {warning}\n")
        
        return log_path
