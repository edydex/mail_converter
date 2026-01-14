"""
Conversion Pipeline Module

Orchestrates the complete PST to PDF conversion process.
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

from .pst_extractor import PSTExtractor, ExtractionResult
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


@dataclass
class PipelineConfig:
    """Configuration for the conversion pipeline"""
    # Input/Output
    pst_path: str
    output_dir: str
    
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
    add_separators: bool = False
    page_size: str = "Letter"  # "Letter" or "A4"
    page_margin: float = 0.5  # Page margin in inches
    merge_folders: bool = False  # True = one combined PDF, False = separate PDF per folder
    
    # Processing
    preserve_pst_structure: bool = True
    rename_emls: bool = True  # Rename EMLs to YYYYMMDD_HHMMSS_subject.eml
    skip_deleted_items: bool = True  # Skip emails from "Deleted Items" folder


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
    Main orchestrator for PST to PDF conversion.
    
    Handles the complete workflow:
    1. Extract EMLs from PST
    2. Parse EMLs and detect duplicates
    3. Convert emails to PDF
    4. Convert attachments to PDF
    5. Merge email + attachments into individual PDFs
    6. Merge all into chronological combined PDF
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
        self.eml_parser = EMLParser()
        self.email_converter = EmailToPDFConverter(page_margin=config.page_margin)
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
            # Stage 1: Extract PST
            self._report_progress(
                PipelineStage.EXTRACTING_PST, 0, 1, 
                self.config.pst_path, "Extracting emails from PST..."
            )
            
            extraction = self.pst_extractor.extract(
                self.config.pst_path,
                str(self.emls_dir),
                preserve_structure=self.config.preserve_pst_structure
            )
            
            if not extraction.success:
                result.errors.extend(extraction.errors)
                result.stage_reached = PipelineStage.EXTRACTING_PST
                return result
            
            result.emails_found = extraction.eml_count
            result.warnings.extend(extraction.warnings)
            
            if self._cancelled:
                return self._handle_cancellation(result)
            
            # Stage 2: Parse emails
            self._report_progress(
                PipelineStage.PARSING_EMAILS, 0, extraction.eml_count,
                "", "Parsing extracted emails..."
            )
            
            eml_files = self.pst_extractor.get_extracted_emls(str(self.emls_dir))
            parsed_emails: List[Tuple[ParsedEmail, Path]] = []
            
            for i, eml_path in enumerate(eml_files):
                if self._cancelled:
                    return self._handle_cancellation(result)
                
                self._report_progress(
                    PipelineStage.PARSING_EMAILS, i + 1, len(eml_files),
                    eml_path.name, f"Parsing {eml_path.name}..."
                )
                
                try:
                    email_data = self.eml_parser.parse_file(str(eml_path))
                    
                    # Rename EML file if enabled (for diagnostics)
                    if self.config.rename_emls:
                        new_name = f"{email_data.get_output_filename()}.eml"
                        new_path = eml_path.parent / new_name
                        # Avoid overwriting if name collision
                        counter = 1
                        while new_path.exists() and new_path != eml_path:
                            new_name = f"{email_data.get_output_filename()}_{counter}.eml"
                            new_path = eml_path.parent / new_name
                            counter += 1
                        if new_path != eml_path:
                            eml_path.rename(new_path)
                            eml_path = new_path
                    
                    # Determine folder path for per-folder duplicate detection
                    try:
                        relative_eml_path = eml_path.relative_to(self.emls_dir)
                        folder_path = str(relative_eml_path.parent)
                        if folder_path == ".":
                            folder_path = ""
                    except ValueError:
                        folder_path = ""
                    
                    # Check for duplicates
                    if self.detect_duplicates:
                        fp = create_fingerprint_from_parsed_email(email_data, str(eml_path))
                        
                        if self.duplicate_detector:
                            # Global detector (merge_folders=True)
                            duplicate = self.duplicate_detector.add_email(fp)
                        else:
                            # Per-folder detector (merge_folders=False)
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
                        folder_lower = folder_path.lower()
                        folder_parts = folder_lower.replace("\\", "/").split("/")
                        is_deleted = (
                            "deleted items" in folder_lower or
                            "deleted" in folder_parts or
                            any(part.startswith("deleted") for part in folder_parts)
                        )
                        if is_deleted:
                            logger.info(f"Skipping email from Deleted Items: {eml_path.name}")
                            continue
                    
                    # Apply date filter (for future use)
                    if self.config.date_from and email_data.date:
                        if email_data.date < self.config.date_from:
                            continue
                    if self.config.date_to and email_data.date:
                        if email_data.date > self.config.date_to:
                            continue
                    
                    parsed_emails.append((email_data, eml_path))
                
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
            # Key: folder relative path, Value: list of (pdf_path, timestamp)
            pdfs_by_folder: Dict[str, List[Tuple[Path, str]]] = {}
            individual_pdfs: List[Tuple[Path, str]] = []  # (path, timestamp)
            
            for i, (email_data, eml_path) in enumerate(parsed_emails):
                if self._cancelled:
                    return self._handle_cancellation(result)
                
                output_name = email_data.get_output_filename()
                timestamp = email_data.get_timestamp_prefix()
                
                # Determine folder structure relative to emls_dir
                try:
                    relative_eml_path = eml_path.relative_to(self.emls_dir)
                    # Get parent folder(s) - e.g., "Exported MailMeter E-Mail/Inbox"
                    folder_path = str(relative_eml_path.parent)
                    if folder_path == ".":
                        folder_path = ""
                except ValueError:
                    folder_path = ""
                
                self._report_progress(
                    PipelineStage.CONVERTING_EMAILS, i + 1, len(parsed_emails),
                    output_name, f"Converting: {output_name}"
                )
                
                try:
                    # Convert email to PDF
                    email_pdf_path = self.temp_dir / f"{output_name}_email.pdf"
                    self.email_converter.convert_email_to_pdf(email_data, email_pdf_path)
                    
                    # Convert attachments
                    # Each item is (path, is_placeholder) - is_placeholder=True for unconverted files
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
                            # Track if this is a placeholder (PARTIAL = not converted, embedded)
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
                            add_separators=self.config.add_separators
                        )
                        
                        if not merge_result.success:
                            result.warnings.extend(merge_result.errors)
                    else:
                        # Just copy email PDF
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
                            # Sanitize folder path for filename
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
                            # Store first combined PDF path (or could store all)
                            if result.combined_pdf_path is None:
                                result.combined_pdf_path = combined_path
                        else:
                            result.warnings.extend(merge_result.errors)
            
            result.stage_reached = PipelineStage.COMPLETE
            
            # Cleanup temp directory
            if not self.config.keep_individual_pdfs:
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            
            # Write log file
            result.log_path = self._write_log(result)
            
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
            
            f.write(f"Input: {self.config.pst_path}\n")
            f.write(f"Output: {self.config.output_dir}\n")
            f.write(f"Date: {result.start_time}\n")
            f.write(f"Duration: {result.duration_seconds:.1f} seconds\n\n")
            
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
