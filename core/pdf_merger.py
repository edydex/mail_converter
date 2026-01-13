"""
PDF Merger Module

Merges multiple PDFs into a single document.
Handles email PDFs with their attachment PDFs, and final chronological merge.

NOTE: Uses pikepdf for actual merging operations because PyPDF2 has issues
with page content becoming corrupted/duplicated during merge operations.
PyPDF2 is still used for reading page counts and extracting metadata.
"""

import os
import io
from pathlib import Path
from typing import List, Optional, Callable, Tuple
from dataclasses import dataclass
import logging

import pikepdf
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

logger = logging.getLogger(__name__)


@dataclass
class MergeItem:
    """Represents an item to be merged"""
    pdf_path: Path
    label: str  # For table of contents
    timestamp: str  # YYYYMMDD_HHMMSS for sorting
    is_attachment: bool = False
    parent_email: Optional[str] = None


@dataclass
class MergeResult:
    """Result of a merge operation"""
    success: bool
    output_path: Optional[Path]
    page_count: int
    errors: List[str]


class PDFMerger:
    """
    Merges PDF files with various options.
    Supports creating table of contents and section dividers.
    """
    
    def __init__(
        self, 
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        page_size: str = "Letter"
    ):
        """
        Initialize PDF Merger.
        
        Args:
            progress_callback: Optional callback(current, total, message)
            page_size: Target page size ("Letter" or "A4")
        """
        self.progress_callback = progress_callback
        self.page_size = letter if page_size == "Letter" else A4
        self.target_width = self.page_size[0]
        self.target_height = self.page_size[1]
    
    def _report_progress(self, current: int, total: int, message: str):
        """Report progress if callback is set."""
        if self.progress_callback:
            self.progress_callback(current, total, message)
    
    def _copy_embedded_files(self, src_pdf, dest_pdf):
        """
        Copy embedded files from source PDF to destination PDF.
        This preserves file attachments during merge operations.
        """
        try:
            # Check if source has embedded files
            if '/Root' not in src_pdf.trailer:
                return
            root = src_pdf.Root
            if '/Names' not in root:
                return
            names = root.Names
            if '/EmbeddedFiles' not in names:
                return
            
            embedded_files = names.EmbeddedFiles
            if '/Names' not in embedded_files:
                return
            
            src_names_array = list(embedded_files.Names)
            
            # Initialize destination Names/EmbeddedFiles if needed
            if '/Names' not in dest_pdf.Root:
                dest_pdf.Root['/Names'] = pikepdf.Dictionary()
            
            if '/EmbeddedFiles' not in dest_pdf.Root.Names:
                dest_pdf.Root.Names['/EmbeddedFiles'] = pikepdf.Dictionary({
                    '/Names': pikepdf.Array()
                })
            
            dest_names_array = list(dest_pdf.Root.Names.EmbeddedFiles.Names)
            
            # Copy each embedded file (array is [name1, filespec1, name2, filespec2, ...])
            for i in range(0, len(src_names_array), 2):
                if i + 1 < len(src_names_array):
                    name = src_names_array[i]
                    filespec = src_names_array[i + 1]
                    
                    # Make the filespec indirect in the destination PDF
                    # This properly copies all the file data
                    copied_filespec = dest_pdf.copy_foreign(filespec)
                    
                    dest_names_array.append(name)
                    dest_names_array.append(dest_pdf.make_indirect(copied_filespec))
            
            dest_pdf.Root.Names.EmbeddedFiles.Names = pikepdf.Array(dest_names_array)
            
        except Exception as e:
            logger.warning(f"Could not copy embedded files: {e}")
    
    def _add_toc_links(self, pdf, toc_page_count: int, toc_entries: List[Tuple[str, int]]):
        """
        Add clickable link annotations to TOC pages.
        
        Args:
            pdf: The pikepdf.Pdf object with TOC pages at the front
            toc_page_count: Number of TOC pages
            toc_entries: List of (title, page_number) tuples with adjusted page numbers
        """
        try:
            # TOC layout parameters (must match _create_table_of_contents)
            page_width = float(letter[0])
            left_margin = 0.75 * inch
            top_margin = inch
            title_height = 18 + 30  # fontSize + spaceAfter
            spacer_height = 20
            row_height = 18 + 8  # leading + padding
            
            # Starting Y position (from top of page)
            start_y = float(letter[1]) - top_margin - title_height - spacer_height
            
            # Entries per page (approximate)
            usable_height = float(letter[1]) - top_margin - inch  # bottom margin
            entries_per_page = int((usable_height - title_height - spacer_height) / row_height)
            
            # Add link annotations for each entry
            for idx, (title, target_page) in enumerate(toc_entries):
                # Which TOC page is this entry on?
                toc_page_idx = idx // entries_per_page
                if toc_page_idx >= toc_page_count:
                    toc_page_idx = toc_page_count - 1
                
                # Y position on this TOC page
                entry_on_page = idx % entries_per_page
                entry_y = start_y - (entry_on_page * row_height)
                
                # If this is not the first TOC page, adjust for no title
                if toc_page_idx > 0:
                    entry_y = float(letter[1]) - top_margin - (entry_on_page * row_height)
                
                # Link rectangle (covers the entry row)
                link_rect = pikepdf.Array([
                    left_margin,           # x1
                    entry_y - row_height,  # y1
                    page_width - left_margin,  # x2
                    entry_y                # y2
                ])
                
                # Create destination (page reference + fit)
                # target_page is 1-based, pikepdf pages are 0-based
                dest_page_idx = target_page - 1
                if dest_page_idx < len(pdf.pages):
                    dest_page = pdf.pages[dest_page_idx]
                    
                    # GoTo action with XYZ destination (go to top of page)
                    dest = pikepdf.Array([
                        dest_page.obj,
                        pikepdf.Name('/XYZ'),
                        0,    # left
                        float(letter[1]),  # top
                        0     # zoom (0 = inherit)
                    ])
                    
                    # Create link annotation
                    link_annot = pikepdf.Dictionary({
                        '/Type': pikepdf.Name('/Annot'),
                        '/Subtype': pikepdf.Name('/Link'),
                        '/Rect': link_rect,
                        '/Border': pikepdf.Array([0, 0, 0]),  # No visible border
                        '/Dest': dest,
                        '/H': pikepdf.Name('/I'),  # Invert on click
                    })
                    
                    # Add to TOC page
                    toc_page = pdf.pages[toc_page_idx]
                    if '/Annots' not in toc_page:
                        toc_page['/Annots'] = pikepdf.Array()
                    toc_page.Annots.append(pdf.make_indirect(link_annot))
                    
        except Exception as e:
            logger.warning(f"Could not add TOC links: {e}")
    
    def _scale_page_to_target(self, page) -> None:
        """
        Scale a PDF page to match the target page size.
        Modifies the page in place.
        """
        # Get current page dimensions
        media_box = page.mediabox
        current_width = float(media_box.width)
        current_height = float(media_box.height)
        
        # Calculate scale factors
        scale_x = self.target_width / current_width
        scale_y = self.target_height / current_height
        
        # Use the smaller scale to fit within target (maintain aspect ratio)
        scale = min(scale_x, scale_y)
        
        # Only scale if significantly different (more than 5% difference)
        if abs(1 - scale) > 0.05:
            # Scale the page
            page.scale(scale, scale)
            
            # Center on the target page size
            new_width = current_width * scale
            new_height = current_height * scale
            
            # Calculate offsets to center
            x_offset = (self.target_width - new_width) / 2
            y_offset = (self.target_height - new_height) / 2
            
            # Update mediabox to target size
            page.mediabox.lower_left = (0, 0)
            page.mediabox.upper_right = (self.target_width, self.target_height)
    
    def merge_email_with_attachments(
        self,
        email_pdf: Path,
        attachment_pdfs: List,  # List of Path or (Path, is_placeholder) tuples
        output_path: Path,
        add_separators: bool = True
    ) -> MergeResult:
        """
        Merge an email PDF with its attachment PDFs.
        
        Args:
            email_pdf: Path to the email body PDF
            attachment_pdfs: List of paths or (path, is_placeholder) tuples
                            If is_placeholder=True, separator is skipped
            output_path: Output path for merged PDF
            add_separators: Whether to add separator pages between attachments
            
        Returns:
            MergeResult with operation details
        """
        errors = []
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Use pikepdf for merging to avoid PyPDF2 page content corruption
        merged_pdf = pikepdf.Pdf.new()
        total_pages = 0
        
        # Add email PDF
        try:
            self._report_progress(0, len(attachment_pdfs) + 1, "Adding email content...")
            
            src_pdf = pikepdf.Pdf.open(str(email_pdf))
            merged_pdf.pages.extend(src_pdf.pages)
            total_pages += len(src_pdf.pages)
        
        except Exception as e:
            errors.append(f"Error reading email PDF: {e}")
            return MergeResult(
                success=False,
                output_path=None,
                page_count=0,
                errors=errors
            )
        
        # Add attachment PDFs
        for i, att_item in enumerate(attachment_pdfs):
            # Handle both old format (just Path) and new format ((Path, is_placeholder))
            if isinstance(att_item, tuple):
                att_pdf, is_placeholder = att_item
            else:
                att_pdf = att_item
                is_placeholder = False
            
            try:
                self._report_progress(
                    i + 1,
                    len(attachment_pdfs) + 1,
                    f"Adding attachment {i + 1}/{len(attachment_pdfs)}..."
                )
                
                # Add separator if requested AND this is not a placeholder PDF
                # (placeholder PDFs already have their own header/info)
                if add_separators and not is_placeholder:
                    separator = self._create_attachment_separator(att_pdf.stem)
                    sep_pdf = pikepdf.Pdf.open(io.BytesIO(separator))
                    merged_pdf.pages.extend(sep_pdf.pages)
                    total_pages += len(sep_pdf.pages)
                
                # Add attachment content using pikepdf (preserves complex XObjects like images)
                src_pdf = pikepdf.Pdf.open(str(att_pdf))
                merged_pdf.pages.extend(src_pdf.pages)
                total_pages += len(src_pdf.pages)
                
                # Copy embedded files from source PDF if any
                self._copy_embedded_files(src_pdf, merged_pdf)
            
            except Exception as e:
                errors.append(f"Error adding attachment {att_pdf.name}: {e}")
                logger.warning(f"Error adding attachment {att_pdf}: {e}")
        
        # Write output
        try:
            merged_pdf.save(str(output_path))
        except Exception as e:
            errors.append(f"Error writing output: {e}")
            return MergeResult(
                success=False,
                output_path=None,
                page_count=0,
                errors=errors
            )
        
        return MergeResult(
            success=len(errors) == 0,
            output_path=output_path,
            page_count=total_pages,
            errors=errors
        )
    
    def merge_chronologically(
        self,
        pdf_files: List[Tuple[Path, str]],  # (path, timestamp)
        output_path: Path,
        add_toc: bool = True,
        add_separators: bool = True
    ) -> MergeResult:
        """
        Merge multiple PDFs in chronological order.
        
        Args:
            pdf_files: List of (pdf_path, timestamp) tuples
            output_path: Output path for merged PDF
            add_toc: Whether to add table of contents
            add_separators: Whether to add separator pages between emails
            
        Returns:
            MergeResult with operation details
        """
        errors = []
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Sort by timestamp (YYYYMMDD_HHMMSS format sorts correctly as string)
        sorted_files = sorted(pdf_files, key=lambda x: x[1])
        
        # Log sorting for debugging
        logger.info(f"Merging {len(sorted_files)} PDFs chronologically")
        if sorted_files:
            logger.info(f"First (oldest): {sorted_files[0][1]} - {sorted_files[0][0].name}")
            logger.info(f"Last (newest): {sorted_files[-1][1]} - {sorted_files[-1][0].name}")
        
        if not sorted_files:
            return MergeResult(
                success=False,
                output_path=None,
                page_count=0,
                errors=["No PDF files to merge"]
            )
        
        # Use pikepdf for merging to avoid PyPDF2 page content corruption issues
        merged_pdf = pikepdf.Pdf.new()
        total_pages = 0
        toc_entries = []  # (title, page_number)
        
        # Track separator PDFs to merge later
        separator_pdfs = []  # List of (insert_position, pdf_bytes)
        
        self._report_progress(0, len(sorted_files), "Starting merge...")
        
        # First pass: collect page counts and TOC entries, merge content
        for i, (pdf_path, timestamp) in enumerate(sorted_files):
            try:
                self._report_progress(
                    i + 1,
                    len(sorted_files),
                    f"Merging {i + 1}/{len(sorted_files)}: {pdf_path.name}"
                )
                
                # Record TOC entry (before any separator)
                toc_entries.append((pdf_path.stem, total_pages + 1))  # +1 for 1-based page numbers
                
                # Add separator if requested (except for first)
                if add_separators and i > 0:
                    separator = self._create_email_separator(pdf_path.stem, timestamp)
                    sep_pdf = pikepdf.Pdf.open(io.BytesIO(separator))
                    merged_pdf.pages.extend(sep_pdf.pages)
                    total_pages += len(sep_pdf.pages)
                    # Update TOC entry to point after separator
                    toc_entries[-1] = (pdf_path.stem, total_pages + 1)
                
                # Add PDF content using pikepdf
                src_pdf = pikepdf.Pdf.open(str(pdf_path))
                merged_pdf.pages.extend(src_pdf.pages)
                total_pages += len(src_pdf.pages)
                
                # Copy embedded files from source PDF if any
                self._copy_embedded_files(src_pdf, merged_pdf)
            
            except Exception as e:
                errors.append(f"Error merging {pdf_path.name}: {e}")
                logger.warning(f"Error merging {pdf_path}: {e}")
        
        # Create final output with optional TOC
        if add_toc and toc_entries:
            # Create TOC - ADJUST page numbers to account for TOC pages at the front
            # First, create a dummy TOC to know how many pages it will be
            dummy_toc = self._create_table_of_contents(toc_entries)
            dummy_reader = PdfReader(io.BytesIO(dummy_toc))
            toc_pages = len(dummy_reader.pages)
            
            # Adjust all TOC entries to account for TOC pages being inserted at front
            adjusted_toc_entries = [(title, page_num + toc_pages) for title, page_num in toc_entries]
            
            # Now create the real TOC with correct page numbers
            toc_pdf_bytes = self._create_table_of_contents(adjusted_toc_entries)
            toc_pdf = pikepdf.Pdf.open(io.BytesIO(toc_pdf_bytes))
            
            # Create new PDF with TOC first, then content
            final_pdf = pikepdf.Pdf.new()
            final_pdf.pages.extend(toc_pdf.pages)
            final_pdf.pages.extend(merged_pdf.pages)
            
            # Copy embedded files from merged_pdf to final_pdf
            self._copy_embedded_files(merged_pdf, final_pdf)
            
            # Add clickable links to TOC entries
            self._add_toc_links(final_pdf, toc_pages, adjusted_toc_entries)
            
            merged_pdf = final_pdf
            total_pages += toc_pages
        
        # Write output with linearization for better compatibility
        try:
            # Linearize makes PDF "web-optimized" and more compatible with various readers
            merged_pdf.save(str(output_path), linearize=True)
        except Exception as e:
            errors.append(f"Error writing output: {e}")
            return MergeResult(
                success=False,
                output_path=None,
                page_count=0,
                errors=errors
            )
        
        self._report_progress(len(sorted_files), len(sorted_files), "Merge complete!")
        
        return MergeResult(
            success=len(errors) == 0,
            output_path=output_path,
            page_count=total_pages,
            errors=errors
        )
    
    def _create_attachment_separator(self, attachment_name: str) -> bytes:
        """Create a separator page for an attachment."""
        buffer = io.BytesIO()
        
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=inch,
            leftMargin=inch,
            topMargin=2*inch,
            bottomMargin=inch
        )
        
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle(
            name='AttSeparator',
            parent=styles['Title'],
            fontSize=16,
            textColor=colors.Color(0.3, 0.3, 0.6),
            spaceAfter=20
        )
        
        subtitle_style = ParagraphStyle(
            name='AttSubtitle',
            parent=styles['Normal'],
            fontSize=12,
            textColor=colors.Color(0.5, 0.5, 0.5),
            alignment=1  # Center
        )
        
        story = [
            Spacer(1, 2*inch),
            Paragraph("ATTACHMENT", title_style),
            Spacer(1, 20),
            Paragraph(self._escape_text(attachment_name), subtitle_style),
        ]
        
        doc.build(story)
        return buffer.getvalue()
    
    def _create_email_separator(self, email_name: str, timestamp: str) -> bytes:
        """Create a separator page between emails."""
        buffer = io.BytesIO()
        
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=inch,
            leftMargin=inch,
            topMargin=2*inch,
            bottomMargin=inch
        )
        
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            name='EmailSeparator',
            parent=styles['Title'],
            fontSize=14,
            textColor=colors.Color(0.2, 0.2, 0.4),
            spaceAfter=20,
            alignment=1
        )
        
        date_style = ParagraphStyle(
            name='DateStyle',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.Color(0.4, 0.4, 0.4),
            alignment=1
        )
        
        # Format timestamp for display
        display_date = timestamp
        if len(timestamp) >= 15:  # YYYYMMDD_HHMMSS format
            try:
                display_date = f"{timestamp[0:4]}-{timestamp[4:6]}-{timestamp[6:8]} {timestamp[9:11]}:{timestamp[11:13]}:{timestamp[13:15]}"
            except:
                pass
        
        story = [
            Spacer(1, 2*inch),
            Paragraph(self._escape_text(email_name), title_style),
            Spacer(1, 10),
            Paragraph(display_date, date_style),
        ]
        
        doc.build(story)
        return buffer.getvalue()
    
    def _create_table_of_contents(self, entries: List[Tuple[str, int]]) -> bytes:
        """Create a table of contents PDF with page numbers (no clickable links).
        
        Note: Internal PDF links are complex to implement correctly across merged PDFs.
        The TOC shows page numbers for manual navigation instead.
        """
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        
        buffer = io.BytesIO()
        
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=inch,
            bottomMargin=inch
        )
        
        # Register a standard font that will be embedded
        # Use Helvetica which is a PDF standard font (always available)
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            name='TOCTitle',
            parent=styles['Title'],
            fontName='Helvetica-Bold',
            fontSize=18,
            spaceAfter=30
        )
        
        entry_style = ParagraphStyle(
            name='TOCEntry',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=18,
        )
        
        page_style = ParagraphStyle(
            name='TOCPage',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10,
            leading=18,
            alignment=2  # Right align
        )
        
        # Calculate available width for TOC entries
        page_width = letter[0] - 1.5*inch  # Account for margins
        
        story = [
            Paragraph("Table of Contents", title_style),
            Spacer(1, 20),
        ]
        
        # Create TOC entries using a Table for proper alignment
        toc_data = []
        for title, page_num in entries:
            # Truncate long titles
            display_title = title[:60] + "..." if len(title) > 60 else title
            display_title = self._escape_text(display_title)
            
            # Page numbers are already adjusted by caller
            toc_data.append([display_title, page_num])
        
        if toc_data:
            # Create table with dot leaders
            col_widths = [page_width - 0.5*inch, 0.5*inch]
            
            # Build table rows with styling - plain text, no links
            table_data = []
            for title, page_num in toc_data:
                title_para = Paragraph(title, entry_style)
                page_para = Paragraph(str(page_num), page_style)
                table_data.append([title_para, page_para])
            
            toc_table = Table(table_data, colWidths=col_widths)
            toc_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ]))
            
            story.append(toc_table)
        
        story.append(PageBreak())
        
        doc.build(story)
        return buffer.getvalue()
    
    def _escape_text(self, text: str) -> str:
        """Escape text for reportlab."""
        if not text:
            return ""
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')
        return text
    
    def simple_merge(self, pdf_files: List[Path], output_path: Path) -> MergeResult:
        """
        Simple merge of PDFs without any extras.
        
        Args:
            pdf_files: List of PDF paths to merge
            output_path: Output path
            
        Returns:
            MergeResult
        """
        errors = []
        
        try:
            merger = PdfMerger()
            
            for i, pdf_path in enumerate(pdf_files):
                self._report_progress(i + 1, len(pdf_files), f"Merging {pdf_path.name}")
                try:
                    merger.append(str(pdf_path))
                except Exception as e:
                    errors.append(f"Error with {pdf_path.name}: {e}")
            
            output_path.parent.mkdir(parents=True, exist_ok=True)
            merger.write(str(output_path))
            merger.close()
            
            # Count pages
            reader = PdfReader(str(output_path))
            page_count = len(reader.pages)
            
            return MergeResult(
                success=len(errors) == 0,
                output_path=output_path,
                page_count=page_count,
                errors=errors
            )
        
        except Exception as e:
            errors.append(f"Merge failed: {e}")
            return MergeResult(
                success=False,
                output_path=None,
                page_count=0,
                errors=errors
            )
