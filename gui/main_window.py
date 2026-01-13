"""
Main Window Module

Main application window for the Mail Converter application.
"""

import os
import sys
import threading
import logging
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional

from core.conversion_pipeline import (
    ConversionPipeline, 
    PipelineConfig, 
    PipelineResult,
    PipelineProgress,
    PipelineStage
)
from core.duplicate_detector import DuplicateCertainty
from .progress_dialog import ProgressDialog
from .settings_dialog import SettingsDialog

logger = logging.getLogger(__name__)


class MainWindow:
    """Main application window."""
    
    def __init__(self, root: tk.Tk):
        """
        Initialize the main window.
        
        Args:
            root: Tkinter root window
        """
        self.root = root
        self.root.title("Mayo's Mail Converter - PST to PDF")
        self.root.geometry("700x550")
        self.root.minsize(600, 450)
        
        # State
        self.pst_path: Optional[str] = None
        self.output_dir: Optional[str] = None
        self.pipeline: Optional[ConversionPipeline] = None
        self.conversion_thread: Optional[threading.Thread] = None
        
        # Settings (defaults)
        self.settings = {
            'ocr_enabled': True,
            'detect_duplicates': True,
            'duplicate_certainty': 'HIGH',
            'keep_individual_pdfs': True,
            'create_combined_pdf': True,
            'add_toc': True,
            'add_separators': False,
            'add_att_separators': False,
            'page_margin': 0.5,
            'skip_deleted_items': True,
            'date_from': None,
            'date_to': None
        }
        
        # Setup UI
        self._setup_styles()
        self._create_widgets()
        self._bind_events()
        
        # Center window
        self._center_window()
    
    def _setup_styles(self):
        """Setup ttk styles."""
        style = ttk.Style()
        
        # Try to use a modern theme
        available_themes = style.theme_names()
        if 'aqua' in available_themes:  # macOS
            style.theme_use('aqua')
        elif 'vista' in available_themes:  # Windows
            style.theme_use('vista')
        elif 'clam' in available_themes:
            style.theme_use('clam')
        
        # Custom styles
        style.configure('Title.TLabel', font=('Helvetica', 16, 'bold'))
        style.configure('Subtitle.TLabel', font=('Helvetica', 11))
        style.configure('Status.TLabel', font=('Helvetica', 10))
        style.configure('Big.TButton', font=('Helvetica', 12), padding=10)
    
    def _create_widgets(self):
        """Create all UI widgets."""
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Title
        title_label = ttk.Label(
            self.main_frame, 
            text="Mayo's Mail Converter",
            style='Title.TLabel'
        )
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 5))
        row += 1
        
        subtitle_label = ttk.Label(
            self.main_frame,
            text="Convert PST emails to PDF with attachments",
            style='Subtitle.TLabel'
        )
        subtitle_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # Separator
        ttk.Separator(self.main_frame).grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        
        # PST File Selection
        ttk.Label(self.main_frame, text="PST File:").grid(
            row=row, column=0, sticky="w", pady=5
        )
        
        self.pst_entry = ttk.Entry(self.main_frame, width=50)
        self.pst_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        
        self.browse_pst_btn = ttk.Button(
            self.main_frame, 
            text="Browse...",
            command=self._browse_pst
        )
        self.browse_pst_btn.grid(row=row, column=2, pady=5)
        row += 1
        
        # Output Directory Selection
        ttk.Label(self.main_frame, text="Output Folder:").grid(
            row=row, column=0, sticky="w", pady=5
        )
        
        self.output_entry = ttk.Entry(self.main_frame, width=50)
        self.output_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        
        self.browse_output_btn = ttk.Button(
            self.main_frame,
            text="Browse...",
            command=self._browse_output
        )
        self.browse_output_btn.grid(row=row, column=2, pady=5)
        row += 1
        
        # Separator
        ttk.Separator(self.main_frame).grid(row=row, column=0, columnspan=3, sticky="ew", pady=15)
        row += 1
        
        # Options Frame
        options_frame = ttk.LabelFrame(self.main_frame, text="Options", padding=10)
        options_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        options_frame.columnconfigure(1, weight=1)
        row += 1
        
        # OCR Option
        self.ocr_var = tk.BooleanVar(value=True)
        ocr_check = ttk.Checkbutton(
            options_frame,
            text="Enable OCR for images and scanned PDFs",
            variable=self.ocr_var
        )
        ocr_check.grid(row=0, column=0, sticky="w", pady=2)
        
        # Duplicate Detection
        self.dup_var = tk.BooleanVar(value=True)
        dup_check = ttk.Checkbutton(
            options_frame,
            text="Detect and skip duplicate emails",
            variable=self.dup_var
        )
        dup_check.grid(row=1, column=0, sticky="w", pady=2)
        
        # Duplicate Certainty
        dup_cert_frame = ttk.Frame(options_frame)
        dup_cert_frame.grid(row=1, column=1, sticky="w", padx=20)
        
        ttk.Label(dup_cert_frame, text="Certainty:").pack(side=tk.LEFT)
        
        self.dup_certainty_var = tk.StringVar(value="HIGH")
        dup_combo = ttk.Combobox(
            dup_cert_frame,
            textvariable=self.dup_certainty_var,
            values=["LOW", "MEDIUM", "HIGH", "EXACT"],
            state="readonly",
            width=10
        )
        dup_combo.pack(side=tk.LEFT, padx=5)
        
        # Combined PDF Option
        self.combined_var = tk.BooleanVar(value=True)
        combined_check = ttk.Checkbutton(
            options_frame,
            text="Create combined chronological PDF",
            variable=self.combined_var
        )
        combined_check.grid(row=2, column=0, sticky="w", pady=2)
        
        # Keep Individual PDFs
        self.individual_var = tk.BooleanVar(value=True)
        individual_check = ttk.Checkbutton(
            options_frame,
            text="Keep individual email PDFs",
            variable=self.individual_var
        )
        individual_check.grid(row=3, column=0, sticky="w", pady=2)
        
        # Merge Folders Option
        self.merge_folders_var = tk.BooleanVar(value=False)
        merge_folders_check = ttk.Checkbutton(
            options_frame,
            text="Merge all folders into one PDF (uncheck for separate PDF per folder)",
            variable=self.merge_folders_var
        )
        merge_folders_check.grid(row=4, column=0, columnspan=2, sticky="w", pady=2)
        
        # Advanced Settings Button
        self.settings_btn = ttk.Button(
            options_frame,
            text="Advanced Settings...",
            command=self._show_settings
        )
        self.settings_btn.grid(row=5, column=1, sticky="e", pady=2)
        
        # Status Frame
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        status_frame.columnconfigure(0, weight=1)
        row += 1
        
        self.status_label = ttk.Label(
            status_frame,
            text="Ready. Select a PST file to begin.",
            style='Status.TLabel'
        )
        self.status_label.grid(row=0, column=0, sticky="w")
        
        # Button Frame
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=20)
        row += 1
        
        self.convert_btn = ttk.Button(
            button_frame,
            text="Convert to PDF",
            style='Big.TButton',
            command=self._start_conversion
        )
        self.convert_btn.pack(side=tk.LEFT, padx=10)
        
        self.cancel_btn = ttk.Button(
            button_frame,
            text="Cancel",
            command=self._cancel_conversion,
            state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=10)
        
        # Footer
        footer_frame = ttk.Frame(self.main_frame)
        footer_frame.grid(row=row, column=0, columnspan=3, sticky="ew")
        
        version_label = ttk.Label(
            footer_frame,
            text="v1.0.0",
            foreground="gray"
        )
        version_label.pack(side=tk.RIGHT)
    
    def _bind_events(self):
        """Bind event handlers."""
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # Drag and drop (basic support)
        self.pst_entry.bind('<Button-1>', lambda e: self._browse_pst() if not self.pst_entry.get() else None)
    
    def _center_window(self):
        """Center the window on screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def _browse_pst(self):
        """Open file dialog to select PST file."""
        filepath = filedialog.askopenfilename(
            title="Select PST File",
            filetypes=[
                ("PST Files", "*.pst"),
                ("All Files", "*.*")
            ]
        )
        
        if filepath:
            self.pst_path = filepath
            self.pst_entry.delete(0, tk.END)
            self.pst_entry.insert(0, filepath)
            
            # Auto-suggest output directory
            if not self.output_entry.get():
                parent_dir = Path(filepath).parent
                output_name = Path(filepath).stem + "_converted"
                suggested_output = parent_dir / output_name
                self.output_entry.insert(0, str(suggested_output))
            
            self._update_status(f"Selected: {Path(filepath).name}")
    
    def _browse_output(self):
        """Open dialog to select output directory."""
        dirpath = filedialog.askdirectory(
            title="Select Output Folder"
        )
        
        if dirpath:
            self.output_dir = dirpath
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, dirpath)
    
    def _show_settings(self):
        """Show advanced settings dialog."""
        dialog = SettingsDialog(self.root, self.settings)
        if dialog.result:
            self.settings.update(dialog.result)
    
    def _update_status(self, message: str):
        """Update status label."""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def _validate_inputs(self) -> bool:
        """Validate inputs before starting conversion."""
        pst_path = self.pst_entry.get().strip()
        output_dir = self.output_entry.get().strip()
        
        if not pst_path:
            messagebox.showerror("Error", "Please select a PST file.")
            return False
        
        if not os.path.isfile(pst_path):
            messagebox.showerror("Error", f"PST file not found:\n{pst_path}")
            return False
        
        if not output_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return False
        
        # Check if output directory exists or can be created
        try:
            Path(output_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot create output folder:\n{e}")
            return False
        
        return True
    
    def _start_conversion(self):
        """Start the conversion process."""
        if not self._validate_inputs():
            return
        
        # Get values from UI
        pst_path = self.pst_entry.get().strip()
        output_dir = self.output_entry.get().strip()
        
        # Create config
        certainty_map = {
            'LOW': DuplicateCertainty.LOW,
            'MEDIUM': DuplicateCertainty.MEDIUM,
            'HIGH': DuplicateCertainty.HIGH,
            'EXACT': DuplicateCertainty.EXACT
        }
        
        config = PipelineConfig(
            pst_path=pst_path,
            output_dir=output_dir,
            ocr_enabled=self.ocr_var.get(),
            detect_duplicates=self.dup_var.get(),
            duplicate_certainty=certainty_map.get(
                self.dup_certainty_var.get(), 
                DuplicateCertainty.HIGH
            ),
            keep_individual_pdfs=self.individual_var.get(),
            create_combined_pdf=self.combined_var.get(),
            add_toc=self.settings.get('add_toc', True),
            add_separators=self.settings.get('add_separators', True),
            page_size=self.settings.get('page_size', 'Letter'),
            page_margin=self.settings.get('page_margin', 0.5),
            merge_folders=self.merge_folders_var.get(),
            rename_emls=self.settings.get('rename_emls', True),
            skip_deleted_items=self.settings.get('skip_deleted_items', True),
            date_from=self.settings.get('date_from'),
            date_to=self.settings.get('date_to')
        )
        
        # Disable UI
        self._set_ui_state(False)
        
        # Create and show progress dialog
        self.progress_dialog = ProgressDialog(self.root)
        
        # Create pipeline with progress callback
        self.pipeline = ConversionPipeline(
            config,
            progress_callback=self._on_progress
        )
        
        # Start conversion in thread
        self.conversion_thread = threading.Thread(
            target=self._run_conversion,
            daemon=True
        )
        self.conversion_thread.start()
        
        # Check thread periodically
        self._check_thread()
    
    def _run_conversion(self):
        """Run conversion in background thread."""
        try:
            self.result = self.pipeline.run()
        except Exception as e:
            logger.exception("Conversion failed")
            self.result = PipelineResult(
                success=False,
                stage_reached=PipelineStage.FAILED,
                errors=[str(e)]
            )
    
    def _on_progress(self, progress: PipelineProgress):
        """Handle progress updates from pipeline."""
        # Schedule UI update on main thread
        self.root.after(0, lambda: self._update_progress(progress))
    
    def _update_progress(self, progress: PipelineProgress):
        """Update progress dialog on main thread."""
        if hasattr(self, 'progress_dialog') and self.progress_dialog:
            self.progress_dialog.update_progress(
                progress.percentage,
                progress.message,
                progress.stage.value
            )
    
    def _check_thread(self):
        """Check if conversion thread is still running."""
        if self.conversion_thread and self.conversion_thread.is_alive():
            # Check again after 100ms
            self.root.after(100, self._check_thread)
        else:
            # Thread finished
            self._conversion_complete()
    
    def _conversion_complete(self):
        """Handle conversion completion."""
        # Close progress dialog
        if hasattr(self, 'progress_dialog') and self.progress_dialog:
            self.progress_dialog.close()
        
        # Re-enable UI
        self._set_ui_state(True)
        
        # Show results
        if hasattr(self, 'result'):
            result = self.result
            
            if result.success:
                message = (
                    f"Conversion complete!\n\n"
                    f"Emails processed: {result.emails_processed}\n"
                    f"Duplicates skipped: {result.duplicates_skipped}\n"
                    f"Attachments converted: {result.attachments_converted}\n"
                    f"Duration: {result.duration_seconds:.1f} seconds"
                )
                
                if result.combined_pdf_path:
                    message += f"\n\nCombined PDF:\n{result.combined_pdf_path}"
                
                messagebox.showinfo("Success", message)
                
                # Offer to open output folder
                if messagebox.askyesno("Open Folder", "Open output folder?"):
                    self._open_folder(str(result.individual_pdfs_dir.parent))
            
            else:
                error_msg = "Conversion failed.\n\n"
                
                if result.errors:
                    error_msg += "Errors:\n"
                    for error in result.errors[:5]:  # Show first 5 errors
                        error_msg += f"• {error}\n"
                
                if result.warnings:
                    error_msg += "\nWarnings:\n"
                    for warning in result.warnings[:5]:
                        error_msg += f"• {warning}\n"
                
                messagebox.showerror("Conversion Failed", error_msg)
        
        self._update_status("Ready")
    
    def _cancel_conversion(self):
        """Cancel ongoing conversion."""
        if self.pipeline:
            self.pipeline.cancel()
            self._update_status("Cancelling...")
    
    def _set_ui_state(self, enabled: bool):
        """Enable or disable UI elements during conversion."""
        state = tk.NORMAL if enabled else tk.DISABLED
        
        self.browse_pst_btn.config(state=state)
        self.browse_output_btn.config(state=state)
        self.pst_entry.config(state=state)
        self.output_entry.config(state=state)
        self.convert_btn.config(state=state)
        self.settings_btn.config(state=state)
        
        # Cancel button is opposite
        self.cancel_btn.config(state=tk.DISABLED if enabled else tk.NORMAL)
    
    def _open_folder(self, path: str):
        """Open folder in file explorer."""
        import subprocess
        import platform
        
        system = platform.system()
        
        try:
            if system == "Darwin":  # macOS
                subprocess.run(["open", path])
            elif system == "Windows":
                subprocess.run(["explorer", path])
            else:  # Linux
                subprocess.run(["xdg-open", path])
        except Exception as e:
            logger.warning(f"Could not open folder: {e}")
    
    def _on_close(self):
        """Handle window close."""
        if self.conversion_thread and self.conversion_thread.is_alive():
            if messagebox.askyesno(
                "Conversion in Progress",
                "A conversion is in progress. Cancel and exit?"
            ):
                if self.pipeline:
                    self.pipeline.cancel()
                self.root.destroy()
        else:
            self.root.destroy()
