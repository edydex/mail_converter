"""
Main Window Module

Main application window for the Mail Converter application.
Features two main tabs:
- PDF Converter: Convert emails to PDF
- Email Tools: Mailbox manipulation (compare, merge, dedupe, filter)
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
    PipelineStage,
    InputType
)
from core.duplicate_detector import DuplicateCertainty
from core.eml_parser import EMLParser
from core.email_to_pdf import EmailToPDFConverter
from core.attachment_converter import AttachmentConverter
from core.pdf_merger import PDFMerger
from .progress_dialog import ProgressDialog
from .settings_dialog import SettingsDialog
from .email_tools_tab import EmailToolsTab

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
        self.root.title("Mayo's Mail Converter - Email to PDF")
        self.root.geometry("750x600")
        self.root.minsize(650, 500)
        
        # State
        self.input_paths: list = []  # Support multiple inputs
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
            'load_remote_images': False,
            'date_from': None,
            'date_to': None
        }
        
        # Backwards compatibility
        self.pst_path: Optional[str] = None
        
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
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # Title
        title_frame = ttk.Frame(self.main_frame)
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        # Create menu bar
        self._create_menu_bar()
        
        title_label = ttk.Label(
            title_frame, 
            text="Mayo's Mail Converter",
            style='Title.TLabel'
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            title_frame,
            text="Email to PDF conversion & mailbox tools",
            style='Subtitle.TLabel'
        )
        subtitle_label.pack()
        
        # Create main notebook (tabs)
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=1, column=0, sticky="nsew")
        
        # Create PDF Converter tab
        self.pdf_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.pdf_tab, text="PDF Converter")
        self._create_pdf_converter_tab()
        
        # Create Email Tools tab
        self.tools_tab = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(self.tools_tab, text="Email Tools")
        self.email_tools = EmailToolsTab(self.tools_tab)
        
        # Footer
        footer_frame = ttk.Frame(self.main_frame)
        footer_frame.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        
        version_label = ttk.Label(
            footer_frame,
            text="v1.3.0",
            foreground="gray"
        )
        version_label.pack(side=tk.RIGHT)
    
    def _create_pdf_converter_tab(self):
        """Create the PDF Converter tab content."""
        frame = self.pdf_tab
        frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Input Selection Label
        ttk.Label(frame, text="Input File(s)/Folder:").grid(
            row=row, column=0, sticky="w", pady=5
        )
        
        self.pst_entry = ttk.Entry(frame, width=45)
        self.pst_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        
        # Button frame for browse options
        browse_frame = ttk.Frame(frame)
        browse_frame.grid(row=row, column=2, columnspan=2, pady=5, sticky="e")
        
        self.browse_pst_btn = ttk.Button(
            browse_frame, 
            text="Files...",
            command=self._browse_files,
            width=8
        )
        self.browse_pst_btn.pack(side=tk.LEFT, padx=2)
        
        self.browse_folder_btn = ttk.Button(
            browse_frame, 
            text="Folder...",
            command=self._browse_folder,
            width=8
        )
        self.browse_folder_btn.pack(side=tk.LEFT, padx=2)
        row += 1
        
        # Selected files indicator
        self.files_count_label = ttk.Label(
            frame,
            text="",
            foreground="gray"
        )
        self.files_count_label.grid(row=row, column=1, sticky="w", padx=5)
        row += 1
        
        # Output Directory Selection
        ttk.Label(frame, text="Output Folder:").grid(
            row=row, column=0, sticky="w", pady=5
        )
        
        self.output_entry = ttk.Entry(frame, width=45)
        self.output_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        
        self.browse_output_btn = ttk.Button(
            frame,
            text="Browse...",
            command=self._browse_output,
            width=8
        )
        self.browse_output_btn.grid(row=row, column=2, pady=5, sticky="w", padx=2)
        row += 1
        
        # Separator
        ttk.Separator(frame).grid(row=row, column=0, columnspan=4, sticky="ew", pady=10)
        row += 1
        
        # Options Frame
        options_frame = ttk.LabelFrame(frame, text="Options", padding=10)
        options_frame.grid(row=row, column=0, columnspan=4, sticky="ew", pady=10)
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
        
        # Combine Folders By Name Option (for multiple PST/MBOX files)
        self.combine_folders_var = tk.BooleanVar(value=False)
        self.combine_folders_check = ttk.Checkbutton(
            options_frame,
            text="Combine folders by name across multiple sources",
            variable=self.combine_folders_var
        )
        self.combine_folders_check.grid(row=5, column=0, columnspan=2, sticky="w", pady=2)
        
        # Add tooltip-like help text
        combine_help = ttk.Label(
            options_frame,
            text="(e.g., 'Inbox' from user1 + 'Inbox' from user2 → single 'Inbox' folder)",
            foreground="gray",
            font=('Helvetica', 9)
        )
        combine_help.grid(row=6, column=0, columnspan=2, sticky="w", padx=20, pady=(0, 5))
        
        # Advanced Settings Button
        self.settings_btn = ttk.Button(
            options_frame,
            text="Advanced Settings...",
            command=self._show_settings
        )
        self.settings_btn.grid(row=7, column=1, sticky="e", pady=2)
        
        # Status Frame
        status_frame = ttk.Frame(frame)
        status_frame.grid(row=row, column=0, columnspan=4, sticky="ew", pady=10)
        status_frame.columnconfigure(0, weight=1)
        row += 1
        
        self.status_label = ttk.Label(
            status_frame,
            text="Ready. Select file(s) or a folder to begin.",
            style='Status.TLabel'
        )
        self.status_label.grid(row=0, column=0, sticky="w")
        
        # Button Frame
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, columnspan=4, pady=15)
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
    
    def _create_menu_bar(self):
        """Create the application menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        help_menu.add_command(label="System Diagnostics...", command=self._show_diagnostics)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self._show_about)
    
    def _show_diagnostics(self):
        """Show system diagnostics dialog."""
        try:
            from utils.system_info import generate_diagnostic_report
            report = generate_diagnostic_report()
        except Exception as e:
            report = f"Error generating diagnostics: {e}"
        
        # Create dialog window
        diag_window = tk.Toplevel(self.root)
        diag_window.title("System Diagnostics")
        diag_window.geometry("700x500")
        diag_window.transient(self.root)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(diag_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget = tk.Text(
            text_frame, 
            wrap=tk.NONE, 
            font=('Courier', 10),
            yscrollcommand=scrollbar.set
        )
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)
        
        # Insert report
        text_widget.insert('1.0', report)
        text_widget.config(state=tk.DISABLED)
        
        # Button frame
        btn_frame = ttk.Frame(diag_window)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        def copy_to_clipboard():
            self.root.clipboard_clear()
            self.root.clipboard_append(report)
            messagebox.showinfo("Copied", "Diagnostic report copied to clipboard!")
        
        def save_to_file():
            filepath = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfilename="diagnostics_report.txt"
            )
            if filepath:
                try:
                    with open(filepath, 'w') as f:
                        f.write(report)
                    messagebox.showinfo("Saved", f"Report saved to:\n{filepath}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not save file: {e}")
        
        ttk.Button(btn_frame, text="Copy to Clipboard", command=copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save to File", command=save_to_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", command=diag_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def _show_about(self):
        """Show about dialog."""
        messagebox.showinfo(
            "About Mayo's Mail Converter",
            "Mayo's Mail Converter v1.3.0\n\n"
            "Convert PST, MBOX, MSG, and EML files to searchable PDFs.\n\n"
            "https://github.com/yourusername/mail_converter"
        )
    
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
    
    def _browse_files(self):
        """Open file dialog to select email files (PST, MBOX, MSG, EML)."""
        filepaths = filedialog.askopenfilenames(
            title="Select Email File(s)",
            filetypes=[
                ("All Email Files", "*.pst *.mbox *.msg *.eml"),
                ("PST Files", "*.pst"),
                ("MBOX Files", "*.mbox"),
                ("MSG Files", "*.msg"),
                ("EML Files", "*.eml"),
                ("All Files", "*.*")
            ]
        )
        
        if filepaths:
            self.input_paths = list(filepaths)
            # For backwards compatibility
            self.pst_path = filepaths[0] if filepaths else None
            
            # Display in entry
            if len(filepaths) == 1:
                display_text = filepaths[0]
            else:
                display_text = f"{len(filepaths)} files selected"
            
            self.pst_entry.delete(0, tk.END)
            self.pst_entry.insert(0, display_text)
            
            # Update files count label
            self._update_files_count()
            
            # Auto-suggest output directory
            if not self.output_entry.get():
                parent_dir = Path(filepaths[0]).parent
                if len(filepaths) == 1:
                    output_name = Path(filepaths[0]).stem + "_converted"
                else:
                    output_name = "emails_converted"
                suggested_output = parent_dir / output_name
                self.output_entry.insert(0, str(suggested_output))
            
            self._update_status(f"Selected {len(filepaths)} file(s)")
    
    def _browse_folder(self):
        """Open dialog to select a folder containing email files."""
        dirpath = filedialog.askdirectory(
            title="Select Folder with Email Files"
        )
        
        if dirpath:
            self.input_paths = [dirpath]
            # For backwards compatibility
            self.pst_path = dirpath
            
            self.pst_entry.delete(0, tk.END)
            self.pst_entry.insert(0, dirpath)
            
            # Count files in folder
            self._update_files_count()
            
            # Auto-suggest output directory
            if not self.output_entry.get():
                parent_dir = Path(dirpath).parent
                output_name = Path(dirpath).name + "_converted"
                suggested_output = parent_dir / output_name
                self.output_entry.insert(0, str(suggested_output))
            
            self._update_status(f"Selected folder: {Path(dirpath).name}")
    
    def _update_files_count(self):
        """Update the file count label based on selected inputs."""
        if not self.input_paths:
            self.files_count_label.config(text="")
            return
        
        if len(self.input_paths) == 1 and Path(self.input_paths[0]).is_dir():
            # Count email files in folder
            folder = Path(self.input_paths[0])
            pst_count = len(list(folder.glob("*.pst")) + list(folder.glob("*.PST")))
            mbox_count = len(list(folder.glob("*.mbox")) + list(folder.glob("*.MBOX")))
            msg_count = len(list(folder.glob("**/*.msg")) + list(folder.glob("**/*.MSG")))
            eml_count = len(list(folder.glob("**/*.eml")) + list(folder.glob("**/*.EML")))
            
            parts = []
            if pst_count: parts.append(f"{pst_count} PST")
            if mbox_count: parts.append(f"{mbox_count} MBOX")
            if msg_count: parts.append(f"{msg_count} MSG")
            if eml_count: parts.append(f"{eml_count} EML")
            
            if parts:
                self.files_count_label.config(text=f"Found: {', '.join(parts)}")
            else:
                self.files_count_label.config(text="No email files found in folder")
        elif len(self.input_paths) > 1:
            # List file types
            exts = {}
            for p in self.input_paths:
                ext = Path(p).suffix.lower().lstrip('.')
                exts[ext] = exts.get(ext, 0) + 1
            
            parts = [f"{count} {ext.upper()}" for ext, count in exts.items()]
            self.files_count_label.config(text=f"Selected: {', '.join(parts)}")
        else:
            self.files_count_label.config(text="")
    
    def _browse_pst(self):
        """Legacy method - redirects to _browse_files."""
        self._browse_files()
    
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
        input_text = self.pst_entry.get().strip()
        output_dir = self.output_entry.get().strip()
        
        if not input_text:
            messagebox.showerror("Error", "Please select email file(s) or a folder.")
            return False
        
        # Check if we have input_paths set
        if not self.input_paths:
            # Try to use the entry text as a single path
            if os.path.exists(input_text):
                self.input_paths = [input_text]
            else:
                messagebox.showerror("Error", f"Path not found:\n{input_text}")
                return False
        
        # Validate all input paths exist
        for path in self.input_paths:
            if not os.path.exists(path):
                messagebox.showerror("Error", f"Path not found:\n{path}")
                return False
            
            # Validate file extension for files (not folders)
            if os.path.isfile(path):
                ext = Path(path).suffix.lower()
                if ext not in ['.pst', '.mbox', '.msg', '.eml']:
                    messagebox.showerror(
                        "Error", 
                        f"Unsupported file type: {ext}\n"
                        "Supported types: .pst, .mbox, .msg, .eml"
                    )
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
        output_dir = self.output_entry.get().strip()
        
        # Check if this is a single EML file - handle with simple converter
        if (len(self.input_paths) == 1 and 
            os.path.isfile(self.input_paths[0]) and
            Path(self.input_paths[0]).suffix.lower() == '.eml'):
            self._convert_eml_file(self.input_paths[0], output_dir)
            return
        
        # Create config for full pipeline conversion
        certainty_map = {
            'LOW': DuplicateCertainty.LOW,
            'MEDIUM': DuplicateCertainty.MEDIUM,
            'HIGH': DuplicateCertainty.HIGH,
            'EXACT': DuplicateCertainty.EXACT
        }
        
        config = PipelineConfig(
            pst_path=self.input_paths[0] if self.input_paths else "",
            input_paths=self.input_paths,
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
            add_separators=self.settings.get('add_separators', False),
            add_att_separators=self.settings.get('add_att_separators', False),
            page_size=self.settings.get('page_size', 'Letter'),
            page_margin=self.settings.get('page_margin', 0.5),
            load_remote_images=self.settings.get('load_remote_images', False),
            merge_folders=self.merge_folders_var.get(),
            combine_folders_by_name=self.combine_folders_var.get(),
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
    
    def _convert_eml_file(self, eml_path: str, output_dir: str):
        """
        Convert a single EML file to PDF.
        
        This is a simplified conversion path for troubleshooting individual emails.
        
        Args:
            eml_path: Path to the EML file
            output_dir: Output directory for the PDF
        """
        try:
            self._update_status("Converting EML file...")
            self._set_ui_state(False)
            
            # Parse the EML file
            parser = EMLParser()
            email_data = parser.parse_file(eml_path)
            
            # Create output path
            eml_name = Path(eml_path).stem
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)
            
            # Convert email to PDF
            converter = EmailToPDFConverter(
                load_remote_images=self.settings.get('load_remote_images', False)
            )
            email_pdf_path = output_path / f"{eml_name}.pdf"
            converter.convert_email_to_pdf(email_data, email_pdf_path, include_headers=True)
            
            # Convert attachments if any
            attachment_pdfs = []
            if email_data.attachments:
                self._update_status(f"Converting {len(email_data.attachments)} attachment(s)...")
                att_output_dir = output_path / "attachments"
                att_converter = AttachmentConverter(att_output_dir)
                
                for att in email_data.attachments:
                    try:
                        result = att_converter.convert_bytes(
                            content=att.content,
                            content_type=att.content_type,
                            filename=att.filename,
                            output_dir=att_output_dir
                        )
                        if result.output_path:
                            attachment_pdfs.append((att.filename, result.output_path))
                    except Exception as e:
                        logger.warning(f"Failed to convert attachment {att.filename}: {e}")
            
            # Merge email with attachments if we have any
            if attachment_pdfs:
                self._update_status("Merging email with attachments...")
                merger = PDFMerger()
                
                # Create combined PDF using merge_email_with_attachments
                combined_path = output_path / f"{eml_name}_with_attachments.pdf"
                
                # Extract just the paths for merging
                att_pdf_paths = [att_pdf for _, att_pdf in attachment_pdfs]
                
                result = merger.merge_email_with_attachments(
                    email_pdf=email_pdf_path,
                    attachment_pdfs=att_pdf_paths,
                    output_path=combined_path,
                    add_separators=self.settings.get('add_att_separators', False)
                )
                
                if result.success:
                    final_pdf = combined_path
                else:
                    logger.warning(f"Merge failed: {result.errors}")
                    final_pdf = email_pdf_path
            else:
                final_pdf = email_pdf_path
            
            self._set_ui_state(True)
            self._update_status("Conversion complete!")
            
            # Show success message
            messagebox.showinfo(
                "Success",
                f"EML converted successfully!\n\n"
                f"Output: {final_pdf}\n"
                f"Attachments: {len(attachment_pdfs)}"
            )
            
            # Offer to open the PDF
            if messagebox.askyesno("Open PDF", "Would you like to open the PDF?"):
                import subprocess
                if sys.platform == 'darwin':
                    subprocess.run(['open', str(final_pdf)])
                elif sys.platform == 'win32':
                    os.startfile(str(final_pdf))
                else:
                    subprocess.run(['xdg-open', str(final_pdf)])
                    
        except Exception as e:
            logger.exception("EML conversion failed")
            self._set_ui_state(True)
            self._update_status("Conversion failed!")
            messagebox.showerror("Error", f"Failed to convert EML file:\n\n{str(e)}")

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
