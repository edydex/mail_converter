"""
Settings Dialog Module

Advanced settings dialog for the Mail Converter application.
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime
from typing import Dict, Any, Optional


class SettingsDialog:
    """Advanced settings dialog."""
    
    def __init__(self, parent: tk.Tk, current_settings: Dict[str, Any]):
        """
        Initialize settings dialog.
        
        Args:
            parent: Parent window
            current_settings: Current settings dictionary
        """
        self.result: Optional[Dict[str, Any]] = None
        self.current = current_settings.copy()
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Advanced Settings")
        self.dialog.geometry("500x450")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._load_current_settings()
        self._center_on_parent(parent)
        
        # Wait for dialog to close
        self.dialog.wait_window()
    
    def _create_widgets(self):
        """Create dialog widgets."""
        # Main notebook for tabs
        notebook = ttk.Notebook(self.dialog)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # === Output Tab ===
        output_frame = ttk.Frame(notebook, padding=15)
        notebook.add(output_frame, text="Output")
        
        # Table of Contents
        self.toc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            output_frame,
            text="Include Table of Contents in combined PDF",
            variable=self.toc_var
        ).pack(anchor=tk.W, pady=5)
        
        # Separator pages
        self.separators_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            output_frame,
            text="Add separator pages between emails",
            variable=self.separators_var
        ).pack(anchor=tk.W, pady=5)
        
        # Attachment separators
        self.att_separators_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            output_frame,
            text="Add separator pages before attachments",
            variable=self.att_separators_var
        ).pack(anchor=tk.W, pady=5)
        
        ttk.Separator(output_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=15)
        
        # Page size
        page_frame = ttk.Frame(output_frame)
        page_frame.pack(anchor=tk.W, pady=5)
        
        ttk.Label(page_frame, text="Page Size:").pack(side=tk.LEFT)
        
        self.page_size_var = tk.StringVar(value="Letter")
        page_combo = ttk.Combobox(
            page_frame,
            textvariable=self.page_size_var,
            values=["Letter", "A4"],
            state="readonly",
            width=10
        )
        page_combo.pack(side=tk.LEFT, padx=10)
        
        # Page margin
        margin_frame = ttk.Frame(output_frame)
        margin_frame.pack(anchor=tk.W, pady=5)
        
        ttk.Label(margin_frame, text="Page Margin:").pack(side=tk.LEFT)
        
        self.page_margin_var = tk.StringVar(value="0.5")
        margin_combo = ttk.Combobox(
            margin_frame,
            textvariable=self.page_margin_var,
            values=["0.25", "0.5", "0.75", "1.0"],
            state="readonly",
            width=10
        )
        margin_combo.pack(side=tk.LEFT, padx=10)
        ttk.Label(margin_frame, text="inches", foreground="gray").pack(side=tk.LEFT)
        
        # === Filtering Tab ===
        filter_frame = ttk.Frame(notebook, padding=15)
        notebook.add(filter_frame, text="Filters")
        
        # Date range
        date_label = ttk.Label(
            filter_frame,
            text="Date Range (leave empty for all emails):",
            font=('Helvetica', 10, 'bold')
        )
        date_label.pack(anchor=tk.W, pady=(0, 10))
        
        # From date
        from_frame = ttk.Frame(filter_frame)
        from_frame.pack(anchor=tk.W, pady=5)
        
        ttk.Label(from_frame, text="From:", width=8).pack(side=tk.LEFT)
        
        self.date_from_var = tk.StringVar()
        self.date_from_entry = ttk.Entry(
            from_frame,
            textvariable=self.date_from_var,
            width=15
        )
        self.date_from_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(from_frame, text="(YYYY-MM-DD)", foreground="gray").pack(side=tk.LEFT)
        
        # To date
        to_frame = ttk.Frame(filter_frame)
        to_frame.pack(anchor=tk.W, pady=5)
        
        ttk.Label(to_frame, text="To:", width=8).pack(side=tk.LEFT)
        
        self.date_to_var = tk.StringVar()
        self.date_to_entry = ttk.Entry(
            to_frame,
            textvariable=self.date_to_var,
            width=15
        )
        self.date_to_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(to_frame, text="(YYYY-MM-DD)", foreground="gray").pack(side=tk.LEFT)
        
        # Clear dates button
        ttk.Button(
            filter_frame,
            text="Clear Dates",
            command=self._clear_dates
        ).pack(anchor=tk.W, pady=10)
        
        ttk.Separator(filter_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=15)
        
        # Future: Sender/Subject filters
        future_label = ttk.Label(
            filter_frame,
            text="Additional filters coming in future versions:\n• Filter by sender\n• Filter by subject keywords\n• Filter by folder",
            foreground="gray"
        )
        future_label.pack(anchor=tk.W, pady=10)
        
        # === Processing Tab ===
        process_frame = ttk.Frame(notebook, padding=15)
        notebook.add(process_frame, text="Processing")
        
        # PST structure
        self.preserve_structure_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            process_frame,
            text="Preserve PST folder structure in extraction",
            variable=self.preserve_structure_var
        ).pack(anchor=tk.W, pady=5)
        
        # Include deleted
        self.include_deleted_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            process_frame,
            text="Include deleted emails (if recoverable)",
            variable=self.include_deleted_var
        ).pack(anchor=tk.W, pady=5)
        
        # Skip Deleted Items folder
        self.skip_deleted_items_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            process_frame,
            text="Skip emails from 'Deleted Items' folder",
            variable=self.skip_deleted_items_var
        ).pack(anchor=tk.W, pady=5)
        
        # Rename EML files for diagnostics
        self.rename_emls_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            process_frame,
            text="Rename extracted emails to YYYYMMDD_HHMMSS_subject.eml",
            variable=self.rename_emls_var
        ).pack(anchor=tk.W, pady=5)
        
        ttk.Separator(process_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=15)
        
        # OCR settings
        ocr_label = ttk.Label(
            process_frame,
            text="OCR Settings:",
            font=('Helvetica', 10, 'bold')
        )
        ocr_label.pack(anchor=tk.W, pady=(0, 10))
        
        # OCR language
        lang_frame = ttk.Frame(process_frame)
        lang_frame.pack(anchor=tk.W, pady=5)
        
        ttk.Label(lang_frame, text="OCR Language:").pack(side=tk.LEFT)
        
        self.ocr_lang_var = tk.StringVar(value="eng")
        lang_combo = ttk.Combobox(
            lang_frame,
            textvariable=self.ocr_lang_var,
            values=["eng", "eng+fra", "eng+deu", "eng+spa", "eng+chi_sim"],
            state="readonly",
            width=15
        )
        lang_combo.pack(side=tk.LEFT, padx=10)
        
        # DPI setting
        dpi_frame = ttk.Frame(process_frame)
        dpi_frame.pack(anchor=tk.W, pady=5)
        
        ttk.Label(dpi_frame, text="OCR DPI:").pack(side=tk.LEFT)
        
        self.ocr_dpi_var = tk.StringVar(value="300")
        dpi_combo = ttk.Combobox(
            dpi_frame,
            textvariable=self.ocr_dpi_var,
            values=["150", "200", "300", "400"],
            state="readonly",
            width=10
        )
        dpi_combo.pack(side=tk.LEFT, padx=10)
        ttk.Label(dpi_frame, text="(higher = better quality, slower)", foreground="gray").pack(side=tk.LEFT)
        
        # === Buttons ===
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            button_frame,
            text="Reset to Defaults",
            command=self._reset_defaults
        ).pack(side=tk.LEFT)
        
        ttk.Button(
            button_frame,
            text="Cancel",
            command=self._cancel
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame,
            text="Save",
            command=self._save
        ).pack(side=tk.RIGHT)
    
    def _load_current_settings(self):
        """Load current settings into UI."""
        self.toc_var.set(self.current.get('add_toc', True))
        self.separators_var.set(self.current.get('add_separators', False))
        self.att_separators_var.set(self.current.get('add_att_separators', False))
        self.page_size_var.set(self.current.get('page_size', 'Letter'))
        self.page_margin_var.set(str(self.current.get('page_margin', 0.5)))
        self.preserve_structure_var.set(self.current.get('preserve_structure', True))
        self.include_deleted_var.set(self.current.get('include_deleted', False))
        self.skip_deleted_items_var.set(self.current.get('skip_deleted_items', True))
        self.rename_emls_var.set(self.current.get('rename_emls', True))
        self.ocr_lang_var.set(self.current.get('ocr_language', 'eng'))
        self.ocr_dpi_var.set(str(self.current.get('ocr_dpi', 300)))
        
        # Date filters
        if self.current.get('date_from'):
            self.date_from_var.set(self.current['date_from'].strftime('%Y-%m-%d'))
        if self.current.get('date_to'):
            self.date_to_var.set(self.current['date_to'].strftime('%Y-%m-%d'))
    
    def _center_on_parent(self, parent: tk.Tk):
        """Center dialog on parent window."""
        self.dialog.update_idletasks()
        
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"+{x}+{y}")
    
    def _clear_dates(self):
        """Clear date filter entries."""
        self.date_from_var.set("")
        self.date_to_var.set("")
    
    def _reset_defaults(self):
        """Reset all settings to defaults."""
        self.toc_var.set(True)
        self.separators_var.set(True)
        self.att_separators_var.set(True)
        self.page_size_var.set("Letter")
        self.page_margin_var.set("0.5")
        self.preserve_structure_var.set(True)
        self.include_deleted_var.set(False)
        self.skip_deleted_items_var.set(True)
        self.rename_emls_var.set(True)
        self.ocr_lang_var.set("eng")
        self.ocr_dpi_var.set("300")
        self._clear_dates()
    
    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse date string to datetime."""
        if not date_str.strip():
            return None
        
        try:
            return datetime.strptime(date_str.strip(), '%Y-%m-%d')
        except ValueError:
            return None
    
    def _save(self):
        """Save settings and close dialog."""
        # Validate dates
        date_from = self._parse_date(self.date_from_var.get())
        date_to = self._parse_date(self.date_to_var.get())
        
        if self.date_from_var.get() and not date_from:
            tk.messagebox.showerror(
                "Invalid Date",
                "Invalid 'From' date format. Use YYYY-MM-DD."
            )
            return
        
        if self.date_to_var.get() and not date_to:
            tk.messagebox.showerror(
                "Invalid Date",
                "Invalid 'To' date format. Use YYYY-MM-DD."
            )
            return
        
        if date_from and date_to and date_from > date_to:
            tk.messagebox.showerror(
                "Invalid Date Range",
                "'From' date must be before 'To' date."
            )
            return
        
        # Collect settings
        self.result = {
            'add_toc': self.toc_var.get(),
            'add_separators': self.separators_var.get(),
            'add_att_separators': self.att_separators_var.get(),
            'page_size': self.page_size_var.get(),
            'page_margin': float(self.page_margin_var.get()),
            'preserve_structure': self.preserve_structure_var.get(),
            'include_deleted': self.include_deleted_var.get(),
            'skip_deleted_items': self.skip_deleted_items_var.get(),
            'rename_emls': self.rename_emls_var.get(),
            'ocr_language': self.ocr_lang_var.get(),
            'ocr_dpi': int(self.ocr_dpi_var.get()),
            'date_from': date_from,
            'date_to': date_to
        }
        
        self.dialog.destroy()
    
    def _cancel(self):
        """Cancel and close dialog without saving."""
        self.result = None
        self.dialog.destroy()
