"""
Email Tools Tab Module

GUI tab for Email Tools: Compare, Merge, Deduplicate, Filter, Convert.
"""

import os
import threading
import logging
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Callable

from core.mailbox_comparator import MailboxComparator, ComparisonConfig, ComparisonResult
from core.mailbox_merger import MailboxMerger, MergeConfig, MergeResult
from core.mailbox_deduplicator import MailboxDeduplicator, DedupeConfig, DedupeResult
from core.mailbox_filter import MailboxFilter, FilterConfig, FilterResult
from core.mailbox_writer import OutputFormat, is_pst_write_available

logger = logging.getLogger(__name__)


class EmailToolsTab:
    """Tab containing Email Tools functionality."""
    
    def __init__(self, parent: ttk.Frame, progress_callback: Callable = None):
        """
        Initialize the Email Tools tab.
        
        Args:
            parent: Parent frame (tab container)
            progress_callback: Optional callback for progress updates
        """
        self.parent = parent
        self.progress_callback = progress_callback
        
        # State
        self.current_operation = None
        self.operation_thread = None
        
        # Check PST availability
        self.pst_available = is_pst_write_available()
        
        # Create sub-notebook for tools
        self._create_widgets()
    
    def _create_widgets(self):
        """Create the tools interface."""
        # Main container
        main_frame = ttk.Frame(self.parent, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tools notebook (sub-tabs for each tool)
        self.tools_notebook = ttk.Notebook(main_frame)
        self.tools_notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create tabs for each tool
        self._create_compare_tab()
        self._create_merge_tab()
        self._create_dedupe_tab()
        self._create_filter_tab()
        self._create_convert_tab()
        
        # Status bar
        self.status_frame = ttk.Frame(main_frame)
        self.status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(
            self.status_frame,
            text="Select a tool to begin.",
            foreground="gray"
        )
        self.status_label.pack(side=tk.LEFT)
    
    def _get_output_formats(self) -> list:
        """Get available output formats."""
        formats = ["EML Folder", "MBOX"]
        if self.pst_available:
            formats.append("PST")
        return formats
    
    def _format_to_enum(self, format_str: str) -> OutputFormat:
        """Convert format string to OutputFormat enum."""
        mapping = {
            "EML Folder": OutputFormat.EML_FOLDER,
            "MBOX": OutputFormat.MBOX,
            "PST": OutputFormat.PST
        }
        return mapping.get(format_str, OutputFormat.EML_FOLDER)
    
    def _update_status(self, message: str):
        """Update status label."""
        self.status_label.config(text=message)
        self.parent.update_idletasks()
    
    # =========================================================================
    # COMPARE TAB
    # =========================================================================
    
    def _create_compare_tab(self):
        """Create the Compare Mailboxes tab."""
        frame = ttk.Frame(self.tools_notebook, padding="15")
        self.tools_notebook.add(frame, text="Compare")
        
        # Description
        desc = ttk.Label(
            frame,
            text="Compare two mailboxes to find common and unique emails.",
            wraplength=500
        )
        desc.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))
        
        # Mailbox A
        ttk.Label(frame, text="Mailbox A:").grid(row=1, column=0, sticky="w", pady=5)
        self.compare_a_entry = ttk.Entry(frame, width=50)
        self.compare_a_entry.grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_mailbox(self.compare_a_entry)
        ).grid(row=1, column=2)
        
        # Mailbox B
        ttk.Label(frame, text="Mailbox B:").grid(row=2, column=0, sticky="w", pady=5)
        self.compare_b_entry = ttk.Entry(frame, width=50)
        self.compare_b_entry.grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_mailbox(self.compare_b_entry)
        ).grid(row=2, column=2)
        
        # Output folder
        ttk.Label(frame, text="Output Folder:").grid(row=3, column=0, sticky="w", pady=5)
        self.compare_output_entry = ttk.Entry(frame, width=50)
        self.compare_output_entry.grid(row=3, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_folder(self.compare_output_entry)
        ).grid(row=3, column=2)
        
        # Output format
        ttk.Label(frame, text="Output Format:").grid(row=4, column=0, sticky="w", pady=5)
        self.compare_format_var = tk.StringVar(value="EML Folder")
        format_combo = ttk.Combobox(
            frame, textvariable=self.compare_format_var,
            values=self._get_output_formats(), state="readonly", width=15
        )
        format_combo.grid(row=4, column=1, sticky="w", padx=5)
        
        # Options frame
        options_frame = ttk.LabelFrame(frame, text="Matching Options", padding=10)
        options_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=15)
        
        self.compare_msgid_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, text="Match by Message-ID",
            variable=self.compare_msgid_var
        ).grid(row=0, column=0, sticky="w")
        
        self.compare_content_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, text="Match by Content (sender, subject, body)",
            variable=self.compare_content_var
        ).grid(row=1, column=0, sticky="w")
        
        tolerance_frame = ttk.Frame(options_frame)
        tolerance_frame.grid(row=2, column=0, sticky="w", pady=5)
        ttk.Label(tolerance_frame, text="Timestamp tolerance:").pack(side=tk.LEFT)
        self.compare_tolerance_var = tk.StringVar(value="15")
        ttk.Entry(
            tolerance_frame, textvariable=self.compare_tolerance_var, width=5
        ).pack(side=tk.LEFT, padx=5)
        ttk.Label(tolerance_frame, text="seconds").pack(side=tk.LEFT)
        
        # Run button
        self.compare_btn = ttk.Button(
            frame, text="Compare Mailboxes",
            command=self._run_compare
        )
        self.compare_btn.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Results info
        self.compare_result_label = ttk.Label(frame, text="", foreground="gray")
        self.compare_result_label.grid(row=7, column=0, columnspan=3)
        
        frame.columnconfigure(1, weight=1)
    
    def _run_compare(self):
        """Run mailbox comparison."""
        # Validate inputs
        mailbox_a = self.compare_a_entry.get().strip()
        mailbox_b = self.compare_b_entry.get().strip()
        output_dir = self.compare_output_entry.get().strip()
        
        if not mailbox_a or not os.path.exists(mailbox_a):
            messagebox.showerror("Error", "Please select a valid Mailbox A")
            return
        if not mailbox_b or not os.path.exists(mailbox_b):
            messagebox.showerror("Error", "Please select a valid Mailbox B")
            return
        if not output_dir:
            messagebox.showerror("Error", "Please select an output folder")
            return
        
        # Create config
        config = ComparisonConfig(
            use_message_id=self.compare_msgid_var.get(),
            use_content=self.compare_content_var.get(),
            timestamp_tolerance_seconds=int(self.compare_tolerance_var.get() or 15),
            output_format=self._format_to_enum(self.compare_format_var.get())
        )
        
        # Disable button
        self.compare_btn.config(state=tk.DISABLED)
        self.compare_result_label.config(text="Comparing... please wait")
        
        # Run in thread
        def run():
            try:
                comparator = MailboxComparator(
                    progress_callback=self._progress_callback
                )
                result = comparator.compare(mailbox_a, mailbox_b, output_dir, config)
                
                # Update UI in main thread
                self.parent.after(0, lambda: self._compare_complete(result))
            except Exception as e:
                self.parent.after(0, lambda: self._operation_error(str(e)))
        
        self.operation_thread = threading.Thread(target=run, daemon=True)
        self.operation_thread.start()
    
    def _compare_complete(self, result: ComparisonResult):
        """Handle comparison completion."""
        self.compare_btn.config(state=tk.NORMAL)
        
        if result.success:
            msg = (
                f"✓ Complete! Common: {result.common_count}, "
                f"Unique to A: {result.unique_to_a_count}, "
                f"Unique to B: {result.unique_to_b_count}"
            )
            self.compare_result_label.config(text=msg, foreground="green")
            
            # Show detailed message
            messagebox.showinfo(
                "Comparison Complete",
                f"Mailbox A: {result.total_in_a} emails\n"
                f"Mailbox B: {result.total_in_b} emails\n\n"
                f"Common: {result.common_count}\n"
                f"Unique to A: {result.unique_to_a_count}\n"
                f"Unique to B: {result.unique_to_b_count}\n\n"
                f"Output saved to: {self.compare_output_entry.get()}"
            )
        else:
            self.compare_result_label.config(
                text=f"✗ Failed: {', '.join(result.errors)}", 
                foreground="red"
            )
    
    # =========================================================================
    # MERGE TAB
    # =========================================================================
    
    def _create_merge_tab(self):
        """Create the Merge Mailboxes tab."""
        frame = ttk.Frame(self.tools_notebook, padding="15")
        self.tools_notebook.add(frame, text="Merge")
        
        # Description
        desc = ttk.Label(
            frame,
            text="Merge multiple mailboxes into one, with optional deduplication.",
            wraplength=500
        )
        desc.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))
        
        # Input mailboxes (listbox)
        ttk.Label(frame, text="Mailboxes to merge:").grid(row=1, column=0, sticky="nw", pady=5)
        
        list_frame = ttk.Frame(frame)
        list_frame.grid(row=1, column=1, sticky="nsew", padx=5)
        
        self.merge_listbox = tk.Listbox(list_frame, height=6, width=50)
        self.merge_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.merge_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.merge_listbox.config(yscrollcommand=scrollbar.set)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=2, sticky="n", pady=5)
        ttk.Button(btn_frame, text="Add...", width=8, command=self._merge_add).pack(pady=2)
        ttk.Button(btn_frame, text="Remove", width=8, command=self._merge_remove).pack(pady=2)
        ttk.Button(btn_frame, text="Clear", width=8, command=self._merge_clear).pack(pady=2)
        
        # Output file
        ttk.Label(frame, text="Output:").grid(row=2, column=0, sticky="w", pady=5)
        self.merge_output_entry = ttk.Entry(frame, width=50)
        self.merge_output_entry.grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_save(self.merge_output_entry)
        ).grid(row=2, column=2)
        
        # Output format
        ttk.Label(frame, text="Output Format:").grid(row=3, column=0, sticky="w", pady=5)
        self.merge_format_var = tk.StringVar(value="MBOX")
        format_combo = ttk.Combobox(
            frame, textvariable=self.merge_format_var,
            values=self._get_output_formats(), state="readonly", width=15
        )
        format_combo.grid(row=3, column=1, sticky="w", padx=5)
        
        # Options
        options_frame = ttk.LabelFrame(frame, text="Options", padding=10)
        options_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=15)
        
        self.merge_dedupe_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, text="Remove duplicates during merge",
            variable=self.merge_dedupe_var
        ).pack(anchor="w")
        
        # Run button
        self.merge_btn = ttk.Button(
            frame, text="Merge Mailboxes",
            command=self._run_merge
        )
        self.merge_btn.grid(row=5, column=0, columnspan=3, pady=20)
        
        self.merge_result_label = ttk.Label(frame, text="", foreground="gray")
        self.merge_result_label.grid(row=6, column=0, columnspan=3)
        
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)
    
    def _merge_add(self):
        """Add mailbox to merge list."""
        paths = filedialog.askopenfilenames(
            title="Select Mailbox(es)",
            filetypes=[
                ("All Mailboxes", "*.pst *.mbox"),
                ("PST Files", "*.pst"),
                ("MBOX Files", "*.mbox"),
                ("All Files", "*.*")
            ]
        )
        for p in paths:
            self.merge_listbox.insert(tk.END, p)
    
    def _merge_remove(self):
        """Remove selected mailbox from list."""
        selection = self.merge_listbox.curselection()
        for idx in reversed(selection):
            self.merge_listbox.delete(idx)
    
    def _merge_clear(self):
        """Clear merge list."""
        self.merge_listbox.delete(0, tk.END)
    
    def _run_merge(self):
        """Run mailbox merge."""
        paths = list(self.merge_listbox.get(0, tk.END))
        output = self.merge_output_entry.get().strip()
        
        if len(paths) < 2:
            messagebox.showerror("Error", "Please add at least 2 mailboxes to merge")
            return
        if not output:
            messagebox.showerror("Error", "Please specify output location")
            return
        
        config = MergeConfig(
            deduplicate=self.merge_dedupe_var.get(),
            output_format=self._format_to_enum(self.merge_format_var.get())
        )
        
        self.merge_btn.config(state=tk.DISABLED)
        self.merge_result_label.config(text="Merging... please wait")
        
        def run():
            try:
                merger = MailboxMerger(progress_callback=self._progress_callback)
                result = merger.merge(paths, output, config)
                self.parent.after(0, lambda: self._merge_complete(result))
            except Exception as e:
                self.parent.after(0, lambda: self._operation_error(str(e)))
        
        self.operation_thread = threading.Thread(target=run, daemon=True)
        self.operation_thread.start()
    
    def _merge_complete(self, result: MergeResult):
        """Handle merge completion."""
        self.merge_btn.config(state=tk.NORMAL)
        
        if result.success:
            msg = (
                f"✓ Merged {result.emails_written} emails "
                f"({result.duplicates_removed} duplicates removed)"
            )
            self.merge_result_label.config(text=msg, foreground="green")
            messagebox.showinfo("Merge Complete", msg)
        else:
            self.merge_result_label.config(
                text=f"✗ Failed: {', '.join(result.errors)}",
                foreground="red"
            )
    
    # =========================================================================
    # DEDUPE TAB
    # =========================================================================
    
    def _create_dedupe_tab(self):
        """Create the Deduplicate tab."""
        frame = ttk.Frame(self.tools_notebook, padding="15")
        self.tools_notebook.add(frame, text="Deduplicate")
        
        desc = ttk.Label(
            frame,
            text="Remove duplicate emails from a mailbox.",
            wraplength=500
        )
        desc.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))
        
        # Input
        ttk.Label(frame, text="Input Mailbox:").grid(row=1, column=0, sticky="w", pady=5)
        self.dedupe_input_entry = ttk.Entry(frame, width=50)
        self.dedupe_input_entry.grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_mailbox(self.dedupe_input_entry)
        ).grid(row=1, column=2)
        
        # Output
        ttk.Label(frame, text="Output:").grid(row=2, column=0, sticky="w", pady=5)
        self.dedupe_output_entry = ttk.Entry(frame, width=50)
        self.dedupe_output_entry.grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_save(self.dedupe_output_entry)
        ).grid(row=2, column=2)
        
        # Format
        ttk.Label(frame, text="Output Format:").grid(row=3, column=0, sticky="w", pady=5)
        self.dedupe_format_var = tk.StringVar(value="MBOX")
        ttk.Combobox(
            frame, textvariable=self.dedupe_format_var,
            values=self._get_output_formats(), state="readonly", width=15
        ).grid(row=3, column=1, sticky="w", padx=5)
        
        # Options
        options_frame = ttk.LabelFrame(frame, text="Options", padding=10)
        options_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=15)
        
        self.dedupe_keep_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame, text="Also save duplicates to separate file",
            variable=self.dedupe_keep_var
        ).pack(anchor="w")
        
        # Run
        self.dedupe_btn = ttk.Button(
            frame, text="Remove Duplicates",
            command=self._run_dedupe
        )
        self.dedupe_btn.grid(row=5, column=0, columnspan=3, pady=20)
        
        self.dedupe_result_label = ttk.Label(frame, text="", foreground="gray")
        self.dedupe_result_label.grid(row=6, column=0, columnspan=3)
        
        frame.columnconfigure(1, weight=1)
    
    def _run_dedupe(self):
        """Run deduplication."""
        input_path = self.dedupe_input_entry.get().strip()
        output_path = self.dedupe_output_entry.get().strip()
        
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid input mailbox")
            return
        if not output_path:
            messagebox.showerror("Error", "Please specify output location")
            return
        
        config = DedupeConfig(
            output_format=self._format_to_enum(self.dedupe_format_var.get()),
            keep_duplicates=self.dedupe_keep_var.get()
        )
        
        self.dedupe_btn.config(state=tk.DISABLED)
        self.dedupe_result_label.config(text="Deduplicating... please wait")
        
        def run():
            try:
                deduper = MailboxDeduplicator(progress_callback=self._progress_callback)
                result = deduper.deduplicate(input_path, output_path, config)
                self.parent.after(0, lambda: self._dedupe_complete(result))
            except Exception as e:
                self.parent.after(0, lambda: self._operation_error(str(e)))
        
        self.operation_thread = threading.Thread(target=run, daemon=True)
        self.operation_thread.start()
    
    def _dedupe_complete(self, result: DedupeResult):
        """Handle dedupe completion."""
        self.dedupe_btn.config(state=tk.NORMAL)
        
        if result.success:
            msg = (
                f"✓ Found {result.duplicates_found} duplicates. "
                f"{result.unique_emails} unique emails saved."
            )
            self.dedupe_result_label.config(text=msg, foreground="green")
            messagebox.showinfo("Deduplication Complete", msg)
        else:
            self.dedupe_result_label.config(
                text=f"✗ Failed: {', '.join(result.errors)}",
                foreground="red"
            )
    
    # =========================================================================
    # FILTER TAB
    # =========================================================================
    
    def _create_filter_tab(self):
        """Create the Filter by Sender/Recipient tab."""
        frame = ttk.Frame(self.tools_notebook, padding="15")
        self.tools_notebook.add(frame, text="Filter")
        
        desc = ttk.Label(
            frame,
            text="Filter emails by sender or recipient email addresses or domains.",
            wraplength=500
        )
        desc.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))
        
        # Input
        ttk.Label(frame, text="Input Mailbox:").grid(row=1, column=0, sticky="w", pady=5)
        self.filter_input_entry = ttk.Entry(frame, width=50)
        self.filter_input_entry.grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_mailbox(self.filter_input_entry)
        ).grid(row=1, column=2)
        
        # Output
        ttk.Label(frame, text="Output:").grid(row=2, column=0, sticky="w", pady=5)
        self.filter_output_entry = ttk.Entry(frame, width=50)
        self.filter_output_entry.grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_folder(self.filter_output_entry)
        ).grid(row=2, column=2)
        
        # Filter criteria
        criteria_frame = ttk.LabelFrame(frame, text="Filter Criteria", padding=10)
        criteria_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)
        
        ttk.Label(
            criteria_frame, 
            text="Enter email addresses or domains (comma-separated):"
        ).grid(row=0, column=0, columnspan=2, sticky="w")
        
        ttk.Label(criteria_frame, text="Sender emails:").grid(row=1, column=0, sticky="w", pady=3)
        self.filter_sender_email_entry = ttk.Entry(criteria_frame, width=40)
        self.filter_sender_email_entry.grid(row=1, column=1, sticky="ew", padx=5)
        
        ttk.Label(criteria_frame, text="Sender domains:").grid(row=2, column=0, sticky="w", pady=3)
        self.filter_sender_domain_entry = ttk.Entry(criteria_frame, width=40)
        self.filter_sender_domain_entry.grid(row=2, column=1, sticky="ew", padx=5)
        
        ttk.Label(criteria_frame, text="Recipient emails:").grid(row=3, column=0, sticky="w", pady=3)
        self.filter_recipient_email_entry = ttk.Entry(criteria_frame, width=40)
        self.filter_recipient_email_entry.grid(row=3, column=1, sticky="ew", padx=5)
        
        ttk.Label(criteria_frame, text="Recipient domains:").grid(row=4, column=0, sticky="w", pady=3)
        self.filter_recipient_domain_entry = ttk.Entry(criteria_frame, width=40)
        self.filter_recipient_domain_entry.grid(row=4, column=1, sticky="ew", padx=5)
        
        criteria_frame.columnconfigure(1, weight=1)
        
        # Format
        ttk.Label(frame, text="Output Format:").grid(row=4, column=0, sticky="w", pady=5)
        self.filter_format_var = tk.StringVar(value="EML Folder")
        ttk.Combobox(
            frame, textvariable=self.filter_format_var,
            values=self._get_output_formats(), state="readonly", width=15
        ).grid(row=4, column=1, sticky="w", padx=5)
        
        # Run
        self.filter_btn = ttk.Button(
            frame, text="Filter Emails",
            command=self._run_filter
        )
        self.filter_btn.grid(row=5, column=0, columnspan=3, pady=20)
        
        self.filter_result_label = ttk.Label(frame, text="", foreground="gray")
        self.filter_result_label.grid(row=6, column=0, columnspan=3)
        
        frame.columnconfigure(1, weight=1)
    
    def _run_filter(self):
        """Run email filter."""
        input_path = self.filter_input_entry.get().strip()
        output_path = self.filter_output_entry.get().strip()
        
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid input mailbox")
            return
        if not output_path:
            messagebox.showerror("Error", "Please specify output location")
            return
        
        # Parse filter criteria
        def parse_list(entry):
            text = entry.get().strip()
            if not text:
                return []
            return [x.strip() for x in text.split(",") if x.strip()]
        
        config = FilterConfig(
            sender_emails=parse_list(self.filter_sender_email_entry),
            sender_domains=parse_list(self.filter_sender_domain_entry),
            recipient_emails=parse_list(self.filter_recipient_email_entry),
            recipient_domains=parse_list(self.filter_recipient_domain_entry),
            output_format=self._format_to_enum(self.filter_format_var.get())
        )
        
        # Check at least one criterion
        if not (config.sender_emails or config.sender_domains or 
                config.recipient_emails or config.recipient_domains):
            messagebox.showerror("Error", "Please enter at least one filter criterion")
            return
        
        self.filter_btn.config(state=tk.DISABLED)
        self.filter_result_label.config(text="Filtering... please wait")
        
        def run():
            try:
                filterer = MailboxFilter(progress_callback=self._progress_callback)
                result = filterer.filter(input_path, output_path, config)
                self.parent.after(0, lambda: self._filter_complete(result))
            except Exception as e:
                self.parent.after(0, lambda: self._operation_error(str(e)))
        
        self.operation_thread = threading.Thread(target=run, daemon=True)
        self.operation_thread.start()
    
    def _filter_complete(self, result: FilterResult):
        """Handle filter completion."""
        self.filter_btn.config(state=tk.NORMAL)
        
        if result.success:
            msg = f"✓ Found {result.matched_emails} matching emails out of {result.total_emails}"
            self.filter_result_label.config(text=msg, foreground="green")
            messagebox.showinfo("Filter Complete", msg)
        else:
            self.filter_result_label.config(
                text=f"✗ Failed: {', '.join(result.errors)}",
                foreground="red"
            )
    
    # =========================================================================
    # CONVERT TAB
    # =========================================================================
    
    def _create_convert_tab(self):
        """Create the Convert Format tab."""
        frame = ttk.Frame(self.tools_notebook, padding="15")
        self.tools_notebook.add(frame, text="Convert")
        
        desc = ttk.Label(
            frame,
            text="Convert between mailbox formats (PST, MBOX, EML folder).",
            wraplength=500
        )
        desc.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))
        
        # Input
        ttk.Label(frame, text="Input:").grid(row=1, column=0, sticky="w", pady=5)
        self.convert_input_entry = ttk.Entry(frame, width=50)
        self.convert_input_entry.grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_mailbox(self.convert_input_entry)
        ).grid(row=1, column=2)
        
        # Output
        ttk.Label(frame, text="Output:").grid(row=2, column=0, sticky="w", pady=5)
        self.convert_output_entry = ttk.Entry(frame, width=50)
        self.convert_output_entry.grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(
            frame, text="Browse...", width=10,
            command=lambda: self._browse_save(self.convert_output_entry)
        ).grid(row=2, column=2)
        
        # Format
        ttk.Label(frame, text="Output Format:").grid(row=3, column=0, sticky="w", pady=5)
        self.convert_format_var = tk.StringVar(value="MBOX")
        ttk.Combobox(
            frame, textvariable=self.convert_format_var,
            values=self._get_output_formats(), state="readonly", width=15
        ).grid(row=3, column=1, sticky="w", padx=5)
        
        # Run
        self.convert_btn = ttk.Button(
            frame, text="Convert",
            command=self._run_convert
        )
        self.convert_btn.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.convert_result_label = ttk.Label(frame, text="", foreground="gray")
        self.convert_result_label.grid(row=5, column=0, columnspan=3)
        
        frame.columnconfigure(1, weight=1)
    
    def _run_convert(self):
        """Run format conversion."""
        input_path = self.convert_input_entry.get().strip()
        output_path = self.convert_output_entry.get().strip()
        
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid input")
            return
        if not output_path:
            messagebox.showerror("Error", "Please specify output location")
            return
        
        # Use merger with dedupe=False for simple conversion
        config = MergeConfig(
            deduplicate=False,
            output_format=self._format_to_enum(self.convert_format_var.get())
        )
        
        self.convert_btn.config(state=tk.DISABLED)
        self.convert_result_label.config(text="Converting... please wait")
        
        def run():
            try:
                merger = MailboxMerger(progress_callback=self._progress_callback)
                result = merger.merge([input_path], output_path, config)
                self.parent.after(0, lambda: self._convert_complete(result))
            except Exception as e:
                self.parent.after(0, lambda: self._operation_error(str(e)))
        
        self.operation_thread = threading.Thread(target=run, daemon=True)
        self.operation_thread.start()
    
    def _convert_complete(self, result: MergeResult):
        """Handle conversion completion."""
        self.convert_btn.config(state=tk.NORMAL)
        
        if result.success:
            msg = f"✓ Converted {result.emails_written} emails"
            self.convert_result_label.config(text=msg, foreground="green")
            messagebox.showinfo("Conversion Complete", msg)
        else:
            self.convert_result_label.config(
                text=f"✗ Failed: {', '.join(result.errors)}",
                foreground="red"
            )
    
    # =========================================================================
    # HELPERS
    # =========================================================================
    
    def _browse_mailbox(self, entry: ttk.Entry):
        """Browse for a mailbox file or folder."""
        # Ask user if file or folder
        choice = messagebox.askquestion(
            "Select Input Type",
            "Do you want to select a file (PST/MBOX)?\n\nClick 'Yes' for file, 'No' for folder."
        )
        
        if choice == 'yes':
            path = filedialog.askopenfilename(
                title="Select Mailbox File",
                filetypes=[
                    ("All Mailboxes", "*.pst *.mbox"),
                    ("PST Files", "*.pst"),
                    ("MBOX Files", "*.mbox"),
                    ("All Files", "*.*")
                ]
            )
        else:
            path = filedialog.askdirectory(title="Select EML Folder")
        
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
    
    def _browse_folder(self, entry: ttk.Entry):
        """Browse for output folder."""
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
    
    def _browse_save(self, entry: ttk.Entry):
        """Browse for save location."""
        path = filedialog.asksaveasfilename(
            title="Save As",
            filetypes=[
                ("MBOX Files", "*.mbox"),
                ("PST Files", "*.pst"),
                ("All Files", "*.*")
            ]
        )
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
    
    def _progress_callback(self, current: int, total: int, message: str):
        """Handle progress updates from operations."""
        self._update_status(message)
    
    def _operation_error(self, error_msg: str):
        """Handle operation error."""
        # Re-enable all buttons
        for btn in [self.compare_btn, self.merge_btn, self.dedupe_btn, 
                    self.filter_btn, self.convert_btn]:
            btn.config(state=tk.NORMAL)
        
        self._update_status(f"Error: {error_msg}")
        messagebox.showerror("Error", f"Operation failed: {error_msg}")
