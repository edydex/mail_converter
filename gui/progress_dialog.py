"""
Progress Dialog Module

Modal dialog showing conversion progress.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional


class ProgressDialog:
    """Modal dialog showing conversion progress."""
    
    def __init__(self, parent: tk.Tk):
        """
        Initialize progress dialog.
        
        Args:
            parent: Parent window
        """
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Converting...")
        self.dialog.geometry("450x200")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        
        # Make modal
        self.dialog.grab_set()
        
        # Prevent closing via X button
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self._create_widgets()
        self._center_on_parent(parent)
    
    def _create_widgets(self):
        """Create dialog widgets."""
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Stage label
        self.stage_label = ttk.Label(
            main_frame,
            text="Initializing...",
            font=('Helvetica', 12, 'bold')
        )
        self.stage_label.pack(pady=(0, 10))
        
        # Progress bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(pady=10, fill=tk.X)
        
        # Percentage label
        self.percent_label = ttk.Label(
            main_frame,
            text="0%"
        )
        self.percent_label.pack()
        
        # Detail message
        self.detail_label = ttk.Label(
            main_frame,
            text="",
            foreground="gray",
            wraplength=400
        )
        self.detail_label.pack(pady=(10, 0))
        
        # Animated indicator
        self.activity_var = tk.StringVar(value="")
        self.activity_label = ttk.Label(
            main_frame,
            textvariable=self.activity_var,
            foreground="gray"
        )
        self.activity_label.pack(pady=(5, 0))
        
        self._animate()
    
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
    
    def _animate(self):
        """Animate activity indicator."""
        if not self.dialog.winfo_exists():
            return
        
        current = self.activity_var.get()
        dots = current.count('.')
        
        if dots >= 3:
            self.activity_var.set("")
        else:
            self.activity_var.set(current + ".")
        
        self.dialog.after(500, self._animate)
    
    def update_progress(
        self,
        percentage: float,
        message: str,
        stage: str
    ):
        """
        Update progress display.
        
        Args:
            percentage: Progress percentage (0-100)
            message: Detail message
            stage: Current stage name
        """
        if not self.dialog.winfo_exists():
            return
        
        self.progress_var.set(percentage)
        self.percent_label.config(text=f"{percentage:.0f}%")
        self.detail_label.config(text=message)
        
        # Format stage name
        stage_display = stage.replace('_', ' ').title()
        self.stage_label.config(text=stage_display)
        
        self.dialog.update_idletasks()
    
    def close(self):
        """Close the dialog."""
        if self.dialog.winfo_exists():
            self.dialog.grab_release()
            self.dialog.destroy()


class IndeterminateProgressDialog:
    """Progress dialog with indeterminate progress bar."""
    
    def __init__(self, parent: tk.Tk, title: str = "Processing..."):
        """
        Initialize indeterminate progress dialog.
        
        Args:
            parent: Parent window
            title: Dialog title
        """
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("350x120")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self._create_widgets()
        self._center_on_parent(parent)
        self._start_animation()
    
    def _create_widgets(self):
        """Create dialog widgets."""
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message
        self.message_label = ttk.Label(
            main_frame,
            text="Please wait...",
            font=('Helvetica', 11)
        )
        self.message_label.pack(pady=(0, 15))
        
        # Indeterminate progress bar
        self.progress_bar = ttk.Progressbar(
            main_frame,
            length=300,
            mode='indeterminate'
        )
        self.progress_bar.pack()
    
    def _center_on_parent(self, parent: tk.Tk):
        """Center dialog on parent."""
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
    
    def _start_animation(self):
        """Start progress bar animation."""
        self.progress_bar.start(10)
    
    def update_message(self, message: str):
        """Update the message displayed."""
        if self.dialog.winfo_exists():
            self.message_label.config(text=message)
            self.dialog.update_idletasks()
    
    def close(self):
        """Close the dialog."""
        if self.dialog.winfo_exists():
            self.progress_bar.stop()
            self.dialog.grab_release()
            self.dialog.destroy()
