#!/usr/bin/env python3
"""
Mayo's Mail Converter - PST to PDF Converter

Main entry point for the application.
"""

import sys
import os
import logging
from pathlib import Path

# Add project root to path for imports
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))


def setup_logging():
    """Configure application logging."""
    log_dir = PROJECT_ROOT / "logs"
    log_dir.mkdir(exist_ok=True)
    
    log_file = log_dir / "mail_converter.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    # Reduce noise from some libraries
    logging.getLogger('PIL').setLevel(logging.WARNING)
    logging.getLogger('pdfminer').setLevel(logging.WARNING)


def check_dependencies():
    """Check for required dependencies and warn if missing."""
    import shutil
    
    warnings = []
    
    # Check for readpst
    if not shutil.which("readpst"):
        warnings.append(
            "readpst (libpst) not found. PST extraction will not work.\n"
            "Install with: brew install libpst (macOS) or apt-get install pst-utils (Linux)"
        )
    
    # Check for tesseract
    if not shutil.which("tesseract"):
        warnings.append(
            "Tesseract not found. OCR will be disabled.\n"
            "Install with: brew install tesseract (macOS) or apt-get install tesseract-ocr (Linux)"
        )
    
    # Check for LibreOffice (including macOS app bundle location)
    macos_libreoffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    has_libreoffice = (
        shutil.which("libreoffice") or 
        shutil.which("soffice") or
        os.path.isfile(macos_libreoffice)
    )
    if not has_libreoffice:
        warnings.append(
            "LibreOffice not found. DOC/PPT/XLS conversion may be limited.\n"
            "Install from: https://www.libreoffice.org/download/"
        )
    
    return warnings


def main():
    """Main entry point."""
    setup_logging()
    logger = logging.getLogger(__name__)
    
    logger.info("Starting Mayo's Mail Converter")
    
    # Check dependencies
    dep_warnings = check_dependencies()
    for warning in dep_warnings:
        logger.warning(warning)
    
    # Import tkinter and create app
    try:
        import tkinter as tk
        from gui.main_window import MainWindow
        
        # Create root window
        root = tk.Tk()
        
        # Set app icon (if available)
        icon_path = PROJECT_ROOT / "assets" / "icon.png"
        if icon_path.exists():
            try:
                icon = tk.PhotoImage(file=str(icon_path))
                root.iconphoto(True, icon)
            except Exception as e:
                logger.debug(f"Could not load icon: {e}")
        
        # Create main window
        app = MainWindow(root)
        
        # Show dependency warnings in UI
        if dep_warnings:
            from tkinter import messagebox
            warning_text = "Some dependencies are missing:\n\n" + "\n\n".join(dep_warnings)
            root.after(500, lambda: messagebox.showwarning("Missing Dependencies", warning_text))
        
        # Run the application
        root.mainloop()
        
    except ImportError as e:
        logger.error(f"Import error: {e}")
        print(f"\nError: Missing required module: {e}")
        print("Please install dependencies with: pip install -r requirements.txt")
        sys.exit(1)
    
    except Exception as e:
        logger.exception(f"Application error: {e}")
        raise
    
    logger.info("Mayo's Mail Converter closed")


if __name__ == "__main__":
    main()
