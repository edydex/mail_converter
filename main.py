#!/usr/bin/env python3
"""
Mayo's Mail Converter - PST to PDF Converter

Main entry point for the application.
"""

import sys
import os
import logging
from pathlib import Path

# Set up bundled library paths BEFORE any other imports
# This must happen before WeasyPrint tries to load GTK/Pango DLLs
def setup_bundled_paths():
    """Configure paths for bundled DLLs on Windows."""
    if sys.platform != 'win32':
        return
    
    # Check if running as PyInstaller bundle
    if hasattr(sys, '_MEIPASS'):
        base_path = Path(sys._MEIPASS)
    else:
        return  # Only needed for bundled app
    
    # Set WEASYPRINT_DLL_DIRECTORIES for GTK/Pango DLLs
    gtk_path = base_path / 'gtk'
    if gtk_path.exists():
        os.environ['WEASYPRINT_DLL_DIRECTORIES'] = str(gtk_path)
    
    # Also add to PATH for other DLLs
    bin_path = base_path / 'bin'
    paths_to_add = []
    if gtk_path.exists():
        paths_to_add.append(str(gtk_path))
    if bin_path.exists():
        paths_to_add.append(str(bin_path))
    
    if paths_to_add:
        os.environ['PATH'] = os.pathsep.join(paths_to_add) + os.pathsep + os.environ.get('PATH', '')

# MUST be called before importing anything that uses WeasyPrint
setup_bundled_paths()

# Add project root to path for imports
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))


def get_log_directory() -> Path:
    """Get the appropriate log directory based on execution context."""
    # For bundled PyInstaller app, use user's documents folder
    if hasattr(sys, '_MEIPASS'):
        if sys.platform == 'win32':
            # Use Documents/MayosMailConverter/logs on Windows
            docs = Path(os.environ.get('USERPROFILE', '')) / 'Documents'
            log_dir = docs / 'MayosMailConverter' / 'logs'
        else:
            # Use ~/Library/Logs on macOS, ~/.local/share on Linux
            home = Path.home()
            if sys.platform == 'darwin':
                log_dir = home / 'Library' / 'Logs' / 'MayosMailConverter'
            else:
                log_dir = home / '.local' / 'share' / 'MayosMailConverter' / 'logs'
    else:
        # Development mode - use project directory
        log_dir = PROJECT_ROOT / "logs"
    
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        # Fallback to temp directory if we can't create the preferred location
        import tempfile
        log_dir = Path(tempfile.gettempdir()) / 'MayosMailConverter' / 'logs'
        log_dir.mkdir(parents=True, exist_ok=True)
    
    return log_dir


def setup_logging():
    """Configure application logging."""
    log_dir = get_log_directory()
    log_file = log_dir / "mail_converter.log"
    
    # Print log location to console so user knows where to find it
    print(f"Log file: {log_file}")
    
    logging.basicConfig(
        level=logging.DEBUG,  # Changed to DEBUG for troubleshooting
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    # Log the log file location
    logging.getLogger(__name__).info(f"Log file location: {log_file}")
    
    # Reduce noise from some libraries
    logging.getLogger('PIL').setLevel(logging.WARNING)
    logging.getLogger('pdfminer').setLevel(logging.WARNING)
    logging.getLogger('fontTools').setLevel(logging.WARNING)
    logging.getLogger('weasyprint').setLevel(logging.WARNING)


def check_dependencies():
    """Check for required dependencies and warn if missing."""
    import shutil
    import platform
    
    warnings = []
    system = platform.system()
    
    # Check for readpst
    readpst_found = False
    
    # First check if bundled (PyInstaller)
    if hasattr(sys, '_MEIPASS'):
        bundled_paths = [
            os.path.join(sys._MEIPASS, 'bin', 'readpst.exe'),
            os.path.join(sys._MEIPASS, 'readpst.exe'),
            os.path.join(sys._MEIPASS, 'bin', 'readpst'),
            os.path.join(sys._MEIPASS, 'readpst'),
        ]
        for path in bundled_paths:
            if os.path.isfile(path):
                readpst_found = True
                break
    
    # Then check system
    if not readpst_found:
        readpst_found = shutil.which("readpst") is not None or shutil.which("readpst.exe") is not None
    
    if not readpst_found:
        if system == "Windows":
            warnings.append(
                "readpst (libpst) not found. PST extraction will not work.\n"
                "This should be bundled with the application - try re-downloading."
            )
        else:
            warnings.append(
                "readpst (libpst) not found. PST extraction will not work.\n"
                "Install with: brew install libpst (macOS) or apt-get install pst-utils (Linux)"
            )
    
    # Check for tesseract (optional, don't warn on Windows - it's complicated there)
    if system != "Windows":
        if not shutil.which("tesseract"):
            warnings.append(
                "Tesseract not found. OCR will be disabled.\n"
                "Install with: brew install tesseract (macOS) or apt-get install tesseract-ocr (Linux)"
            )
    
    # Check for LibreOffice (including macOS app bundle location)
    macos_libreoffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    windows_libreoffice_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    
    has_libreoffice = (
        shutil.which("libreoffice") or 
        shutil.which("soffice") or
        os.path.isfile(macos_libreoffice) or
        any(os.path.isfile(p) for p in windows_libreoffice_paths)
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
