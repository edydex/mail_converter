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


def setup_bundled_environment():
    """Set up environment for bundled GTK libraries on Windows."""
    if not hasattr(sys, '_MEIPASS'):
        return  # Not running as bundled app
    
    if sys.platform != 'win32':
        return  # Only needed on Windows
    
    base_path = Path(sys._MEIPASS)
    gtk_path = base_path / 'gtk'
    
    if not gtk_path.exists():
        return  # No GTK libraries bundled
    
    # Add GTK to PATH so DLLs can be found
    os.environ['PATH'] = str(gtk_path) + os.pathsep + os.environ.get('PATH', '')
    
    # Set GDK-Pixbuf loader path
    pixbuf_loaders = gtk_path / 'gdk-pixbuf-2.0'
    if pixbuf_loaders.exists():
        os.environ['GDK_PIXBUF_MODULE_FILE'] = str(pixbuf_loaders / 'loaders.cache')
    
    # Set fontconfig path
    fonts_dir = gtk_path / 'etc' / 'fonts'
    if fonts_dir.exists():
        os.environ['FONTCONFIG_PATH'] = str(fonts_dir)
    
    # Set GLib schemas path
    schemas_dir = gtk_path / 'share' / 'glib-2.0' / 'schemas'
    if schemas_dir.exists():
        os.environ['GSETTINGS_SCHEMA_DIR'] = str(schemas_dir)


# Set up bundled environment BEFORE any other imports
setup_bundled_environment()


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
