"""
System Information Diagnostic Module

Collects detailed system information for troubleshooting,
particularly around display scaling and DPI settings.
"""

import sys
import os
import platform
import logging
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

logger = logging.getLogger(__name__)


def get_system_info() -> Dict[str, Any]:
    """
    Collect comprehensive system information for diagnostics.
    
    Returns a dictionary with all relevant system details.
    """
    info = {
        'timestamp': datetime.now().isoformat(),
        'platform': {},
        'python': {},
        'display': {},
        'dpi': {},
        'environment': {},
        'weasyprint': {},
        'libraries': {},
    }
    
    # Platform info
    info['platform'] = {
        'system': platform.system(),
        'release': platform.release(),
        'version': platform.version(),
        'machine': platform.machine(),
        'processor': platform.processor(),
        'node': platform.node(),
    }
    
    # Python info
    info['python'] = {
        'version': sys.version,
        'executable': sys.executable,
        'is_bundled': hasattr(sys, '_MEIPASS'),
        'bundled_path': getattr(sys, '_MEIPASS', None),
    }
    
    # Display/DPI info (Windows-specific)
    if sys.platform == 'win32':
        info['display'] = get_windows_display_info()
        info['dpi'] = get_windows_dpi_info()
    else:
        info['display'] = {'note': 'Display info collection only implemented for Windows'}
        info['dpi'] = {'note': 'DPI info collection only implemented for Windows'}
    
    # Environment variables related to display/scaling
    dpi_env_vars = [
        'GDK_SCALE', 'GDK_DPI_SCALE', 'QT_SCALE_FACTOR', 
        'QT_AUTO_SCREEN_SCALE_FACTOR', 'WEASYPRINT_DLL_DIRECTORIES',
        'GTK_THEME', 'PATH'
    ]
    info['environment'] = {
        var: os.environ.get(var, '<not set>')
        for var in dpi_env_vars
    }
    # Truncate PATH for readability
    if info['environment'].get('PATH'):
        path_val = info['environment']['PATH']
        if len(path_val) > 200:
            info['environment']['PATH'] = path_val[:200] + '... (truncated)'
    
    # WeasyPrint info
    info['weasyprint'] = get_weasyprint_info()
    
    # Library versions
    info['libraries'] = get_library_versions()
    
    return info


def get_windows_display_info() -> Dict[str, Any]:
    """Get Windows-specific display information."""
    display_info = {}
    
    try:
        import ctypes
        from ctypes import wintypes
        
        # Get screen dimensions
        user32 = ctypes.windll.user32
        display_info['screen_width_raw'] = user32.GetSystemMetrics(0)  # SM_CXSCREEN
        display_info['screen_height_raw'] = user32.GetSystemMetrics(1)  # SM_CYSCREEN
        
        # Get virtual screen (all monitors)
        display_info['virtual_screen_width'] = user32.GetSystemMetrics(78)  # SM_CXVIRTUALSCREEN
        display_info['virtual_screen_height'] = user32.GetSystemMetrics(79)  # SM_CYVIRTUALSCREEN
        
        # Get primary monitor info via GetDC
        try:
            gdi32 = ctypes.windll.gdi32
            hdc = user32.GetDC(0)
            display_info['physical_width_mm'] = gdi32.GetDeviceCaps(hdc, 4)   # HORZSIZE
            display_info['physical_height_mm'] = gdi32.GetDeviceCaps(hdc, 6)  # VERTSIZE
            display_info['pixels_per_inch_x'] = gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            display_info['pixels_per_inch_y'] = gdi32.GetDeviceCaps(hdc, 90)  # LOGPIXELSY
            display_info['device_scale_factor_x'] = gdi32.GetDeviceCaps(hdc, 10)  # HORZRES
            display_info['device_scale_factor_y'] = gdi32.GetDeviceCaps(hdc, 12)  # VERTRES
            user32.ReleaseDC(0, hdc)
        except Exception as e:
            display_info['gdi_error'] = str(e)
            
    except Exception as e:
        display_info['error'] = str(e)
    
    return display_info


def get_windows_dpi_info() -> Dict[str, Any]:
    """Get Windows DPI awareness and scaling information."""
    dpi_info = {}
    
    try:
        import ctypes
        
        # Check DPI awareness mode
        try:
            shcore = ctypes.windll.shcore
            awareness = ctypes.c_int()
            result = shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
            dpi_info['dpi_awareness_result'] = result
            dpi_info['dpi_awareness_value'] = awareness.value
            dpi_info['dpi_awareness_mode'] = {
                0: 'DPI_UNAWARE',
                1: 'SYSTEM_DPI_AWARE', 
                2: 'PER_MONITOR_DPI_AWARE'
            }.get(awareness.value, f'UNKNOWN ({awareness.value})')
        except Exception as e:
            dpi_info['dpi_awareness_error'] = str(e)
        
        # Try to get actual DPI for primary monitor
        try:
            # GetDpiForSystem requires Windows 10 1607+
            user32 = ctypes.windll.user32
            try:
                system_dpi = user32.GetDpiForSystem()
                dpi_info['system_dpi'] = system_dpi
                dpi_info['scale_factor'] = system_dpi / 96.0
                dpi_info['scale_percentage'] = f"{int(system_dpi / 96.0 * 100)}%"
            except AttributeError:
                dpi_info['system_dpi_note'] = 'GetDpiForSystem not available (requires Win10 1607+)'
        except Exception as e:
            dpi_info['system_dpi_error'] = str(e)
            
        # Try GetDpiForWindow on desktop window
        try:
            user32 = ctypes.windll.user32
            hwnd_desktop = user32.GetDesktopWindow()
            try:
                desktop_dpi = user32.GetDpiForWindow(hwnd_desktop)
                dpi_info['desktop_window_dpi'] = desktop_dpi
            except AttributeError:
                pass
        except Exception as e:
            dpi_info['desktop_dpi_error'] = str(e)
            
        # Check if we successfully set DPI unaware earlier
        # If awareness is 0, our SetProcessDpiAwareness(0) worked
        if dpi_info.get('dpi_awareness_value') == 0:
            dpi_info['dpi_override_status'] = 'SUCCESS - Process is DPI Unaware'
        elif dpi_info.get('dpi_awareness_value') == 1:
            dpi_info['dpi_override_status'] = 'PARTIAL - Process is System DPI Aware (may have issues)'
        elif dpi_info.get('dpi_awareness_value') == 2:
            dpi_info['dpi_override_status'] = 'FAILED - Process is Per-Monitor DPI Aware (likely to have issues)'
        
    except Exception as e:
        dpi_info['error'] = str(e)
    
    return dpi_info


def get_weasyprint_info() -> Dict[str, Any]:
    """Get WeasyPrint availability and configuration info."""
    wp_info = {}
    
    try:
        from weasyprint import __version__ as wp_version
        wp_info['available'] = True
        wp_info['version'] = wp_version
        
        # Try to get Cairo/Pango info
        try:
            import cairocffi
            wp_info['cairo_version'] = cairocffi.cairo_version_string()
        except Exception:
            pass
            
        try:
            import pangocffi
            wp_info['pango_available'] = True
        except ImportError:
            wp_info['pango_available'] = False
            
    except ImportError:
        wp_info['available'] = False
        wp_info['reason'] = 'WeasyPrint not installed'
    except OSError as e:
        wp_info['available'] = False
        wp_info['reason'] = f'Native library error: {e}'
    except Exception as e:
        wp_info['available'] = False
        wp_info['reason'] = str(e)
    
    return wp_info


def get_library_versions() -> Dict[str, str]:
    """Get versions of key libraries."""
    libs = {}
    
    library_names = [
        ('PIL', 'PIL'),
        ('reportlab', 'reportlab'),
        ('pypdf', 'pypdf'),
        ('tkinter', 'tkinter'),
        ('cairocffi', 'cairocffi'),
    ]
    
    for display_name, import_name in library_names:
        try:
            mod = __import__(import_name)
            version = getattr(mod, '__version__', getattr(mod, 'VERSION', 'installed'))
            libs[display_name] = str(version)
        except ImportError:
            libs[display_name] = 'not installed'
        except Exception as e:
            libs[display_name] = f'error: {e}'
    
    return libs


def format_system_info(info: Dict[str, Any], indent: int = 0) -> str:
    """Format system info dictionary as readable text."""
    lines = []
    prefix = "  " * indent
    
    for key, value in info.items():
        if isinstance(value, dict):
            lines.append(f"{prefix}{key}:")
            lines.append(format_system_info(value, indent + 1))
        else:
            lines.append(f"{prefix}{key}: {value}")
    
    return "\n".join(lines)


def generate_diagnostic_report() -> str:
    """
    Generate a full diagnostic report as a string.
    
    This can be shown to the user or written to a file.
    """
    info = get_system_info()
    
    report = []
    report.append("=" * 60)
    report.append("MAYO'S MAIL CONVERTER - DIAGNOSTIC REPORT")
    report.append("=" * 60)
    report.append("")
    
    # Platform
    report.append("PLATFORM:")
    report.append("-" * 40)
    for k, v in info['platform'].items():
        report.append(f"  {k}: {v}")
    report.append("")
    
    # Python
    report.append("PYTHON:")
    report.append("-" * 40)
    for k, v in info['python'].items():
        report.append(f"  {k}: {v}")
    report.append("")
    
    # Display (Windows)
    if info['display']:
        report.append("DISPLAY INFO:")
        report.append("-" * 40)
        for k, v in info['display'].items():
            report.append(f"  {k}: {v}")
        report.append("")
    
    # DPI (Windows)
    if info['dpi']:
        report.append("DPI/SCALING INFO:")
        report.append("-" * 40)
        for k, v in info['dpi'].items():
            report.append(f"  {k}: {v}")
        report.append("")
    
    # Environment
    report.append("ENVIRONMENT VARIABLES:")
    report.append("-" * 40)
    for k, v in info['environment'].items():
        report.append(f"  {k}: {v}")
    report.append("")
    
    # WeasyPrint
    report.append("WEASYPRINT:")
    report.append("-" * 40)
    for k, v in info['weasyprint'].items():
        report.append(f"  {k}: {v}")
    report.append("")
    
    # Libraries
    report.append("LIBRARIES:")
    report.append("-" * 40)
    for k, v in info['libraries'].items():
        report.append(f"  {k}: {v}")
    report.append("")
    
    report.append("=" * 60)
    report.append("END OF REPORT")
    report.append("=" * 60)
    
    return "\n".join(report)


def log_system_info():
    """Log system info at startup for debugging."""
    try:
        info = get_system_info()
        logger.info("System Info:")
        logger.info(f"  Platform: {info['platform']['system']} {info['platform']['release']}")
        logger.info(f"  Python: {sys.version.split()[0]}, Bundled: {info['python']['is_bundled']}")
        
        if sys.platform == 'win32' and info['dpi']:
            logger.info(f"  DPI Awareness: {info['dpi'].get('dpi_awareness_mode', 'unknown')}")
            logger.info(f"  System DPI: {info['dpi'].get('system_dpi', 'unknown')}")
            logger.info(f"  Scale Factor: {info['dpi'].get('scale_percentage', 'unknown')}")
            logger.info(f"  Override Status: {info['dpi'].get('dpi_override_status', 'unknown')}")
            
        if info['weasyprint']:
            logger.info(f"  WeasyPrint: {info['weasyprint'].get('version', 'N/A')} (available: {info['weasyprint'].get('available')})")
    except Exception as e:
        logger.warning(f"Could not collect system info: {e}")


if __name__ == '__main__':
    # When run directly, print the diagnostic report
    print(generate_diagnostic_report())
