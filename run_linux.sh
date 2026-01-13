#!/bin/bash
# Mayo's Mail Converter - Linux Launcher/Installer
# This script will install all dependencies if needed, then launch the app.
# 
# Usage: 
#   chmod +x run_linux.sh
#   ./run_linux.sh
#
# Note: First run may require sudo for system packages (libpst, tkinter)

set -e  # Exit on error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/venv"
REQUIREMENTS_FILE="$SCRIPT_DIR/requirements.txt"
MAIN_SCRIPT="$SCRIPT_DIR/main.py"

echo -e "${BLUE}"
echo "=============================================="
echo "   Mayo's Mail Converter - Linux"
echo "=============================================="
echo -e "${NC}"

# Function to check if a command exists
command_exists() {
    command -v "$1" &> /dev/null
}

# Function to print status
print_status() {
    echo -e "${GREEN}[✓]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[!]${NC} $1"
}

print_error() {
    echo -e "${RED}[✗]${NC} $1"
}

print_info() {
    echo -e "${BLUE}[i]${NC} $1"
}

# Detect package manager
detect_package_manager() {
    if command_exists apt-get; then
        echo "apt"
    elif command_exists dnf; then
        echo "dnf"
    elif command_exists yum; then
        echo "yum"
    elif command_exists pacman; then
        echo "pacman"
    elif command_exists zypper; then
        echo "zypper"
    else
        echo "unknown"
    fi
}

PKG_MANAGER=$(detect_package_manager)

# Track if we need to install anything
NEEDS_SYSTEM_INSTALL=false
NEEDS_PYTHON_SETUP=false
SYSTEM_PACKAGES_TO_INSTALL=()

# =============================================================================
# CHECK SYSTEM DEPENDENCIES
# =============================================================================

echo ""
echo "Checking system dependencies..."
echo ""

# Check for Python 3
if ! command_exists python3; then
    print_warning "Python 3 not found - will install"
    NEEDS_SYSTEM_INSTALL=true
    case $PKG_MANAGER in
        apt) SYSTEM_PACKAGES_TO_INSTALL+=("python3" "python3-venv" "python3-pip") ;;
        dnf|yum) SYSTEM_PACKAGES_TO_INSTALL+=("python3" "python3-pip") ;;
        pacman) SYSTEM_PACKAGES_TO_INSTALL+=("python" "python-pip") ;;
        zypper) SYSTEM_PACKAGES_TO_INSTALL+=("python3" "python3-pip") ;;
    esac
else
    PYTHON_VERSION=$(python3 --version 2>&1 | cut -d' ' -f2)
    print_status "Python $PYTHON_VERSION installed"
fi

# Check for pip
if ! python3 -m pip --version &>/dev/null; then
    print_warning "pip not found - will install"
    NEEDS_SYSTEM_INSTALL=true
    case $PKG_MANAGER in
        apt) SYSTEM_PACKAGES_TO_INSTALL+=("python3-pip") ;;
        dnf|yum) SYSTEM_PACKAGES_TO_INSTALL+=("python3-pip") ;;
        pacman) SYSTEM_PACKAGES_TO_INSTALL+=("python-pip") ;;
        zypper) SYSTEM_PACKAGES_TO_INSTALL+=("python3-pip") ;;
    esac
else
    print_status "pip installed"
fi

# Check for venv module
if ! python3 -m venv --help &>/dev/null; then
    print_warning "python3-venv not found - will install"
    NEEDS_SYSTEM_INSTALL=true
    case $PKG_MANAGER in
        apt) SYSTEM_PACKAGES_TO_INSTALL+=("python3-venv") ;;
        dnf|yum) ;; # Usually included with python3
        pacman) ;; # Usually included with python
        zypper) SYSTEM_PACKAGES_TO_INSTALL+=("python3-venv") ;;
    esac
else
    print_status "python3-venv available"
fi

# Check for tkinter
if ! python3 -c "import tkinter" &>/dev/null; then
    print_warning "tkinter not found - will install"
    NEEDS_SYSTEM_INSTALL=true
    case $PKG_MANAGER in
        apt) SYSTEM_PACKAGES_TO_INSTALL+=("python3-tk") ;;
        dnf|yum) SYSTEM_PACKAGES_TO_INSTALL+=("python3-tkinter") ;;
        pacman) SYSTEM_PACKAGES_TO_INSTALL+=("tk") ;;
        zypper) SYSTEM_PACKAGES_TO_INSTALL+=("python3-tk") ;;
    esac
else
    print_status "tkinter available"
fi

# Check for libpst (readpst)
if ! command_exists readpst; then
    print_warning "libpst (readpst) not found - will install"
    NEEDS_SYSTEM_INSTALL=true
    case $PKG_MANAGER in
        apt) SYSTEM_PACKAGES_TO_INSTALL+=("pst-utils") ;;
        dnf|yum) SYSTEM_PACKAGES_TO_INSTALL+=("libpst") ;;
        pacman) SYSTEM_PACKAGES_TO_INSTALL+=("libpst") ;;
        zypper) SYSTEM_PACKAGES_TO_INSTALL+=("libpst") ;;
    esac
else
    print_status "libpst (readpst) installed"
fi

# Check for virtual environment
if [ ! -d "$VENV_DIR" ]; then
    print_warning "Virtual environment not found - will create"
    NEEDS_PYTHON_SETUP=true
else
    print_status "Virtual environment exists"
fi

# =============================================================================
# INSTALL SYSTEM DEPENDENCIES (requires sudo)
# =============================================================================

if [ "$NEEDS_SYSTEM_INSTALL" = true ] && [ ${#SYSTEM_PACKAGES_TO_INSTALL[@]} -gt 0 ]; then
    echo ""
    echo -e "${YELLOW}System packages need to be installed:${NC}"
    echo "  ${SYSTEM_PACKAGES_TO_INSTALL[*]}"
    echo ""
    echo "This requires sudo access."
    echo ""
    read -p "Continue with installation? (Y/n): " -n 1 -r
    echo ""
    
    if [[ ! $REPLY =~ ^[Nn]$ ]]; then
        # Remove duplicates from package list
        UNIQUE_PACKAGES=($(echo "${SYSTEM_PACKAGES_TO_INSTALL[@]}" | tr ' ' '\n' | sort -u | tr '\n' ' '))
        
        print_info "Installing system packages..."
        
        case $PKG_MANAGER in
            apt)
                sudo apt-get update
                sudo apt-get install -y "${UNIQUE_PACKAGES[@]}"
                ;;
            dnf)
                sudo dnf install -y "${UNIQUE_PACKAGES[@]}"
                ;;
            yum)
                sudo yum install -y "${UNIQUE_PACKAGES[@]}"
                ;;
            pacman)
                sudo pacman -S --noconfirm "${UNIQUE_PACKAGES[@]}"
                ;;
            zypper)
                sudo zypper install -y "${UNIQUE_PACKAGES[@]}"
                ;;
            *)
                print_error "Unknown package manager. Please install manually:"
                echo "  - python3, python3-pip, python3-venv"
                echo "  - python3-tk (or python3-tkinter)"
                echo "  - libpst (or pst-utils)"
                exit 1
                ;;
        esac
        
        print_status "System packages installed"
    else
        print_error "Installation cancelled. Cannot run without dependencies."
        exit 1
    fi
fi

# =============================================================================
# SETUP PYTHON VIRTUAL ENVIRONMENT
# =============================================================================

if [ "$NEEDS_PYTHON_SETUP" = true ] || [ ! -f "$VENV_DIR/bin/activate" ]; then
    echo ""
    print_info "Setting up Python virtual environment..."
    
    # Create virtual environment
    python3 -m venv "$VENV_DIR"
    print_status "Virtual environment created"
    
    NEEDS_PYTHON_SETUP=true
fi

# Activate virtual environment
source "$VENV_DIR/bin/activate"
print_status "Virtual environment activated"

# Install/update Python packages if needed
if [ "$NEEDS_PYTHON_SETUP" = true ] || [ "$REQUIREMENTS_FILE" -nt "$VENV_DIR/.requirements_installed" ]; then
    echo ""
    print_info "Installing Python packages..."
    
    # Upgrade pip
    pip install --upgrade pip --quiet
    
    # Install requirements
    if [ -f "$REQUIREMENTS_FILE" ]; then
        pip install -r "$REQUIREMENTS_FILE" --quiet
        touch "$VENV_DIR/.requirements_installed"
        print_status "Python packages installed"
    else
        print_warning "requirements.txt not found - some features may not work"
    fi
fi

# =============================================================================
# VERIFY INSTALLATION
# =============================================================================

echo ""
echo "Verifying installation..."

# Check critical imports
if python3 -c "import tkinter" 2>/dev/null; then
    print_status "tkinter available"
else
    print_error "tkinter not available - GUI won't work"
fi

if python3 -c "from weasyprint import HTML" 2>/dev/null; then
    print_status "WeasyPrint available"
else
    print_warning "WeasyPrint not available - using fallback PDF generation"
    print_info "For better PDF output, install WeasyPrint system dependencies:"
    case $PKG_MANAGER in
        apt) echo "  sudo apt-get install libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0" ;;
        dnf|yum) echo "  sudo dnf install pango gdk-pixbuf2" ;;
        pacman) echo "  sudo pacman -S pango gdk-pixbuf2" ;;
    esac
fi

if python3 -c "import pikepdf" 2>/dev/null; then
    print_status "pikepdf available"
else
    print_error "pikepdf not available - PDF merging won't work"
fi

# =============================================================================
# LAUNCH APPLICATION
# =============================================================================

echo ""
echo -e "${GREEN}=============================================="
echo "   Launching Mayo's Mail Converter..."
echo "==============================================${NC}"
echo ""

# Run the application
cd "$SCRIPT_DIR"
python3 "$MAIN_SCRIPT"

# Deactivate on exit
deactivate 2>/dev/null || true
