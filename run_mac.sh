#!/bin/bash
# Mayo's Mail Converter - macOS Launcher/Installer
# This script will install all dependencies if needed, then launch the app.
# 
# Usage: 
#   chmod +x run_mac.sh
#   ./run_mac.sh

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
echo "   Mayo's Mail Converter - macOS"
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

# Check if running on macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    print_error "This script is for macOS. Use run_linux.sh for Linux."
    exit 1
fi

# Track if we need to install anything
NEEDS_SYSTEM_INSTALL=false
NEEDS_PYTHON_SETUP=false

# =============================================================================
# CHECK SYSTEM DEPENDENCIES
# =============================================================================

echo ""
echo "Checking system dependencies..."
echo ""

# Check for Homebrew
if ! command_exists brew; then
    print_warning "Homebrew not found - will install"
    NEEDS_SYSTEM_INSTALL=true
else
    print_status "Homebrew installed"
fi

# Check for Python 3
if ! command_exists python3; then
    print_warning "Python 3 not found - will install"
    NEEDS_SYSTEM_INSTALL=true
else
    PYTHON_VERSION=$(python3 --version 2>&1 | cut -d' ' -f2)
    print_status "Python $PYTHON_VERSION installed"
fi

# Check for libpst (readpst)
if ! command_exists readpst; then
    print_warning "libpst (readpst) not found - will install"
    NEEDS_SYSTEM_INSTALL=true
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
# INSTALL SYSTEM DEPENDENCIES (requires user interaction for Homebrew)
# =============================================================================

if [ "$NEEDS_SYSTEM_INSTALL" = true ]; then
    echo ""
    echo -e "${YELLOW}System packages need to be installed.${NC}"
    echo "This may require your password for Homebrew installation."
    echo ""
    read -p "Continue with installation? (Y/n): " -n 1 -r
    echo ""
    
    if [[ ! $REPLY =~ ^[Nn]$ ]]; then
        # Install Homebrew if needed
        if ! command_exists brew; then
            print_info "Installing Homebrew..."
            /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
            
            # Add Homebrew to PATH for Apple Silicon
            if [[ $(uname -m) == "arm64" ]]; then
                eval "$(/opt/homebrew/bin/brew shellenv)"
                # Add to shell profile if not already there
                if ! grep -q 'brew shellenv' ~/.zprofile 2>/dev/null; then
                    echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
                fi
            fi
            print_status "Homebrew installed"
        fi
        
        # Install Python if needed
        if ! command_exists python3; then
            print_info "Installing Python 3..."
            brew install python@3.11
            print_status "Python installed"
        fi
        
        # Install libpst if needed
        if ! command_exists readpst; then
            print_info "Installing libpst (for PST file extraction)..."
            brew install libpst
            print_status "libpst installed"
        fi
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
    print_info "Try: brew install python-tk@3.11"
fi

if python3 -c "from weasyprint import HTML" 2>/dev/null; then
    print_status "WeasyPrint available"
else
    print_warning "WeasyPrint not available - using fallback PDF generation"
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

# Deactivate on exit (this won't run if app crashes, but that's fine)
deactivate 2>/dev/null || true
