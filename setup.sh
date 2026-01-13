#!/bin/bash
#
# Mail Converter Setup Script
# Installs dependencies for macOS, Linux, or shows instructions for Windows
#

set -e

echo "==========================================="
echo "   Mayo's Mail Converter - Setup Script"
echo "==========================================="
echo ""

# Detect OS
OS="$(uname -s)"

install_macos() {
    echo "Detected: macOS"
    echo ""
    
    # Check for Homebrew
    if ! command -v brew &> /dev/null; then
        echo "Homebrew not found. Installing..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    fi
    
    echo "Installing system dependencies..."
    brew install libpst tesseract poppler
    
    # wkhtmltopdf
    if ! command -v wkhtmltopdf &> /dev/null; then
        echo "Installing wkhtmltopdf..."
        brew install --cask wkhtmltopdf
    fi
    
    # LibreOffice (optional but recommended)
    if ! command -v libreoffice &> /dev/null; then
        read -p "Install LibreOffice for better DOC/PPT support? (y/n) " -n 1 -r
        echo
        if [[ $REPLY =~ ^[Yy]$ ]]; then
            brew install --cask libreoffice
        fi
    fi
    
    echo ""
    echo "System dependencies installed!"
}

install_linux() {
    echo "Detected: Linux"
    echo ""
    
    if command -v apt-get &> /dev/null; then
        echo "Using apt-get..."
        sudo apt-get update
        sudo apt-get install -y libpst-dev pst-utils tesseract-ocr poppler-utils wkhtmltopdf libreoffice
    elif command -v dnf &> /dev/null; then
        echo "Using dnf..."
        sudo dnf install -y libpst tesseract poppler-utils wkhtmltopdf libreoffice
    elif command -v pacman &> /dev/null; then
        echo "Using pacman..."
        sudo pacman -S libpst tesseract poppler wkhtmltopdf libreoffice-fresh
    else
        echo "Could not detect package manager. Please install manually:"
        echo "  - libpst (pst-utils)"
        echo "  - tesseract"
        echo "  - poppler-utils"
        echo "  - wkhtmltopdf"
        echo "  - libreoffice"
    fi
    
    echo ""
    echo "System dependencies installed!"
}

show_windows_instructions() {
    echo "Detected: Windows (or running in Git Bash/WSL)"
    echo ""
    echo "Please install the following manually:"
    echo ""
    echo "1. Tesseract OCR:"
    echo "   https://github.com/UB-Mannheim/tesseract/wiki"
    echo ""
    echo "2. Poppler (for PDF processing):"
    echo "   https://github.com/oschwartz10612/poppler-windows/releases"
    echo ""
    echo "3. wkhtmltopdf (for HTML to PDF):"
    echo "   https://wkhtmltopdf.org/downloads.html"
    echo ""
    echo "4. LibreOffice (for Office document conversion):"
    echo "   https://www.libreoffice.org/download/"
    echo ""
    echo "5. libpst (for PST extraction):"
    echo "   Either use WSL or download from:"
    echo "   https://www.five-ten-sg.com/libpst/"
    echo ""
    echo "After installing, add them to your PATH."
}

# Install based on OS
case "$OS" in
    Darwin)
        install_macos
        ;;
    Linux)
        install_linux
        ;;
    MINGW*|MSYS*|CYGWIN*)
        show_windows_instructions
        ;;
    *)
        echo "Unknown OS: $OS"
        echo "Please install dependencies manually."
        ;;
esac

echo ""
echo "==========================================="
echo "   Setting up Python environment"
echo "==========================================="
echo ""

# Check for Python
if ! command -v python3 &> /dev/null; then
    echo "Python 3 not found. Please install Python 3.8 or later."
    exit 1
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install requirements
echo "Installing Python dependencies..."
pip install -r requirements.txt

echo ""
echo "==========================================="
echo "   Setup Complete!"
echo "==========================================="
echo ""
echo "To run the application:"
echo ""
echo "  source venv/bin/activate"
echo "  python main.py"
echo ""
echo "To build a Windows executable:"
echo ""
echo "  source venv/bin/activate"
echo "  pip install pyinstaller"
echo "  pyinstaller mail_converter.spec"
echo ""
