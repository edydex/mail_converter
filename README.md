# Mayo's Mail Converter

Convert PST email archives to PDF for e-discovery, litigation support, and records management.

## 🎯 What It Does

- Extracts all emails from Outlook PST files
- Converts each email + all attachments into a single, searchable PDF
- Names files chronologically: `YYYYMMDD_HHMMSS_Subject.pdf`
- Creates a combined PDF with everything in date order
- Supports documents, images, nested emails, and more

**Perfect for:** Law firms, legal assistants, compliance teams, anyone dealing with discovery requests.

## 📦 Downloads

### Windows (Recommended)
**[Download Latest Release](https://github.com/edydex/mail_converter/releases/latest)**

1. Download `MailConverter_vX.X.X_Windows.zip`
2. Extract anywhere
3. Double-click `MailConverter.exe`
4. No installation required!

**Bundled:** PST extraction, PDF processing  
**Optional:** [LibreOffice](https://www.libreoffice.org/download/) for Office docs, [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) for OCR

---

### macOS / Linux (Power Users)

<details>
<summary>Click to expand installation instructions</summary>

#### macOS
```bash
# Install dependencies
brew install libpst tesseract poppler libreoffice

# Clone and setup
git clone https://github.com/edydex/mail_converter.git
cd mail_converter
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Run
python main.py
```

#### Linux (Ubuntu/Debian)
```bash
# Install dependencies
sudo apt-get update
sudo apt-get install pst-utils tesseract-ocr poppler-utils libreoffice

# Clone and setup
git clone https://github.com/edydex/mail_converter.git
cd mail_converter
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Run
python main.py
```

</details>

---

## 📁 Output Structure

```
output/
├── 1_extracted_emls/              # Raw email files
├── 2_individual_pdfs/             # Each email as its own PDF
│   ├── 20240115_093045_Meeting_Notes.pdf
│   ├── 20240115_143022_Contract_Review.pdf
│   └── ...
├── 3_combined_output/
│   └── combined_chronological.pdf  # Everything merged, in date order
└── conversion_log.txt
```

## ✅ Supported Formats

| Category | Formats |
|----------|---------|
| **Email** | PST, EML, MSG |
| **Documents** | PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, TXT, CSV, HTML |
| **Images** | JPG, PNG, GIF, BMP, TIFF |
| **Other** | Nested email attachments, scanned PDFs (with OCR) |

## 🔧 Optional: Better Document Conversion

For best results converting Office documents (Word, Excel, PowerPoint):

**Windows:** Install [LibreOffice](https://www.libreoffice.org/download/) (free, portable version works)

The app will automatically detect LibreOffice if installed. Without it, Office docs are embedded as attachments rather than converted.

## 📋 License

**PolyForm Noncommercial License 1.0.0**

- ✅ Free for personal use
- ✅ Free for non-profits, educational institutions, government
- ❌ Commercial use requires a separate license

See [LICENSE](LICENSE) for details.

## 🤝 Contributing

Issues and pull requests welcome! This is a side project, so response times may vary.
  - `attachment_converter.py` - Convert various formats to PDF
  - `pdf_merger.py` - Merge PDFs and handle OCR
  - `duplicate_detector.py` - Duplicate email detection
- `utils/`
  - `ocr_handler.py` - OCR processing wrapper
  - `file_utils.py` - File handling utilities
  - `progress_tracker.py` - Progress tracking for GUI
- `gui/`
  - `main_window.py` - Main application window
  - `progress_dialog.py` - Progress dialog
  - `settings_dialog.py` - Settings/options dialog

## License

MIT License
