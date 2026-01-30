# Mayo's Mail Converter

Convert PST email archives to PDF for e-discovery, litigation support, and records management.

## What It Does

- Extracts all emails from Outlook PST files
- Converts each email + all attachments into a single, searchable PDF
- Names files chronologically: `YYYYMMDD_HHMMSS_Subject.pdf`
- Creates a combined PDF with everything in date order
- Supports documents, images, nested emails, and more

**Perfect for:** Law firms, legal assistants, compliance teams, anyone dealing with discovery requests.

---

## Downloads

### Windows (Easiest)

**[Download Latest Release](https://github.com/edydex/mail_converter/releases/latest)**

1. Download `MayosMailConverter_Portable.zip`
2. Extract anywhere
3. Double-click `MayosMailConverter.exe`
4. No installation required!

**Included:** PST extraction (readpst), PDF processing (Poppler)  
**Optional:** [LibreOffice](https://www.libreoffice.org/download/) for Office docs

---

### macOS

One script does everything - installs dependencies, sets up Python environment, launches the app:

```bash
# Download and run
git clone https://github.com/edydex/mail_converter.git
cd mail_converter
./run_mac.sh
```

First run will:
- Install Homebrew (if needed)
- Install Python 3 and libpst via Homebrew
- Create a virtual environment
- Install all Python packages
- Launch the app

Subsequent runs just launch the app instantly.

---

### Linux (Ubuntu/Debian/Fedora/Arch)

Same deal - one script handles everything:

```bash
# Download and run
git clone https://github.com/edydex/mail_converter.git
cd mail_converter
./run_linux.sh
```

First run will:
- Install system packages via apt/dnf/pacman (requires sudo)
- Create a virtual environment
- Install all Python packages
- Launch the app

Supported package managers: apt, dnf, yum, pacman, zypper

---

## Output Structure

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

---

## Supported Formats

| Category | Formats |
|----------|---------|
| **Email** | PST, EML, MSG |
| **Documents** | PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, TXT, CSV, HTML |
| **Images** | JPG, PNG, GIF, BMP, TIFF |
| **Other** | Nested email attachments |

---

## Optional: Better Document Conversion

For best results converting Office documents (Word, Excel, PowerPoint):

Install [LibreOffice](https://www.libreoffice.org/download/) (free)

The app will automatically detect LibreOffice if installed. Without it, Office docs are embedded as attachments rather than converted to PDF.

---

## License

**PolyForm Noncommercial License 1.0.0**

- Free for personal use
- Free for non-profits, educational institutions, government
- Commercial use requires a separate license

See [LICENSE](LICENSE) for details.

---

## Contributing

Issues and pull requests welcome.

---

## Project Structure

```
mail_converter/
├── main.py                    # Entry point
├── run_mac.sh                 # macOS launcher/installer
├── run_linux.sh               # Linux launcher/installer
├── core/
│   ├── pst_extractor.py       # PST extraction via readpst
│   ├── email_parser.py        # Email parsing
│   ├── email_to_pdf.py        # Email to PDF conversion
│   ├── attachment_converter.py # Convert attachments to PDF
│   ├── pdf_merger.py          # Merge PDFs
│   └── conversion_pipeline.py # Main processing pipeline
├── gui/
│   ├── main_window.py         # Main application window
│   ├── progress_dialog.py     # Progress dialog
│   └── settings_dialog.py     # Settings dialog
└── assets/
    ├── icon.png               # Application icon
    └── icon.ico               # Windows icon
```

---

## PST Writing Limitation & Future Work

### The Problem

The Email Tools feature can compare, merge, deduplicate, and filter mailboxes.

### PST Output with Date Preservation ✅

**Good news!** PST writing with preserved sent/received dates is now fully supported on Windows using Extended MAPI via pywin32. No commercial libraries required.

| Output Format | Dates Preserved | Platform |
|---------------|-----------------|----------|
| **EML Folder** | ✅ Yes | All |
| **MBOX** | ✅ Yes | All |
| **PST** | ✅ Yes | Windows (with Outlook) |

**Requirements for PST output:**
- Windows with Microsoft Outlook installed
- Outlook must be running when converting
- pywin32 package (included in requirements.txt)

The converter automatically uses Extended MAPI when available, which properly sets `PR_MESSAGE_DELIVERY_TIME` and `PR_CLIENT_SUBMIT_TIME` properties on messages.


