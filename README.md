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

The Email Tools feature can compare, merge, deduplicate, and filter mailboxes. However, **writing to PST format with preserved sent/received dates** is not currently possible in an open-source way.

Microsoft's PST format is proprietary. While reading PST files is well-supported (via `libpst`/`readpst`), **writing** PST files with correct dates requires using Microsoft's Extended MAPI API at a low level.

| Output Format | Dates Preserved | Platform |
|---------------|-----------------|----------|
| **EML Folder** | ✅ Yes | All |
| **MBOX** | ✅ Yes | All |
| **PST** | ❌ No (shows import date) | Windows only |

**Current recommendation:** Use EML Folder or MBOX output. These formats preserve all email data including dates perfectly.

---

### For Closed-Source / Commercial Projects: Redemption Library

If you're building a **closed-source commercial product**, the [Redemption library](https://www.dimastr.com/redemption/) ($299.99 one-time) solves this:

- Provides Extended MAPI access via COM
- Allows setting `SentOn` and `ReceivedTime` on messages
- Works with Python via `win32com.client`
- **License**: One-time purchase, royalty-free distribution, unlimited end users

**Note:** Redemption's distributable license explicitly **cannot be used in open-source projects**.

The code in `mailbox_writer.py` already detects Redemption if installed and uses it automatically.

---

### For Open-Source: What Would Need To Be Built

To create an open-source solution for PST writing with date preservation, someone would need to build a Python extension that wraps Extended MAPI directly. Here's what's involved:

#### Required MAPI Interfaces to Wrap

```c
// Session management
MAPIInitialize()
MAPILogonEx()          // Returns IMAPISession

// Store access
IMAPISession::OpenMsgStore()   // Returns IMsgStore
IMsgStore::OpenEntry()         // Returns IMAPIFolder

// Message creation
IMAPIFolder::CreateMessage()   // Returns IMessage

// Property setting (THE KEY PART)
IMessage::SetProps()           // Set PR_MESSAGE_DELIVERY_TIME, PR_CLIENT_SUBMIT_TIME
IMessage::SaveChanges()        // Must set props BEFORE first save!
```

#### Key MAPI Properties for Dates

| Property | Tag | Description |
|----------|-----|-------------|
| `PR_MESSAGE_DELIVERY_TIME` | `0x0E060040` | ReceivedTime |
| `PR_CLIENT_SUBMIT_TIME` | `0x00390040` | SentOn |
| `PR_MESSAGE_FLAGS` | `0x0E070003` | Must clear `MSGFLAG_UNSENT` |

#### Technical Challenges

1. **COM Interface Binding**: MAPI uses `IUnknown`-based COM, not `IDispatch` (automation). Can't use `win32com.client` directly.

2. **Structure Marshaling**: Need to handle `SPropValue`, `ENTRYID`, `SRowSet`, `FILETIME` structures.

3. **Bitness Matching**: Must match Outlook's bitness (32 or 64-bit).

4. **Session Management**: MAPI sessions are tricky - initialization, profile selection, cleanup.

5. **PyWin32's MAPI Module**: `win32com.mapi` exists but is incomplete. Doesn't expose `SetProps()` on messages.

#### Estimated Effort

- **6-12 months** for a working implementation
- Requires C/C++ and Python extension development experience
- Deep knowledge of Windows COM and MAPI internals
- Ongoing maintenance as Outlook versions change

#### Reference Resources

- **MFCMAPI**: Microsoft's official MAPI sample app - [github.com/microsoft/mfcmapi](https://github.com/microsoft/mfcmapi)
- **libpff**: Open-source PST reader (read-only) - [github.com/libyal/libpff](https://github.com/libyal/libpff)
- **MSDN MAPI Reference**: [docs.microsoft.com/en-us/office/client-developer/outlook/mapi](https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/)

#### Contribution Welcome

If you have Extended MAPI experience and want to tackle this, contributions are very welcome! The goal would be a minimal Python extension (`mapi_writer.pyd`) that exposes just enough to:

1. Open/create a PST store
2. Create a message in a folder
3. Set date properties before first save
4. Save the message

Even a proof-of-concept would be valuable.


