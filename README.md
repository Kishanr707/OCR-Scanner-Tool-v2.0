# OCR Scanner Tool v2.0

A desktop application that scans visiting cards and contact documents, extracts name, phone number, and email using AI vision, and saves everything into a persistent master Excel sheet.

Built for internal/industrial use — functional, fast, and offline-friendly once packaged.

---

## What It Does

- Accepts image files, PDFs, and Word documents as input
- Uses AI vision to extract **name**, **phone number**, and **email** from each file
- Detects and skips unreadable or irrelevant files with clear error messages
- Checks for duplicate entries before saving
- Appends all results to a single persistent `contacts.xlsx` file
- Native desktop GUI — no browser, no localhost, no dependencies for the end user

---

## Supported File Types

| Type | Extensions |
|---|---|
| Images | `.jpg` `.jpeg` `.png` `.tiff` `.bmp` `.webp` |
| PDF (text-based) | `.pdf` |
| Word Document | `.docx` |

> **Note:** Scanned PDFs with no text layer are not supported. Export the card as a JPG or PNG instead.

---

## Project Structure

```
OCR-Scanner-Tool-v2.0/
│
├── main.py              ← GUI entry point
├── scanner.py           ← AI API calls and extraction logic
├── excel_manager.py     ← contacts.xlsx read/write
├── config.py            ← API key and app constants
├── build.spec           ← PyInstaller configuration
├── build.bat            ← One-click EXE builder
└── README.md
```

---

## Setup & Installation

### Requirements

- Python 3.12+
- A valid Gemini API key

### Install Dependencies

```bash
pip install google-genai pdfplumber python-docx openpyxl pillow pyinstaller
```

### Add Your API Key

Open `config.py` and replace the placeholder:

```python
GEMINI_API_KEY = "YOUR_API_KEY_HERE"
```

Get your key from: https://aistudio.google.com/app/apikey

### Run Directly

```bash
python main.py
```

---

## Building the EXE

1. Make sure your API key is set in `config.py`
2. Double-click `build.bat`
3. Wait for the build to complete
4. Output will be at:

```
dist\VisitingCardScanner\
```

---

## Running the EXE

Navigate to:

```
dist\VisitingCardScanner\
```

Open `VisitingCardScanner.exe`

The app launches as a standalone desktop window. No Python, no browser, no installation needed on the target machine.

> `contacts.xlsx` will be created automatically in the same folder as the EXE on first scan.

---

## How to Use

1. Click **+ Add File** to add up to 10 files
2. Either paste a full file path or click **BROWSE** to select files manually
3. Click **GET DETAILS** to start scanning
4. Results appear per file:
   - **✓ Green** — successfully extracted all fields
   - **⚠ Yellow** — partial data, one or more fields missing
   - **✕ Red** — error, file skipped with reason shown
5. Click **OPEN CONTACTS SHEET** to open `contacts.xlsx` directly in Excel

---

## Output — contacts.xlsx

| # | Name | Phone | Email | Source File | Scanned At |
|---|---|---|---|---|---|
| 1 | Olivia Wilson | +123-456-7890 | hello@example.com | card.jpg | 2026-03-21 10:30 |

- Created automatically on first scan
- Every scan appends new rows — existing data is never overwritten
- Duplicate emails are detected and skipped automatically

---

## Notes

- Accuracy is highest on clear, well-lit images of visiting cards
- For DOCX files, contact info should be in the first section of the document
- The app truncates long documents before sending to the API to keep costs minimal
- All processing happens via API call — internet connection required during scan
