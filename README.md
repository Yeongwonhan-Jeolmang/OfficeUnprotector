# OfficeUnprotector

A command-line tool to remove password protection from PDF and Office files.

## Requirements

Python 3.10+ and the following packages:

```bash
pip install pypdf msoffcrypto-tool lxml
```

## Usage

```bash
python unprotect.py <file> [password] [options]
python unprotect.py <glob> [password] [options]
```

| Argument | Required | Description |
|---|---|---|
| `file` / `glob` | Yes | One or more file paths or glob patterns |
| `password` | No | Password to open the file (omit if only edit/sheet protection is set) |
| `--output` / `-o` | No | Output path — single-file mode only (default: `unlocked_<filename>`) |
| `--in-place` | No | Overwrite the original file instead of writing a new one |
| `--check` | No | Report protection status without modifying any files |

`--output` and `--in-place` are mutually exclusive. The original file is never modified unless `--in-place` is set.

## Examples

```bash
# Unprotect a PDF
python unprotect.py document.pdf mypassword

# Unprotect a Word document
python unprotect.py report.docx mypassword

# Unprotect an Excel file
python unprotect.py spreadsheet.xlsx mypassword

# Unprotect a PowerPoint presentation
python unprotect.py slides.pptx mypassword

# Custom output path
python unprotect.py document.pdf mypassword --output clean.pdf

# File with only sheet/edit protection (no open password needed)
python unprotect.py spreadsheet.xlsx

# Check protection status without modifying anything
python unprotect.py document.pdf --check

# Overwrite the original file in place
python unprotect.py report.docx mypassword --in-place

# Process multiple files at once
python unprotect.py file1.xlsx file2.xlsx mypassword

# Process all Excel files in a directory
python unprotect.py "*.xlsx" mypassword
```

## Supported Formats

| Format | Extensions | What gets removed |
|---|---|---|
| PDF | `.pdf` | Open password / encryption |
| Word | `.docx` | Open password + `documentProtection` / `writeProtection` |
| Excel | `.xlsx` `.xlsm` | Open password + sheet and workbook protection |
| PowerPoint | `.pptx` | Open password + `modifyVerifier` / `writeProtection` + locked OLE objects per-slide |

> **Legacy formats** (`.doc`, `.xls`, `.ppt`) are not supported. Open the file in LibreOffice or Microsoft Office, save it as the modern format, then retry.

## How it works

Modern Office files (`.docx`, `.xlsx`, `.pptx`) are ZIP archives containing XML. Protection sits on two independent layers:

**Layer 1 — file encryption** (password required to open): handled by `msoffcrypto-tool`, which decrypts the file into a temporary copy before any further processing.

**Layer 2 — edit/structure protection** (file opens fine, but editing is locked): handled by direct XML manipulation via `lxml`. The relevant protection elements are removed from the XML inside the ZIP, rewriting only those entries while preserving compression and metadata for everything else.

For PDFs, `pypdf` handles both decryption and reconstruction.

## Troubleshooting

**`ModuleNotFoundError`** — run `pip install pypdf msoffcrypto-tool lxml`

**`Error: Wrong password`** — double-check the password; the original file is left untouched

**`Unsupported file type`** — only `.pdf`, `.docx`, `.xlsx`, `.xlsm`, `.pptx` are supported; convert legacy formats first

**`--output` rejected with multiple files** — `--output` only works for a single file; use `--in-place` or omit it to get the default `unlocked_<filename>` naming for each file