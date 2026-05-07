# OfficeUnprotector

<p align="center">
  <a href="https://github.com/Yeongwonhan-Jeolmang/OfficeUnprotector/actions/workflows/ci.yml">
    <img src="https://github.com/Yeongwonhan-Jeolmang/OfficeUnprotector/actions/workflows/ci.yml/badge.svg" alt="CI">
  </a>
  <a href="https://www.python.org/downloads/">
    <img src="https://img.shields.io/badge/python-3.10%2B-blue?logo=python&logoColor=white" alt="Python 3.10+">
  </a>
  <a href="https://github.com/Yeongwonhan-Jeolmang/OfficeUnprotector/blob/main/LICENSE">
    <img src="https://img.shields.io/github/license/Yeongwonhan-Jeolmang/OfficeUnprotector" alt="License: MIT">
  </a>
  <a href="https://github.com/Yeongwonhan-Jeolmang/OfficeUnprotector/issues">
    <img src="https://img.shields.io/github/issues/Yeongwonhan-Jeolmang/OfficeUnprotector" alt="Open Issues">
  </a>
  <a href="https://github.com/Yeongwonhan-Jeolmang/OfficeUnprotector/commits/main">
    <img src="https://img.shields.io/github/last-commit/Yeongwonhan-Jeolmang/OfficeUnprotector" alt="Last Commit">
  </a>
</p>

<p align="center">
  A command-line tool to remove password protection and editing locks from PDF and Office files.
</p>

---

## Supported Formats

| Format | Extensions | What gets removed |
|---|---|---|
| PDF | `.pdf` | Open password / encryption |
| Word | `.docx` | Open password + `documentProtection` / `writeProtection` |
| Excel | `.xlsx` `.xlsm` | Open password + sheet and workbook protection |
| PowerPoint | `.pptx` | Open password + `modifyVerifier` / `writeProtection` + locked OLE objects per-slide |

> **Legacy formats** (`.doc`, `.xls`, `.ppt`) are not supported. Open the file in LibreOffice or Microsoft Office, save it as the modern format, then retry.

---

## Requirements

Python 3.10+ and the following packages:

```bash
pip install pypdf msoffcrypto-tool lxml
```

---

## Usage

```bash
python unprotect.py <file> [options]
python unprotect.py <glob> [options]
```

| Argument | Required | Description |
|---|---|---|
| `file` / `glob` | Yes | One or more file paths or glob patterns |
| `--password` / `-p` | No | Password to open the file (omit if only edit/sheet protection is set) |
| `--output` / `-o` | No | Output path — single-file mode only (default: `unlocked_<filename>`) |
| `--in-place` | No | Overwrite the original file instead of writing a new one |
| `--check` | No | Report protection status without modifying any files |

`--output` and `--in-place` are mutually exclusive. The original file is never modified unless `--in-place` is set.

---

## Examples

```bash
# Check protection status without modifying anything
python unprotect.py document.pdf --check
python unprotect.py "*.xlsx" --check

# Unprotect a PDF (open password)
python unprotect.py document.pdf -p mypassword

# Unprotect a Word document
python unprotect.py report.docx -p mypassword

# Unprotect an Excel file
python unprotect.py spreadsheet.xlsx -p mypassword

# Unprotect a PowerPoint presentation
python unprotect.py slides.pptx -p mypassword

# File with only sheet/edit protection (no open password needed)
python unprotect.py spreadsheet.xlsx

# Custom output path
python unprotect.py document.pdf -p mypassword --output clean.pdf

# Overwrite the original file in place
python unprotect.py report.docx -p mypassword --in-place

# Process multiple files at once
python unprotect.py file1.xlsx file2.xlsx -p mypassword

# Process all Excel files in a directory
python unprotect.py "*.xlsx" -p mypassword
```

---

## How It Works

Modern Office files (`.docx`, `.xlsx`, `.pptx`) are ZIP archives containing XML. Protection sits on two independent layers:

**Layer 1 — file encryption** (password required to open): handled by `msoffcrypto-tool`, which decrypts the file into a temporary copy before any further processing.

**Layer 2 — edit/structure protection** (file opens fine, but editing is locked): handled by direct XML manipulation via `lxml`. The relevant protection elements are removed from the XML inside the ZIP, rewriting only those entries while preserving compression and metadata for everything else.

For PDFs, `pypdf` handles both decryption and reconstruction. PDFs with only an owner password (print/edit restrictions, but no open password) are unlocked automatically without needing `--password`.

---

## Troubleshooting

**`ModuleNotFoundError`** — run `pip install pypdf msoffcrypto-tool lxml`

**`Error: Wrong password`** — double-check the password; the original file is left untouched

**`Error: PDF is encrypted and requires a password`** — pass the open password with `--password` / `-p`

**`Unsupported file type`** — only `.pdf`, `.docx`, `.xlsx`, `.xlsm`, `.pptx` are supported; convert legacy formats first

**`--output` rejected with multiple files** — `--output` only works for a single file; use `--in-place` or omit it to get the default `unlocked_<filename>` naming for each file

---

## License

MIT — see [LICENSE](LICENSE) for details.