# OfficeUnprotector
A simple command-line tool to remove password protection from PDF and Office 365 files.

## Requirements

Python 3.7+ and the following packages:

```bash
pip install pypdf msoffcrypto-tool openpyxl lxml
```

## Usage

```bash
python unprotect.py <file> [password] [--output <output_file>]
```

### Arguments

| Argument | Required | Description |
|---|---|---|
| `file` | Yes | Path to the protected file |
| `password` | No | Password to unlock the file (omit if no open password) |
| `--output` / `-o` | No | Custom output path (default: `unlocked_<filename>`) |

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
```

## Supported Formats

| Format | Extensions | What it removes |
|---|---|---|
| PDF | `.pdf` | Open password / encryption |
| Word | `.docx` `.doc` | Open password + document & write protection |
| Excel | `.xlsx` `.xlsm` `.xls` | Open password + sheet & workbook protection |
| PowerPoint | `.pptx` `.ppt` | Open password + modify & write protection |

## How it works

Each Office file type has two layers of protection the tool handles:

**Layer 1 — File encryption** (password to open): handled by `msoffcrypto-tool`, which decrypts the file before any further processing.

**Layer 2 — Edit/structure protection** (file opens but editing is locked):
- **Excel** — removes sheet and workbook protection via `openpyxl`
- **Word** — strips `documentProtection` and `writeProtection` tags from `settings.xml` inside the file
- **PowerPoint** — strips `modifyVerifier` and `writeProtection` tags from `presentation.xml` inside the file

## Notes

- The original file is **never modified**. Output is always written to a new file.
- If a file has no open password but has edit/sheet protection, you can omit the password argument.
- Wrong passwords will exit with an error and leave the original file untouched.
- `.doc` and `.ppt` (legacy binary formats) support file-level decryption but may have limited edit-protection removal compared to the modern XML-based formats.

## Troubleshooting

**`ModuleNotFoundError`** — run `pip install pypdf msoffcrypto-tool openpyxl lxml`

**`Error: Wrong password`** — double-check the password and try again

**`Unsupported file type`** — only `.pdf`, `.docx`, `.doc`, `.xlsx`, `.xlsm`, `.xls`, `.pptx`, `.ppt` are supported
