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

### Password options *(mutually exclusive)*

| Flag | Description |
|---|---|
| `--password PASSWORD`, `-p` | Password to open the file (omit if only edit/sheet protection is set) |
| `--password-file FILE` | Read the password from the first line of a file (avoids shell history exposure) |
| `--password-list WORDLIST` | Try every line of a file as a password (brute-force / forgot-password mode) |

If none of these are given and the file turns out to be encrypted, you'll be prompted securely via `getpass` when running interactively. In non-interactive (piped/scripted) contexts no prompt fires and the tool errors cleanly.

### Output options *(mutually exclusive)*

| Flag | Description |
|---|---|
| `--output PATH`, `-o` | Output path — single-file mode only (default: `unlocked_<filename>`) |
| `--in-place` | Overwrite the original file instead of writing a new one |
| `--output-dir DIR` | Write all unlocked files into this directory (created if needed) |

### Safety flags

| Flag | Description |
|---|---|
| `--backup` | Create a `.bak` copy before modifying (requires `--in-place`) |
| `--no-overwrite` | Skip files whose output already exists (useful for resumable batch jobs) |

### Inspection

| Flag | Description |
|---|---|
| `--check`, `--dry-run` | Report protection status without modifying any files |
| `--json` | With `--check`: emit machine-readable JSON instead of human-readable text |

### Directory walking

| Flag | Description |
|---|---|
| `--recursive`, `-r` | Expand `**` in glob patterns to walk entire directory trees |

### Error handling

| Flag | Description |
|---|---|
| `--fail-fast` | Stop on the first error (default: continue and summarise at the end) |

### Verbosity *(mutually exclusive)*

| Flag | Description |
|---|---|
| `--verbose`, `-v` | Show which XML elements were removed, which files were rewritten, etc. |
| `--quiet`, `-q` | Suppress all output except errors (useful for scripting) |

---

## Examples

```bash
# Check protection status without modifying anything
python unprotect.py document.pdf --check
python unprotect.py "*.xlsx" --check

# Machine-readable check, pipe to jq
python unprotect.py --check --json "*.xlsx" | jq '.[] | select(.status == "protected")'

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

# Overwrite the original file in place, with a safety backup
python unprotect.py report.docx -p mypassword --in-place --backup

# Process multiple files at once
python unprotect.py file1.xlsx file2.xlsx -p mypassword

# Batch unlock all Excel files in a directory tree
python unprotect.py "**/*.xlsx" --recursive --output-dir ./unlocked/

# Password from a file (avoids shell history exposure)
python unprotect.py secret.xlsx --password-file ~/.mysecret

# Brute-force a forgotten password
python unprotect.py locked.xlsx --password-list common_passwords.txt

# Skip files already unlocked from a previous run
python unprotect.py "*.docx" --output-dir ./out/ --no-overwrite

# Suppress all output (for use in scripts)
python unprotect.py "*.pdf" -p mypassword --output-dir ./out/ --quiet

# Stop on first failure (CI pipelines)
python unprotect.py "*.xlsx" --output-dir ./out/ --fail-fast
```

When processing more than one file, a summary is printed at the end:

```
Done: 47 unlocked, 2 already open, 1 failed
```

---

## --check / --json output

Human-readable:
```
[PROTECTED]  report.xlsx  — workbook structure, xl/worksheets/sheet1.xml
[OPEN]       slides.pptx  — no protection detected
```

JSON (`--check --json`):
```json
[
  {
    "file": "report.xlsx",
    "status": "protected",
    "layers": ["workbook structure", "xl/worksheets/sheet1.xml"],
    "message": "workbook structure, xl/worksheets/sheet1.xml"
  },
  {
    "file": "slides.pptx",
    "status": "open",
    "layers": [],
    "message": "no protection detected"
  }
]
```

---

## How It Works

Modern Office files (`.docx`, `.xlsx`, `.pptx`) are ZIP archives containing XML. Protection sits on two independent layers:

**Layer 1 — file encryption** (password required to open): handled by `msoffcrypto-tool`, which decrypts the file into a temporary copy before any further processing.

**Layer 2 — edit/structure protection** (file opens fine, but editing is locked): handled by direct XML manipulation via `lxml`. The relevant protection elements are removed from the XML inside the ZIP, rewriting only those entries while preserving compression and metadata for everything else.

For PDFs, `pypdf` handles both decryption and reconstruction. PDFs with only an owner password (print/edit restrictions, but no open password) are unlocked automatically without needing `--password`.

---

## Library API

`unprotect.py` can be imported directly into other Python projects. All functions raise `UnprotectError` instead of calling `sys.exit()`.

```python
from unprotect import unprotect_file, check_file, FileResult, UnprotectError

# Inspect a file (read-only)
result = check_file("report.xlsx")
# result.status  → "protected" | "open" | "failed"
# result.layers  → ["workbook structure", "xl/worksheets/sheet1.xml"]

# Unlock a file
try:
    result = unprotect_file(
        input_path="report.xlsx",
        password="secret",       # optional
        output_path="out.xlsx",  # optional; defaults to unlocked_report.xlsx
        in_place=False,
        output_dir=None,
        backup=False,
        no_overwrite=False,
    )
    # result.status → "unlocked" | "skipped"
except UnprotectError as e:
    print(f"Failed (code {e.code}): {e}")
```

---

## Exit Codes

| Code | Meaning |
|---|---|
| `0` | Success |
| `1` | General error (unsupported format, missing dependency, etc.) |
| `2` | Wrong or missing password / invalid argument combination |
| `3` | Unsupported file extension |
| `4` | File not found |

---

## Troubleshooting

**`ModuleNotFoundError`** — run `pip install pypdf msoffcrypto-tool lxml`

**`Error: Wrong password`** — double-check the password; the original file is left untouched

**`Error: PDF is encrypted and requires a password`** — pass the open password with `--password` / `-p`

**`Unsupported file type`** — only `.pdf`, `.docx`, `.xlsx`, `.xlsm`, `.pptx` are supported; convert legacy formats first

**`--output` rejected with multiple files** — `--output` only works for a single file; use `--output-dir` for batch jobs or omit it to get the default `unlocked_<filename>` naming

---

## License

MIT — see [LICENSE](LICENSE) for details.