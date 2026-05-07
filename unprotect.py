#!/usr/bin/env python3
"""
unprotect.py — Remove password protection from PDF and Office files.

Can also be used as a library:
    from unprotect import unprotect_file, check_file
"""

from __future__ import annotations

import getpass
import glob
import json
import logging
import os
import shutil
import sys
import zipfile
import argparse
from dataclasses import dataclass, field
from typing import Callable

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
# Two loggers:
#   log        — user-facing messages (INFO = normal, DEBUG = --verbose)
#   error_log  — always printed unless a caller suppresses it
# ---------------------------------------------------------------------------
log = logging.getLogger("unprotect")
_handler = logging.StreamHandler(sys.stdout)
_handler.setFormatter(logging.Formatter("%(message)s"))
log.addHandler(_handler)
log.setLevel(logging.INFO)


# ---------------------------------------------------------------------------
# Public result type (enables library use)
# ---------------------------------------------------------------------------
@dataclass
class FileResult:
    path: str
    status: str                      # "unlocked" | "open" | "protected" | "failed" | "skipped"
    message: str = ""
    layers: list[str] = field(default_factory=list)   # for --check / --json


# ---------------------------------------------------------------------------
# Internal helpers — no sys.exit() so they're library-safe
# ---------------------------------------------------------------------------

class UnprotectError(Exception):
    """Raised instead of sys.exit() so library callers can catch it."""
    def __init__(self, message: str, code: int = 1):
        super().__init__(message)
        self.code = code


def _msoffcrypto_decrypt(input_path: str, password: str, tmp_path: str) -> bool:
    """Decrypt an Office file using msoffcrypto. Returns True if it was encrypted."""
    try:
        import msoffcrypto
    except ImportError:
        raise UnprotectError("Missing dependency! Please run: pip install msoffcrypto-tool")

    with open(input_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        if not office_file.is_encrypted():
            return False
        if not password:
            raise UnprotectError(
                "File is encrypted but no password was provided. "
                "Use --password, --password-file, or supply one interactively.",
                code=2,
            )
        try:
            office_file.load_key(password=password)
            with open(tmp_path, "wb") as out:
                office_file.decrypt(out)
        except Exception as e:
            raise UnprotectError(f"Couldn't decrypt file — {e}", code=2)
    return True


def _cleanup(path: str):
    if path and os.path.exists(path):
        os.remove(path)


def _rewrite_zip(zip_path: str, filename_in_zip: str, new_content: bytes):
    tmp_zip = zip_path + ".zip.tmp"
    try:
        with zipfile.ZipFile(zip_path, "r") as zin, \
                zipfile.ZipFile(tmp_zip, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == filename_in_zip:
                    data = new_content
                zout.writestr(item, data, compress_type=item.compress_type)
        os.replace(tmp_zip, zip_path)
        log.debug("  Rewrote %s inside zip", filename_in_zip)
    except Exception:
        _cleanup(tmp_zip)
        raise


def _resolve_output(
    input_path: str,
    output_arg: str | None,
    in_place: bool,
    output_dir: str | None,
) -> str:
    if in_place:
        return input_path
    if output_arg:
        return output_arg
    base = os.path.basename(input_path)
    name, ext = os.path.splitext(base)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        return os.path.join(output_dir, f"unlocked_{name}{ext}")
    directory = os.path.dirname(input_path) or "."
    return os.path.join(directory, f"unlocked_{name}{ext}")


def _check_collision(input_path: str, output_path: str, in_place: bool):
    if in_place:
        return
    if os.path.realpath(input_path) == os.path.realpath(output_path):
        raise UnprotectError(
            "Input and output paths resolve to the same file. "
            "Use --in-place to overwrite, or choose a different --output path."
        )


def _make_backup(path: str) -> str:
    backup = path + ".bak"
    shutil.copy2(path, backup)
    log.debug("  Backup created: %s", backup)
    return backup


# ---------------------------------------------------------------------------
# Protection detection helpers
# ---------------------------------------------------------------------------

def _detect_xml_layers(input_path: str, ext: str) -> list[str]:
    """Return a list of protection layer names found inside the zip."""
    layers: list[str] = []
    try:
        with zipfile.ZipFile(input_path, "r") as z:
            names = z.namelist()
            import lxml.etree as etree

            if ext in (".xlsx", ".xlsm"):
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                wb_name = next((n for n in names if n.endswith("workbook.xml")), None)
                if wb_name:
                    root = etree.fromstring(z.read(wb_name))
                    if root.find(f"{{{ns}}}workbookProtection") is not None:
                        layers.append("workbook structure")
                for n in names:
                    if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                        root = etree.fromstring(z.read(n))
                        if root.find(f"{{{ns}}}sheetProtection") is not None:
                            layers.append(n)

            elif ext == ".docx":
                ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                settings_name = next((n for n in names if n.endswith("settings.xml")), None)
                if settings_name:
                    root = etree.fromstring(z.read(settings_name))
                    for tag in ("documentProtection", "writeProtection"):
                        if root.find(f"{{{ns}}}{tag}") is not None:
                            layers.append(tag)

            elif ext == ".pptx":
                ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
                prs_name = next((n for n in names if n.endswith("presentation.xml")), None)
                if prs_name:
                    root = etree.fromstring(z.read(prs_name))
                    for tag in ("modifyVerifier", "writeProtection"):
                        if root.find(f"{{{ns}}}{tag}") is not None:
                            layers.append(tag)
    except Exception as e:
        log.debug("  Could not inspect XML layers: %s", e)
    return layers


def check_file(input_path: str, password: str | None = None) -> FileResult:
    """
    Public API: inspect a file and return a FileResult describing its protection state.
    Does not modify any file.
    """
    ext = os.path.splitext(input_path)[1].lower()

    if ext in (".xlsx", ".xlsm", ".docx", ".pptx"):
        try:
            import msoffcrypto
            with open(input_path, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                encrypted = of.is_encrypted()
        except ImportError:
            encrypted = False
        except Exception:
            encrypted = False

        if encrypted:
            return FileResult(
                path=input_path,
                status="protected",
                message="file-level password required to open",
                layers=["file encryption"],
            )
        layers = _detect_xml_layers(input_path, ext)
        if layers:
            return FileResult(
                path=input_path,
                status="protected",
                message=", ".join(layers),
                layers=layers,
            )
        return FileResult(path=input_path, status="open", message="no protection detected")

    elif ext == ".pdf":
        try:
            from pypdf import PdfReader
            r = PdfReader(input_path)
            if r.is_encrypted:
                return FileResult(
                    path=input_path,
                    status="protected",
                    message="PDF open password set",
                    layers=["PDF encryption"],
                )
            return FileResult(path=input_path, status="open", message="no PDF encryption detected")
        except ImportError:
            return FileResult(path=input_path, status="failed", message="Missing dependency: pip install pypdf")

    else:
        return FileResult(
            path=input_path,
            status="failed",
            message=f"unsupported extension '{ext}'",
        )


def _print_check_result(result: FileResult, use_json: bool, json_accumulator: list | None):
    if use_json:
        if json_accumulator is not None:
            json_accumulator.append({
                "file": result.path,
                "status": result.status,
                "layers": result.layers,
                "message": result.message,
            })
        return

    label_map = {
        "protected": "[PROTECTED]",
        "open":      "[OPEN]     ",
        "failed":    "[ERROR]    ",
    }
    label = label_map.get(result.status, "[UNKNOWN]  ")
    log.info("%s  %s  — %s", label, result.path, result.message)


# ---------------------------------------------------------------------------
# Unprotect implementations
# ---------------------------------------------------------------------------

def unprotect_pdf(input_path: str, password: str | None, output_path: str) -> FileResult:
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        raise UnprotectError("Missing dependency! Please run: pip install pypdf")

    reader = PdfReader(input_path)
    if reader.is_encrypted:
        result = reader.decrypt(password or "")
        if result == 0:
            if password:
                raise UnprotectError("Wrong password!", code=2)
            else:
                raise UnprotectError("PDF is encrypted and requires a password. Use --password.", code=1)

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)

    log.debug("  PDF pages copied, encryption stripped")
    return FileResult(path=output_path, status="unlocked", message=f"PDF unprotected → {output_path}")


def _strip_excel_xml_protection(xlsx_path: str):
    import lxml.etree as etree

    with zipfile.ZipFile(xlsx_path, "r") as z:
        names = z.namelist()
        wb_name = next((n for n in names if n.endswith("workbook.xml")), None)
        wb_xml = z.read(wb_name) if wb_name else None
        sheet_names = [n for n in names
                       if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]

    if wb_name and wb_xml is not None:
        wb_root = etree.fromstring(wb_xml)
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        removed = [el for el in wb_root.findall(f"{{{ns}}}workbookProtection")]
        for el in removed:
            wb_root.remove(el)
            log.debug("  Removed workbookProtection from %s", wb_name)
        if removed:
            new_xml = etree.tostring(wb_root, xml_declaration=True, encoding="UTF-8", standalone=True)
            _rewrite_zip(xlsx_path, wb_name, new_xml)

    for sheet_name in sheet_names:
        with zipfile.ZipFile(xlsx_path, "r") as z:
            sheet_xml = z.read(sheet_name)
        root = etree.fromstring(sheet_xml)
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        removed = [el for el in root.findall(f"{{{ns}}}sheetProtection")]
        for el in removed:
            root.remove(el)
            log.debug("  Removed sheetProtection from %s", sheet_name)
        if removed:
            new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            _rewrite_zip(xlsx_path, sheet_name, new_xml)


def unprotect_excel(input_path: str, password: str | None, output_path: str) -> FileResult:
    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".xls":
        raise UnprotectError("Legacy .xls format is not supported.")

    tmp_path = output_path + ".tmp.xlsx"
    try:
        was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_path)
        work_path = tmp_path if was_encrypted else input_path
        if os.path.realpath(work_path) != os.path.realpath(output_path):
            shutil.copy2(work_path, output_path)
        _strip_excel_xml_protection(output_path)
    finally:
        _cleanup(tmp_path)

    return FileResult(path=output_path, status="unlocked", message=f"Excel unprotected → {output_path}")


def unprotect_word(input_path: str, password: str | None, output_path: str) -> FileResult:
    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".doc":
        raise UnprotectError("Legacy .doc format is not supported.")

    tmp_path = output_path + ".tmp.docx"
    try:
        was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_path)
        work_path = tmp_path if was_encrypted else input_path
        if os.path.realpath(work_path) != os.path.realpath(output_path):
            shutil.copy2(work_path, output_path)

        import lxml.etree as etree
        with zipfile.ZipFile(output_path, "r") as z:
            settings_name = next((n for n in z.namelist() if n.endswith("settings.xml")), None)
            if settings_name is None:
                return FileResult(path=output_path, status="unlocked",
                                  message=f"Word unprotected (no settings.xml) → {output_path}")
            settings_xml = z.read(settings_name)

        root = etree.fromstring(settings_xml)
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        changed = False
        for tag in ("documentProtection", "writeProtection"):
            for el in root.findall(f"{{{ns}}}{tag}"):
                root.remove(el)
                log.debug("  Removed %s from %s", tag, settings_name)
                changed = True
        if changed:
            new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            _rewrite_zip(output_path, settings_name, new_xml)
    finally:
        _cleanup(tmp_path)

    return FileResult(path=output_path, status="unlocked", message=f"Word unprotected → {output_path}")


def unprotect_powerpoint(input_path: str, password: str | None, output_path: str) -> FileResult:
    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".ppt":
        raise UnprotectError("Legacy .ppt format is not supported.")

    tmp_path = output_path + ".tmp.pptx"
    try:
        was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_path)
        work_path = tmp_path if was_encrypted else input_path
        if os.path.realpath(work_path) != os.path.realpath(output_path):
            shutil.copy2(work_path, output_path)

        import lxml.etree as etree
        with zipfile.ZipFile(output_path, "r") as z:
            names = z.namelist()
            prs_name = next((n for n in names if n.endswith("presentation.xml")), None)
            if prs_name is None:
                return FileResult(path=output_path, status="unlocked",
                                  message=f"PowerPoint unprotected (no presentation.xml) → {output_path}")
            prs_xml = z.read(prs_name)

        root = etree.fromstring(prs_xml)
        ns_pml = "http://schemas.openxmlformats.org/presentationml/2006/main"
        changed = False
        for tag in ("modifyVerifier", "writeProtection"):
            for el in root.findall(f"{{{ns_pml}}}{tag}"):
                root.remove(el)
                log.debug("  Removed %s from %s", tag, prs_name)
                changed = True
        if changed:
            new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            _rewrite_zip(output_path, prs_name, new_xml)

        with zipfile.ZipFile(output_path, "r") as z:
            slide_names = [n for n in z.namelist()
                           if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
        for slide_name in slide_names:
            with zipfile.ZipFile(output_path, "r") as z:
                slide_xml = z.read(slide_name)
            slide_root = etree.fromstring(slide_xml)
            mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
            slide_changed = False
            for ac in slide_root.findall(f".//{{{mc_ns}}}AlternateContent"):
                parent = ac.getparent()
                if parent is not None:
                    raw = etree.tostring(ac)
                    if b"oleObj" in raw and b"locked" in raw:
                        parent.remove(ac)
                        log.debug("  Removed locked AlternateContent from %s", slide_name)
                        slide_changed = True
            if slide_changed:
                new_slide_xml = etree.tostring(slide_root, xml_declaration=True,
                                               encoding="UTF-8", standalone=True)
                _rewrite_zip(output_path, slide_name, new_slide_xml)
    finally:
        _cleanup(tmp_path)

    return FileResult(path=output_path, status="unlocked", message=f"PowerPoint unprotected → {output_path}")


# ---------------------------------------------------------------------------
# Supported format registry
# ---------------------------------------------------------------------------
SUPPORTED: dict[str, tuple[str, Callable]] = {
    ".pdf":  ("PDF",        unprotect_pdf),
    ".xlsx": ("Excel",      unprotect_excel),
    ".xlsm": ("Excel",      unprotect_excel),
    ".xls":  ("Excel",      unprotect_excel),
    ".docx": ("Word",       unprotect_word),
    ".doc":  ("Word",       unprotect_word),
    ".pptx": ("PowerPoint", unprotect_powerpoint),
    ".ppt":  ("PowerPoint", unprotect_powerpoint),
}


# ---------------------------------------------------------------------------
# Public library API
# ---------------------------------------------------------------------------

def unprotect_file(
    input_path: str,
    password: str | None = None,
    output_path: str | None = None,
    in_place: bool = False,
    output_dir: str | None = None,
    backup: bool = False,
    no_overwrite: bool = False,
) -> FileResult:
    """
    Remove protection from a single file.

    Parameters
    ----------
    input_path  : Path to the protected file.
    password    : Open password, if any.
    output_path : Explicit output path (single-file mode).
    in_place    : Overwrite the original file.
    output_dir  : Directory for output (batch-friendly).
    backup      : Create a .bak copy before modifying (only with in_place).
    no_overwrite: Skip if the output file already exists.

    Returns
    -------
    FileResult with status "unlocked", "open", "skipped", or raises UnprotectError.
    """
    if not os.path.exists(input_path):
        raise UnprotectError(f"File not found: {input_path}", code=4)

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED:
        supported_list = ", ".join(SUPPORTED.keys())
        raise UnprotectError(
            f"Unsupported file type '{ext}'. Supported: {supported_list}", code=3
        )

    label, handler = SUPPORTED[ext]
    resolved_output = _resolve_output(input_path, output_path, in_place, output_dir)
    _check_collision(input_path, resolved_output, in_place)

    if no_overwrite and os.path.exists(resolved_output):
        return FileResult(
            path=input_path,
            status="skipped",
            message=f"output already exists, skipping: {resolved_output}",
        )

    if backup and in_place and os.path.exists(input_path):
        _make_backup(input_path)

    log.debug("Processing %s file: %s → %s", label, input_path, resolved_output)
    return handler(input_path, password, resolved_output)


# ---------------------------------------------------------------------------
# Password helpers
# ---------------------------------------------------------------------------

def _load_password_file(path: str) -> str:
    with open(path, "r") as f:
        line = f.readline().rstrip("\n")
    if not line:
        raise UnprotectError(f"--password-file '{path}' is empty.")
    return line


def _try_password_list(
    input_path: str,
    wordlist_path: str,
    output_path: str | None,
    in_place: bool,
    output_dir: str | None,
    backup: bool,
    no_overwrite: bool,
) -> FileResult:
    """Try each line of wordlist_path as a password; return on first success."""
    with open(wordlist_path, "r", errors="replace") as f:
        passwords = [line.rstrip("\n") for line in f if line.strip()]

    log.info("Trying %d passwords from %s …", len(passwords), wordlist_path)
    for i, pwd in enumerate(passwords, 1):
        log.debug("  Attempt %d/%d: %s", i, len(passwords), pwd)
        try:
            result = unprotect_file(
                input_path=input_path,
                password=pwd,
                output_path=output_path,
                in_place=in_place,
                output_dir=output_dir,
                backup=backup,
                no_overwrite=no_overwrite,
            )
            log.info("  Password found: %s (attempt %d/%d)", pwd, i, len(passwords))
            return result
        except UnprotectError as e:
            if e.code == 2:   # wrong password — keep trying
                continue
            raise             # other errors bubble up

    raise UnprotectError(
        f"No password from '{wordlist_path}' worked ({len(passwords)} tried).", code=2
    )


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def _expand_paths(patterns: list[str], recursive: bool) -> list[str]:
    paths: list[str] = []
    for pattern in patterns:
        expanded = glob.glob(pattern, recursive=recursive)
        if expanded:
            paths.extend(expanded)
        else:
            paths.append(pattern)  # keep as-is so "file not found" error fires naturally
    # Deduplicate while preserving order
    seen: set[str] = set()
    result: list[str] = []
    for p in paths:
        rp = os.path.realpath(p)
        if rp not in seen:
            seen.add(rp)
            result.append(p)
    return result


def main(argv: list[str] | None = None):
    parser = argparse.ArgumentParser(
        description="Remove password protection from PDF and Office files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  unprotect.py report.xlsx --password secret
  unprotect.py "*.xlsx" --output-dir ./unlocked/ --quiet
  unprotect.py "**/*.docx" --recursive --in-place --backup
  unprotect.py report.xlsx --password-list words.txt
  unprotect.py *.pdf --check --json
""",
    )

    # Positional
    parser.add_argument("files", nargs="+",
                        help="Path(s) or glob pattern(s) to protected file(s)")

    # Password options (mutually exclusive group)
    pw_group = parser.add_mutually_exclusive_group()
    pw_group.add_argument("--password", "-p", default=None,
                          help="Password to unlock the file")
    pw_group.add_argument("--password-file", metavar="FILE",
                          help="File containing the password (first line used)")
    pw_group.add_argument("--password-list", metavar="WORDLIST",
                          help="Try each line of WORDLIST as a password (brute-force)")

    # Output options
    out_group = parser.add_mutually_exclusive_group()
    out_group.add_argument("--output", "-o", default=None,
                           help="Output path (single-file mode only)")
    out_group.add_argument("--in-place", action="store_true",
                           help="Overwrite the original file(s)")
    out_group.add_argument("--output-dir", metavar="DIR",
                           help="Directory to write unlocked files into")

    # Safety
    parser.add_argument("--backup", action="store_true",
                        help="Create a .bak copy before modifying (with --in-place)")
    parser.add_argument("--no-overwrite", action="store_true",
                        help="Skip files whose output path already exists")

    # Check / dry-run
    parser.add_argument("--check", "--dry-run", action="store_true",
                        help="Inspect protection without modifying any file (alias: --dry-run)")
    parser.add_argument("--json", dest="json_output", action="store_true",
                        help="Emit machine-readable JSON for --check output")

    # Directory walking
    parser.add_argument("--recursive", "-r", action="store_true",
                        help="Expand ** in glob patterns recursively")

    # Error handling
    parser.add_argument("--fail-fast", action="store_true",
                        help="Stop immediately on the first error")

    # Verbosity (mutually exclusive)
    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument("--verbose", "-v", action="store_true",
                           help="Show detailed per-element info")
    verbosity.add_argument("--quiet", "-q", action="store_true",
                           help="Suppress all output except errors")

    args = parser.parse_args(argv)

    # ---- Configure logging level ----------------------------------------
    if args.verbose:
        log.setLevel(logging.DEBUG)
    elif args.quiet:
        log.setLevel(logging.ERROR)
    else:
        log.setLevel(logging.INFO)

    # ---- Mutual-exclusion guards -----------------------------------------
    if args.backup and not args.in_place:
        parser.error("--backup only makes sense with --in-place")

    if args.json_output and not args.check:
        parser.error("--json only makes sense with --check / --dry-run")

    # ---- Resolve password ------------------------------------------------
    password: str | None = args.password

    if args.password_file:
        try:
            password = _load_password_file(args.password_file)
        except (OSError, UnprotectError) as e:
            log.error("Error reading --password-file: %s", e)
            sys.exit(1)

    # (--password-list is handled per-file below; no global password needed)

    # ---- Expand file patterns -------------------------------------------
    input_paths = _expand_paths(args.files, recursive=args.recursive)

    if args.output and len(input_paths) > 1:
        parser.error("--output can only be used when processing a single file.")

    # ---- Interactive password prompt (TTY only) --------------------------
    # Prompt once before the loop if we need a password and stdin is a TTY.
    if (
        not args.check
        and not args.password_list
        and password is None
        and sys.stdin.isatty()
        and len(input_paths) > 0
    ):
        try:
            prompted = getpass.getpass("Password (leave blank if none): ")
            if prompted:
                password = prompted
        except (KeyboardInterrupt, EOFError):
            print()
            sys.exit(130)

    # ---- Main processing loop -------------------------------------------
    json_results: list[dict] = []
    counts = {"unlocked": 0, "open": 0, "skipped": 0, "failed": 0}
    exit_code = 0

    for path in input_paths:
        try:
            if args.check:
                result = check_file(path, password)
                _print_check_result(result, args.json_output, json_results)
                counts[result.status if result.status in counts else "failed"] += 1

            elif args.password_list:
                result = _try_password_list(
                    input_path=path,
                    wordlist_path=args.password_list,
                    output_path=args.output,
                    in_place=args.in_place,
                    output_dir=args.output_dir,
                    backup=args.backup,
                    no_overwrite=args.no_overwrite,
                )
                log.info("%s", result.message)
                counts[result.status if result.status in counts else "failed"] += 1

            else:
                result = unprotect_file(
                    input_path=path,
                    password=password,
                    output_path=args.output,
                    in_place=args.in_place,
                    output_dir=args.output_dir,
                    backup=args.backup,
                    no_overwrite=args.no_overwrite,
                )
                log.info("%s", result.message)
                if result.status == "skipped":
                    log.info("  Skipped: %s", result.message)
                counts[result.status if result.status in counts else "failed"] += 1

        except UnprotectError as e:
            log.error("Error [%s]: %s", path, e)
            counts["failed"] += 1
            exit_code = e.code
            if args.fail_fast:
                log.error("--fail-fast: stopping after first error.")
                break
        except Exception as e:
            log.error("Unexpected error [%s]: %s", path, e)
            counts["failed"] += 1
            exit_code = 1
            if args.fail_fast:
                log.error("--fail-fast: stopping after first error.")
                break

    # ---- JSON output (--check --json) -----------------------------------
    if args.json_output:
        print(json.dumps(json_results, indent=2))

    # ---- Summary (batch jobs with >1 file) ------------------------------
    if len(input_paths) > 1 and not args.quiet:
        parts: list[str] = []
        if counts["unlocked"]:
            parts.append(f"{counts['unlocked']} unlocked")
        if counts["open"]:
            parts.append(f"{counts['open']} already open")
        if counts["skipped"]:
            parts.append(f"{counts['skipped']} skipped (output exists)")
        if counts["failed"]:
            parts.append(f"{counts['failed']} failed")
        if not args.json_output:
            log.info("\nDone: %s", ", ".join(parts) if parts else "nothing processed")

    sys.exit(exit_code)


if __name__ == "__main__":
    main()