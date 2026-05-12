#!/usr/bin/env python3
"""
unprotect.py — Remove password protection from PDF and Office files.

Enhanced version with legacy binary support, conversion, extra protection removal,
metadata preservation, PDF cleaning, and parallel processing.
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
import subprocess
from dataclasses import dataclass, field
from typing import Callable
from concurrent.futures import ProcessPoolExecutor, as_completed

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
log = logging.getLogger("unprotect")
_handler = logging.StreamHandler(sys.stdout)
_handler.setFormatter(logging.Formatter("%(message)s"))
log.addHandler(_handler)
log.setLevel(logging.INFO)


# ---------------------------------------------------------------------------
# Public result type
# ---------------------------------------------------------------------------
@dataclass
class FileResult:
    path: str
    status: str                      # "unlocked" | "open" | "protected" | "failed" | "skipped"
    message: str = ""
    layers: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------
class UnprotectError(Exception):
    def __init__(self, message: str, code: int = 1):
        super().__init__(message)
        self.code = code


def _msoffcrypto_decrypt(input_path: str, password: str, tmp_path: str) -> bool:
    try:
        import msoffcrypto
    except ImportError:
        raise UnprotectError("Missing dependency: pip install msoffcrypto-tool")

    with open(input_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        if not office_file.is_encrypted():
            return False
        if not password:
            raise UnprotectError(
                "File is encrypted but no password provided.", code=2
            )
        try:
            office_file.load_key(password=password)
            with open(tmp_path, "wb") as out:
                office_file.decrypt(out)
        except Exception as e:
            raise UnprotectError(f"Decryption failed: {e}", code=2)
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
    convert: bool = False,
) -> str:
    if in_place:
        return input_path
    if output_arg:
        return output_arg
    base = os.path.basename(input_path)
    name, ext = os.path.splitext(base)
    if convert:
        new_ext = {'.doc': '.docx', '.xls': '.xlsx', '.ppt': '.pptx'}.get(ext, ext)
        base = name + new_ext
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        return os.path.join(output_dir, f"unlocked_{base}" if not convert else base)
    directory = os.path.dirname(input_path) or "."
    return os.path.join(directory, f"unlocked_{base}" if not convert else base)


def _check_collision(input_path: str, output_path: str, in_place: bool):
    if in_place:
        return
    if os.path.realpath(input_path) == os.path.realpath(output_path):
        raise UnprotectError(
            "Input and output resolve to same file. Use --in-place to overwrite."
        )


def _make_backup(path: str) -> str:
    backup = path + ".bak"
    shutil.copy2(path, backup)
    log.debug("  Backup created: %s", backup)
    return backup


def _convert_with_libreoffice(input_path: str, output_path: str) -> None:
    if not shutil.which("libreoffice"):
        raise UnprotectError(
            "libreoffice not found in PATH. Install LibreOffice to use --convert."
        )
    out_dir = os.path.dirname(output_path) or "."
    out_ext = os.path.splitext(output_path)[1][1:]
    cmd = [
        "libreoffice", "--headless", "--convert-to", out_ext,
        "--outdir", out_dir, input_path
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        converted_name = os.path.basename(input_path)
        name_no_ext = os.path.splitext(converted_name)[0]
        expected = os.path.join(out_dir, f"{name_no_ext}.{out_ext}")
        if os.path.exists(expected) and expected != output_path:
            shutil.move(expected, output_path)
    except subprocess.CalledProcessError as e:
        raise UnprotectError(f"Conversion failed: {e.stderr}")


# ---------------------------------------------------------------------------
# XML stripping helpers
# ---------------------------------------------------------------------------
def _strip_excel_xml_protection(xlsx_path: str):
    try:
        import lxml.etree as etree
    except ImportError:
        log.warning("lxml not installed, cannot strip Excel XML protections")
        return
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


def _strip_word_xml_protection(docx_path: str):
    try:
        import lxml.etree as etree
    except ImportError:
        log.warning("lxml not installed, cannot strip Word XML protections")
        return
    with zipfile.ZipFile(docx_path, "r") as z:
        settings_name = next((n for n in z.namelist() if n.endswith("settings.xml")), None)
        if settings_name is None:
            return
        settings_xml = z.read(settings_name)
    root = etree.fromstring(settings_xml)
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    changed = False
    for tag in ("documentProtection", "writeProtection", "readOnlyRecommended"):
        for el in root.findall(f"{{{ns}}}{tag}"):
            root.remove(el)
            log.debug("  Removed %s from %s", tag, settings_name)
            changed = True
    if changed:
        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _rewrite_zip(docx_path, settings_name, new_xml)


def _strip_pptx_protections(pptx_path: str):
    try:
        import lxml.etree as etree
    except ImportError:
        log.warning("lxml not installed, cannot strip PowerPoint XML protections")
        return
    with zipfile.ZipFile(pptx_path, "r") as z:
        names = z.namelist()
        prs_name = next((n for n in names if n.endswith("presentation.xml")), None)
        if prs_name is None:
            return
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
        _rewrite_zip(pptx_path, prs_name, new_xml)

    # Remove locked OLE objects
    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_names = [n for n in z.namelist()
                       if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    for slide_name in slide_names:
        with zipfile.ZipFile(pptx_path, "r") as z:
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
            _rewrite_zip(pptx_path, slide_name, new_slide_xml)


def _remove_extra_protections_office(zip_path: str, file_type: str):
    """Remove digital signatures, etc."""
    with zipfile.ZipFile(zip_path, "r") as z:
        names = z.namelist()
        sig_entries = [n for n in names if "_xmlsignatures" in n or n.endswith("signatures.xml")]
    for entry in sig_entries:
        _rewrite_zip(zip_path, entry, b"")
        log.debug("  Removed signature entry: %s", entry)


# ---------------------------------------------------------------------------
# Protection detection (extended)
# ---------------------------------------------------------------------------
def _detect_xml_layers(input_path: str, ext: str) -> list[str]:
    layers = []
    try:
        import lxml.etree as etree
    except ImportError:
        return layers  # cannot detect without lxml

    try:
        with zipfile.ZipFile(input_path, "r") as z:
            names = z.namelist()

            if ext in (".xlsx", ".xlsm"):
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                wb_name = next((n for n in names if n.endswith("workbook.xml")), None)
                if wb_name:
                    root = etree.fromstring(z.read(wb_name))
                    wb_prot = root.find(f"{{{ns}}}workbookProtection")
                    if wb_prot is not None:
                        if wb_prot.get("lockStructure") == "1":
                            layers.append("workbook structure locked")
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
                    for tag in ("documentProtection", "writeProtection", "readOnlyRecommended"):
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

            if any("signature" in name.lower() for name in names):
                layers.append("digital signature")
    except Exception:
        pass
    return layers


def check_file(input_path: str, password: str | None = None) -> FileResult:
    ext = os.path.splitext(input_path)[1].lower()

    if ext in (".xlsx", ".xlsm", ".docx", ".pptx", ".doc", ".xls", ".ppt"):
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
                message="file-level encryption",
                layers=["file encryption"],
            )
        if ext in (".xlsx", ".xlsm", ".docx", ".pptx"):
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
            reader = PdfReader(input_path)
            if reader.is_encrypted:
                return FileResult(
                    path=input_path,
                    status="protected",
                    message="PDF encryption (open or owner password)",
                    layers=["PDF encryption"],
                )
            return FileResult(path=input_path, status="open", message="no PDF encryption")
        except ImportError:
            return FileResult(path=input_path, status="failed", message="pip install pypdf")

    return FileResult(path=input_path, status="failed", message=f"unsupported extension '{ext}'")


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
# Unprotect implementations (enhanced & fixed)
# ---------------------------------------------------------------------------
def unprotect_pdf(input_path: str, password: str | None, output_path: str, **kwargs) -> FileResult:
    try:
        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import ArrayObject, NameObject
    except ImportError:
        raise UnprotectError("Missing dependency: pip install pypdf")

    reader = PdfReader(input_path)

    # Handle encrypted PDFs
    if reader.is_encrypted:
        success = reader.decrypt(password or "")

        # Retry empty password for owner-only encryption
        if success == 0 and password:
            success = reader.decrypt("")

        if success == 0:
            raise UnprotectError(
                "Wrong password (or file uses unsupported encryption)",
                code=2,
            )

    writer = PdfWriter()

    # Add pages, removing redaction annotations on the fly
    for page in reader.pages:
        if "/Annots" in page:
            filtered_annots = ArrayObject(
                [
                    annot
                    for annot in page["/Annots"]
                    if annot.get("/Subtype") != "/Redact"
                ]
            )
            page[NameObject("/Annots")] = filtered_annots
            log.debug("  Removed redaction annotations")
        writer.add_page(page)

    # Remove JavaScript and form restrictions from writer root
    root = writer._root_object

    # Remove JavaScript
    if "/Names" in root:
        names = root["/Names"]
        if "/JavaScript" in names:
            del names["/JavaScript"]
            log.debug("  Removed embedded JavaScript")
        if len(names) == 0:
            del root["/Names"]

    # Remove AcroForm
    if "/AcroForm" in root:
        del root["/AcroForm"]
        log.debug("  Removed AcroForm")

    # Write cleaned PDF
    with open(output_path, "wb") as f:
        writer.write(f)

    return FileResult(
        path=output_path,
        status="unlocked",
        message=f"PDF cleaned → {output_path}",
    )


def _handle_office_common(
    input_path: str,
    password: str | None,
    output_path: str,
    convert: bool,
    is_legacy: bool,
    strip_func: Callable,
    new_ext: str,
    log_label: str,
) -> FileResult:
    """
    Common logic for Excel, Word, PowerPoint.
    - If legacy and convert: decrypt (if needed) -> convert -> strip -> move.
    - If legacy and not convert: decrypt (if needed) -> copy (no stripping).
    - If OpenXML: decrypt (if needed) -> copy -> strip.
    """
    ext = os.path.splitext(input_path)[1].lower()
    tmp_decrypted = None
    tmp_converted = None

    try:
        if is_legacy and convert:
            # --- Legacy to OpenXML conversion ---
            # Step 1: decrypt to a temporary legacy file (if encrypted)
            tmp_legacy = output_path + ".legacy"
            was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_legacy)
            source_for_conversion = tmp_legacy if was_encrypted else input_path

            # Step 2: convert to new format (OpenXML)
            tmp_converted = output_path + ".tmp" + new_ext
            _convert_with_libreoffice(source_for_conversion, tmp_converted)

            # Step 3: strip XML protection from the converted file
            strip_func(tmp_converted)
            _remove_extra_protections_office(tmp_converted, log_label)

            # Step 4: move to final output
            shutil.move(tmp_converted, output_path)
            log.info("  Converted %s → %s", ext, new_ext)

            # Cleanup
            if was_encrypted:
                _cleanup(tmp_legacy)
            return FileResult(path=output_path, status="unlocked", message=f"{log_label} unprotected & converted → {output_path}")

        else:
            # --- Legacy without conversion, or OpenXML ---
            # Decrypt if needed
            tmp_decrypted = output_path + ".tmp" + (new_ext if not is_legacy else ext)
            was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_decrypted)
            work_path = tmp_decrypted if was_encrypted else input_path

            # Copy to output (avoid self-copy when in-place)
            if os.path.realpath(work_path) != os.path.realpath(output_path):
                shutil.copy2(work_path, output_path)

            # Strip XML protection only for OpenXML files (not legacy binaries)
            if not is_legacy:
                strip_func(output_path)
                _remove_extra_protections_office(output_path, log_label)

            return FileResult(path=output_path, status="unlocked", message=f"{log_label} unprotected → {output_path}")

    finally:
        _cleanup(tmp_decrypted)
        _cleanup(tmp_converted)


def unprotect_excel(input_path: str, password: str | None, output_path: str, convert: bool = False, **kwargs) -> FileResult:
    ext = os.path.splitext(input_path)[1].lower()
    return _handle_office_common(
        input_path, password, output_path, convert,
        is_legacy=(ext == ".xls"),
        strip_func=_strip_excel_xml_protection,
        new_ext=".xlsx",
        log_label="Excel"
    )


def unprotect_word(input_path: str, password: str | None, output_path: str, convert: bool = False, **kwargs) -> FileResult:
    ext = os.path.splitext(input_path)[1].lower()
    return _handle_office_common(
        input_path, password, output_path, convert,
        is_legacy=(ext == ".doc"),
        strip_func=_strip_word_xml_protection,
        new_ext=".docx",
        log_label="Word"
    )


def unprotect_powerpoint(input_path: str, password: str | None, output_path: str, convert: bool = False, **kwargs) -> FileResult:
    ext = os.path.splitext(input_path)[1].lower()
    return _handle_office_common(
        input_path, password, output_path, convert,
        is_legacy=(ext == ".ppt"),
        strip_func=_strip_pptx_protections,
        new_ext=".pptx",
        log_label="PowerPoint"
    )


# Aliases for legacy formats (they just reuse the above logic with convert flag)
def unprotect_doc(input_path: str, password: str | None, output_path: str, convert: bool = False, **kwargs) -> FileResult:
    return unprotect_word(input_path, password, output_path, convert)


def unprotect_xls(input_path: str, password: str | None, output_path: str, convert: bool = False, **kwargs) -> FileResult:
    return unprotect_excel(input_path, password, output_path, convert)


def unprotect_ppt(input_path: str, password: str | None, output_path: str, convert: bool = False, **kwargs) -> FileResult:
    return unprotect_powerpoint(input_path, password, output_path, convert)


# ---------------------------------------------------------------------------
# Supported format registry (extended)
# ---------------------------------------------------------------------------
SUPPORTED: dict[str, tuple[str, Callable]] = {
    ".pdf":  ("PDF",        unprotect_pdf),
    ".xlsx": ("Excel",      unprotect_excel),
    ".xlsm": ("Excel",      unprotect_excel),
    ".xls":  ("Excel (legacy)", unprotect_xls),
    ".docx": ("Word",       unprotect_word),
    ".doc":  ("Word (legacy)",   unprotect_doc),
    ".pptx": ("PowerPoint", unprotect_powerpoint),
    ".ppt":  ("PowerPoint (legacy)", unprotect_ppt),
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
    convert: bool = False,
) -> FileResult:
    if not os.path.exists(input_path):
        raise UnprotectError(f"File not found: {input_path}", code=4)

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED:
        raise UnprotectError(f"Unsupported file type '{ext}'. Supported: {', '.join(SUPPORTED.keys())}", code=3)

    label, handler = SUPPORTED[ext]
    resolved_output = _resolve_output(input_path, output_path, in_place, output_dir, convert)
    _check_collision(input_path, resolved_output, in_place)

    if no_overwrite and os.path.exists(resolved_output):
        return FileResult(path=input_path, status="skipped", message=f"output exists: {resolved_output}")

    if backup and in_place and os.path.exists(input_path):
        _make_backup(input_path)

    log.debug("Processing %s: %s → %s", label, input_path, resolved_output)
    return handler(input_path, password, resolved_output, convert=convert)


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
    convert: bool,
) -> FileResult:
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
                convert=convert,
            )
            log.info("  Password found: %s (attempt %d/%d)", pwd, i, len(passwords))
            return result
        except UnprotectError as e:
            if e.code == 2:
                continue
            raise
    raise UnprotectError(f"No password from '{wordlist_path}' worked ({len(passwords)} tried).", code=2)


# ---------------------------------------------------------------------------
# Parallel worker and CLI entry point
# ---------------------------------------------------------------------------
def _expand_paths(patterns: list[str], recursive: bool) -> list[str]:
    paths = []
    for pattern in patterns:
        expanded = glob.glob(pattern, recursive=recursive)
        if expanded:
            paths.extend(expanded)
        else:
            paths.append(pattern)
    seen = set()
    result = []
    for p in paths:
        rp = os.path.realpath(p)
        if rp not in seen:
            seen.add(rp)
            result.append(p)
    return result


def _process_one(args_tuple):
    """Worker function for parallel execution."""
    path, cli_args, password, convert = args_tuple
    # Recreate logger for worker (avoid duplicate handlers)
    worker_log = logging.getLogger(f"unprotect.{os.getpid()}")
    worker_log.setLevel(logging.WARNING)
    try:
        if cli_args.check:
            result = check_file(path, password)
            return result
        elif cli_args.password_list:
            result = _try_password_list(
                input_path=path,
                wordlist_path=cli_args.password_list,
                output_path=cli_args.output,
                in_place=cli_args.in_place,
                output_dir=cli_args.output_dir,
                backup=cli_args.backup,
                no_overwrite=cli_args.no_overwrite,
                convert=convert,
            )
            return result
        else:
            result = unprotect_file(
                input_path=path,
                password=password,
                output_path=cli_args.output,
                in_place=cli_args.in_place,
                output_dir=cli_args.output_dir,
                backup=cli_args.backup,
                no_overwrite=cli_args.no_overwrite,
                convert=convert,
            )
            return result
    except UnprotectError as e:
        return FileResult(path=path, status="failed", message=str(e))
    except Exception as e:
        return FileResult(path=path, status="failed", message=f"Unexpected: {e}")


def main(argv: list[str] | None = None):
    parser = argparse.ArgumentParser(
        description="Remove password protection from PDF and Office files (now with legacy formats, conversion, parallel processing).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  unprotect.py report.xlsx --password secret
  unprotect.py "*.xls" --convert --output-dir ./unlocked/
  unprotect.py "**/*.doc" --recursive --in-place --backup --jobs 4
  unprotect.py report.xlsx --password-list words.txt --convert
  unprotect.py *.pdf --check --json
""",
    )
    parser.add_argument("files", nargs="+", help="Path(s) or glob pattern(s)")
    pw_group = parser.add_mutually_exclusive_group()
    pw_group.add_argument("--password", "-p", default=None)
    pw_group.add_argument("--password-file", metavar="FILE")
    pw_group.add_argument("--password-list", metavar="WORDLIST")
    out_group = parser.add_mutually_exclusive_group()
    out_group.add_argument("--output", "-o", default=None)
    out_group.add_argument("--in-place", action="store_true")
    out_group.add_argument("--output-dir", metavar="DIR")
    parser.add_argument("--backup", action="store_true")
    parser.add_argument("--no-overwrite", action="store_true")
    parser.add_argument("--check", "--dry-run", action="store_true")
    parser.add_argument("--json", dest="json_output", action="store_true")
    parser.add_argument("--recursive", "-r", action="store_true")
    parser.add_argument("--fail-fast", action="store_true")
    parser.add_argument("--convert", action="store_true", help="Convert legacy binary formats to OpenXML (requires LibreOffice)")
    parser.add_argument("--jobs", "-j", type=int, default=1, help="Number of parallel workers (default: 1)")
    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument("--verbose", "-v", action="store_true")
    verbosity.add_argument("--quiet", "-q", action="store_true")

    args = parser.parse_args(argv)

    if args.verbose:
        log.setLevel(logging.DEBUG)
    elif args.quiet:
        log.setLevel(logging.ERROR)
    else:
        log.setLevel(logging.INFO)

    if args.backup and not args.in_place:
        parser.error("--backup only makes sense with --in-place")
    if args.json_output and not args.check:
        parser.error("--json only with --check")
    if args.convert and args.check:
        parser.error("--convert cannot be used with --check (no conversion in dry-run)")

    password = args.password
    if args.password_file:
        try:
            password = _load_password_file(args.password_file)
        except (OSError, UnprotectError) as e:
            log.error("Error reading --password-file: %s", e)
            sys.exit(1)

    input_paths = _expand_paths(args.files, recursive=args.recursive)

    if args.output and len(input_paths) > 1:
        parser.error("--output only for single file")

    if not args.check and not args.password_list and password is None and sys.stdin.isatty() and input_paths:
        try:
            prompted = getpass.getpass("Password (leave blank if none): ")
            if prompted:
                password = prompted
        except (KeyboardInterrupt, EOFError):
            print()
            sys.exit(130)

    # Parallel processing
    if args.jobs > 1:
        log.debug("Using %d parallel workers", args.jobs)
        results = []
        with ProcessPoolExecutor(max_workers=args.jobs) as executor:
            futures = {
                executor.submit(_process_one, (path, args, password, args.convert)): path
                for path in input_paths
            }
            for future in as_completed(futures):
                try:
                    res = future.result()
                    results.append(res)
                except Exception as e:
                    results.append(FileResult(path=futures[future], status="failed", message=f"Error: {e}"))
                    if args.fail_fast:
                        for f in futures:
                            f.cancel()
                        break
        ordered = {res.path: res for res in results}
        results = [ordered[p] for p in input_paths if p in ordered]
    else:
        results = []
        for path in input_paths:
            try:
                if args.check:
                    res = check_file(path, password)
                elif args.password_list:
                    res = _try_password_list(
                        input_path=path,
                        wordlist_path=args.password_list,
                        output_path=args.output,
                        in_place=args.in_place,
                        output_dir=args.output_dir,
                        backup=args.backup,
                        no_overwrite=args.no_overwrite,
                        convert=args.convert,
                    )
                else:
                    res = unprotect_file(
                        input_path=path,
                        password=password,
                        output_path=args.output,
                        in_place=args.in_place,
                        output_dir=args.output_dir,
                        backup=args.backup,
                        no_overwrite=args.no_overwrite,
                        convert=args.convert,
                    )
                results.append(res)
            except UnprotectError as e:
                results.append(FileResult(path=path, status="failed", message=f"Error: {e}"))
                if args.fail_fast:
                    break
            except Exception as e:
                results.append(FileResult(path=path, status="failed", message=f"Unexpected: {e}"))
                if args.fail_fast:
                    break

    json_results = []
    counts = {"unlocked": 0, "open": 0, "skipped": 0, "failed": 0}
    exit_code = 0
    for res in results:
        if args.check:
            _print_check_result(res, args.json_output, json_results if args.json_output else None)
        else:
            if res.status == "skipped":
                log.info("Skipped: %s", res.message)
            else:
                log.info("%s", res.message)
        counts[res.status if res.status in counts else "failed"] += 1
        if res.status == "failed":
            exit_code = 1

    if args.json_output:
        print(json.dumps(json_results, indent=2))

    if len(input_paths) > 1 and not args.quiet and not args.json_output:
        parts = [f"{counts[k]} {k}" for k in ("unlocked", "open", "skipped", "failed") if counts[k]]
        log.info("\nDone: %s", ", ".join(parts) if parts else "nothing processed")

    sys.exit(exit_code)


if __name__ == "__main__":
    main()