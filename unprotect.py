#!/usr/bin/env python3
import sys
import os
import argparse
import glob
import shutil
import zipfile
from typing import Callable

# Helper functions

def _msoffcrypto_decrypt(input_path: str, password: str, tmp_path: str) -> bool:
    """Decrypt an encrypted Office file using msoffcrypto-tool.
    Returns True if the file was encrypted and successfully decrypted."""
    try:
        import msoffcrypto
    except ImportError:
        print("Missing dependency! Please run: pip install msoffcrypto-tool")
        sys.exit(1)
    
    with open(input_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        if not office_file.is_encrypted():
            return False # not encrypted so skip the decryption step
        if not password:
            print("Error: File is encrypted but no password was provided!")
            sys.exit(2)
        try:
            office_file.load_key(password=password)
            with open(tmp_path, "wb") as out:
                office_file.decrypt(out)
        except Exception as e:
            print(f"Error: Couldn't decrypt file - {e}")
            sys.exit(2)
    return True


def _cleanup(path: str):
    if path and os.path.exists(path):
        os.remove(path)


def _rewrite_zip(zip_path: str, filename_in_zip: str, new_content: bytes):
    """Replace a single file inside a zip archive in-place
    preserving compression type and metadata fopr every other entry"""
    tmp_zip = zip_path + ".zip.tmp"
    with zipfile.ZipFile(zip_path, "r") as zin, \
        zipfile.ZipFile(tmp_zip, "w") as zout: # No forced ZIP_DEFLATED
        for item in zin.infolist():
            if item.filename == filename_in_zip:
                # Write replacement with same compression as original
                zout.writestr(item, new_content, compress_type=item.compress_type)
            else:
                # Copy verbatim, preserving all metadata++
                zout.writestr(item, zin.read(item.filename), compress_type=item.compress_type)
    os.replace(tmp_zip, zip_path)

def _resolve_output(input_path: str, output_arg: str | None, in_place: bool) -> str:

    if in_place:
        return input_path
    if output_arg:
        return output_arg
    base = os.path.basename(input_path)
    name, ext = os.path.splitext(base)
    directory = os.path.dirname(input_path) or "."
    return os.path.join(directory, f"unlocked_{name}{ext}")

def _check_collision(input_path: str, output_path: str, in_place: bool):
    if in_place:
        return # intentional overwrite
    if os.path.realpath(input_path) == os.path.realpath(output_path):
        print("Error: Input and output paths resolve to the same file. Use --in-place to overwrite, or choose a different --output path.")
    sys.exit(1)

# --check / dry-run functions

def check_protection(input_path: str, password: str | None) -> None:
    """Report whether a file is protected without modifying it."""
    ext = os.path.splitext(input_path)[1].lower()

    # Encryption check (all Office formats)
    if ext in (".xlsx", ".xlsm", ".docx", ".pptx"):
        try:
            import msoffcrypto
            with open(input_path, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                encrypted = of.is_encrypted()
        except ImportError:
            encrypted = False
        if encrypted:
            print(f"[ENCRYPTED]  {input_path}  (file-level password required to open)")
        else:
            _check_xml_protection(input_path, ext)

    elif ext == ".pdf":
        try:
            from pypdf import PdfReader
            r = PdfReader(input_path)
            if r.is_encrypted:
                print(f"[ENCRYPTED]  {input_path}  (PDF open password set)")
            else:
                print(f"[OPEN]       {input_path}  (no PDF encryption detected)")
        except ImportError:
            print("Missing dependency! Please run: pip install pypdf")
    else:
        print(f"[UNKNOWN]    {input_path}  (unsupported extension '{ext}')")


def _check_xml_protection(input_path: str, ext: str) -> None:
    """Inspect the XML inside an (unencrypted) Office ZIP for protection tags."""
    protected_items: list[str] = []
    try:
        with zipfile.ZipFile(input_path, "r") as z:
            names = z.namelist()

            if ext in (".xlsx", ".xlsm"):
                # Workbook-level
                wb_name = next((n for n in names if n.endswith("workbook.xml")), None)
                if wb_name:
                    import lxml.etree as etree  # type: ignore[import]
                    root = etree.fromstring(z.read(wb_name))
                    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    if root.find(f"{{{ns}}}workbookProtection") is not None:
                        protected_items.append("workbook structure")
                # Per-sheet
                for n in names:
                    if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                        import lxml.etree as etree  # type: ignore[import]
                        root = etree.fromstring(z.read(n))
                        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                        if root.find(f"{{{ns}}}sheetProtection") is not None:
                            protected_items.append(n)

            elif ext == ".docx":
                settings_name = next((n for n in names if n.endswith("settings.xml")), None)
                if settings_name:
                    import lxml.etree as etree  # type: ignore[import]
                    root = etree.fromstring(z.read(settings_name))
                    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    for tag in ("documentProtection", "writeProtection"):
                        if root.find(f"{{{ns}}}{tag}") is not None:
                            protected_items.append(tag)

            elif ext == ".pptx":
                prs_name = next((n for n in names if n.endswith("presentation.xml")), None)
                if prs_name:
                    import lxml.etree as etree  # type: ignore[import]
                    root = etree.fromstring(z.read(prs_name))
                    ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
                    for tag in ("modifyVerifier", "writeProtection"):
                        if root.find(f"{{{ns}}}{tag}") is not None:
                            protected_items.append(tag)

    except Exception as e:
        print(f"[ERROR]      {input_path}  — could not inspect: {e}")
        return

    if protected_items:
        print(f"[PROTECTED]  {input_path}  — {', '.join(protected_items)}")
    else:
        print(f"[OPEN]       {input_path}  — no protection detected")

# PDF Functions

def unprotect_pdf(input_path: str, password: str | None, output_path: str) -> int:
    """Returns 0 on success, 1 on error, 2 on wrong password"""
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        print("Missing dependency! Please run: pip install pypdf")
        return 1
    
    reader = PdfReader(input_path)

    if reader.is_encrypted:
        if not password:
            print("Error: PDF is encrypted but no password was given.")
            return False
        result = reader.decrypt(password)
        if result == 0:
            print("Error: Wrong password!")
            return 2

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"PDF has been unprotected: {output_path}")
    return 0

    # Excel

def unprotect_excel(input_path: str, password: str | None, output_path: str) -> int:
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".xls":
        print("Error: Legacy .xls format is not supported. The binary BIFF format requires a separate tool (e.g. LibreOffice). Convert to .xlsx first, then retry.")
        return 1

    tmp_path = output_path + ".tmp.xlsx"
    try:
        was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_path)
        work_path = tmp_path if was_encrypted else input_path

        # Use direct XML manipulation (more reliable than openpyxl for hashed passwords)
        shutil.copy2(work_path, output_path)
        _strip_excel_xml_protection(output_path)

    except SystemExit:
        raise
    except Exception as e:
        print(f"Error removing Excel protection: {e}")
        return 1
    finally:
        _cleanup(tmp_path)

    print(f"Excel file unprotected: {output_path}")
    return 0

def _strip_excel_xml_protection(xlsx_path: str):
    """Remove workbook-level and per-sheet protection by editing XML directly.
    This handles modern hash-based protection (hashValue/saltValue) that
    openpyxl's high-level API may leave behind."""
    import lxml.etree as etree  # type: ignore[import]

    with zipfile.ZipFile(xlsx_path, "r") as z:
        names = z.namelist()

        # Workbook protection
        wb_name = next((n for n in names if n.endswith("workbook.xml")), None)
        if wb_name:
            wb_xml = z.read(wb_name)
            wb_root = etree.fromstring(wb_xml)
            ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            for el in wb_root.findall(f"{{{ns}}}workbookProtection"):
                wb_root.remove(el)
            new_wb_xml = etree.tostring(wb_root, xml_declaration=True, encoding="UTF-8", standalone=True)
            _rewrite_zip(xlsx_path, wb_name, new_wb_xml)

        # Per-sheet protection
        sheet_names = [n for n in names
                       if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]

    for sheet_name in sheet_names:
        with zipfile.ZipFile(xlsx_path, "r") as z:
            sheet_xml = z.read(sheet_name)
        import lxml.etree as etree  # type: ignore[import]
        root = etree.fromstring(sheet_xml)
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        changed = False
        for el in root.findall(f"{{{ns}}}sheetProtection"):
            root.remove(el)
            changed = True
        if changed:
            new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            _rewrite_zip(xlsx_path, sheet_name, new_xml)

# Word
def unprotect_word(input_path: str, password: str | None, output_path: str) -> int:
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".doc":
        print("Error: Legacy .doc format is not supported. The binary format requires a separate tool (e.g. LibreOffice). Convert to .docx first, then retry.")
        return 1

    tmp_path = output_path + ".tmp.docx"
    try:
        was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_path)
        work_path = tmp_path if was_encrypted else input_path

        shutil.copy2(work_path, output_path)

        import lxml.etree as etree  # type: ignore[import]

        with zipfile.ZipFile(output_path, "r") as z:
            settings_name = next((n for n in z.namelist() if n.endswith("settings.xml")), None)
            if settings_name is None:
                print(f"✓ Word file unprotected (no settings.xml found): {output_path}")
                return 0
            settings_xml = z.read(settings_name)

        root = etree.fromstring(settings_xml)
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        changed = False
        for tag in ("documentProtection", "writeProtection"):
            for el in root.findall(f"{{{ns}}}{tag}"):
                root.remove(el)
                changed = True

        if changed:
            new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            _rewrite_zip(output_path, settings_name, new_xml)

    except ImportError:
        print("Missing dependency! Please run: pip install lxml")
        return 1
    except SystemExit:
        raise
    except Exception as e:
        print(f"Error removing Word protection: {e}")
        return 1
    finally:
        _cleanup(tmp_path)

    print(f"Word file unprotected: {output_path}")
    return 0

# Powerpoint

def unprotect_powerpoint(input_path: str, password: str | None, output_path: str) -> int:
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".ppt":
        print(
            "Error: Legacy .ppt format is not supported. "
            "The binary format requires a separate tool (e.g. LibreOffice). "
            "Convert to .pptx first, then retry."
        )
        return 1

    tmp_path = output_path + ".tmp.pptx"
    try:
        was_encrypted = _msoffcrypto_decrypt(input_path, password or "", tmp_path)
        work_path = tmp_path if was_encrypted else input_path

        shutil.copy2(work_path, output_path)

        import lxml.etree as etree  # type: ignore[import]

        # Presentation-level protection
        with zipfile.ZipFile(output_path, "r") as z:
            names = z.namelist()
            prs_name = next((n for n in names if n.endswith("presentation.xml")), None)
            if prs_name is None:
                print(f"✓ PowerPoint unprotected (no presentation.xml): {output_path}")
                return 0
            prs_xml = z.read(prs_name)

        root = etree.fromstring(prs_xml)
        ns_pml = "http://schemas.openxmlformats.org/presentationml/2006/main"
        changed = False
        for tag in ("modifyVerifier", "writeProtection"):
            for el in root.findall(f"{{{ns_pml}}}{tag}"):
                root.remove(el)
                changed = True

        if changed:
            new_xml = etree.tostring(root, xml_declaration=True,
                                     encoding="UTF-8", standalone=True)
            _rewrite_zip(output_path, prs_name, new_xml)

        # Per-slide protection (oleObj / AlternateContent locks)
        with zipfile.ZipFile(output_path, "r") as z:
            slide_names = [n for n in z.namelist()
                           if n.startswith("ppt/slides/slide") and n.endswith(".xml")]

        for slide_name in slide_names:
            with zipfile.ZipFile(output_path, "r") as z:
                slide_xml = z.read(slide_name)
            slide_root = etree.fromstring(slide_xml)

            # Remove mc:AlternateContent blocks that wrap locked OLE objects
            mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
            slide_changed = False
            for ac in slide_root.findall(f".//{{{mc_ns}}}AlternateContent"):
                parent = ac.getparent()
                if parent is not None:
                    # Only remove if it wraps a locked oleObj
                    if b"oleObj" in etree.tostring(ac) and b"locked" in etree.tostring(ac):
                        parent.remove(ac)
                        slide_changed = True

            if slide_changed:
                new_slide_xml = etree.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone=True)
                _rewrite_zip(output_path, slide_name, new_slide_xml)

    except ImportError:
        print("Missing dependency! Please run: pip install lxml")
        return 1
    except SystemExit:
        raise
    except Exception as e:
        print(f"Error removing PowerPoint protection: {e}")
        return 1
    finally:
        _cleanup(tmp_path)

    print(f"PowerPoint unprotected: {output_path}")
    return 0

# Dispatch table

# Maps extension -> (label, handler)
# .xls / .doc / .ppt are actually intentionally kept so they get a clear error message :D
# rather than an "unsupported file type" error.
SUPPORTED: dict[str, tuple[str, Callable]] = {
    ".pdf":  ("PDF",        unprotect_pdf),
    ".xlsx": ("Excel",      unprotect_excel),
    ".xlsm": ("Excel",      unprotect_excel),
    ".xls":  ("Excel",      unprotect_excel),   # will self-reject with a VERY helpful message (trust)
    ".docx": ("Word",       unprotect_word),
    ".doc":  ("Word",       unprotect_word),    # will self-reject
    ".pptx": ("PowerPoint", unprotect_powerpoint),
    ".ppt":  ("PowerPoint", unprotect_powerpoint),  # will self-reject
}

# Main

def process_file(input_path: str, password: str | None,
                 output_arg: str | None, in_place: bool,
                 check_only: bool) -> int:
    """Process a single file. Returns an exit code."""

    if not os.path.exists(input_path):
        print(f"Error: File not found: {input_path}")
        return 4

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED:
        supported_list = ", ".join(SUPPORTED.keys())
        print(f"Error: Unsupported file type '{ext}'. Supported: {supported_list}")
        return 3

    if check_only:
        check_protection(input_path, password)
        return 0

    label, handler = SUPPORTED[ext]
    output_path = _resolve_output(input_path, output_arg, in_place)
    _check_collision(input_path, output_path, in_place)

    print(f"Processing {label} file: {input_path}")
    return handler(input_path, password, output_path)

def main():
    parser = argparse.ArgumentParser(description="Remove password protection from PDF and Office 365 files.")
    parser.add_argument("file", help="Path to the protected file")
    parser.add_argument("password", nargs="?", default=None, help="Password to unlock the file (omit if no open password)")
    parser.add_argument("--output", "-o", default=None, help="Output file path (default: unlocked_<filename>)")
    args = parser.parse_args()

    input_path = args.file
    password = args.password

    if not os.path.exists(input_path):
        print(f"Error: File not found: {input_path}")
        sys.exit(1)
    
    ext = os.path.splitext(input_path)[1].lower()

    if ext not in SUPPORTED:
        supported_list = ", ".join(SUPPORTED.keys())
        print(f"Error: Unsupported file type '{ext}'. Supported: {supported_list}")
        sys.exit(1)

    label, handler = SUPPORTED[ext]
    base = os.path.basename(input_path)
    name, suffix = os.path.splitext(base)

    output_path = args.output or os.path.join(os.path.dirname(input_path) or ".", f"unlocked_{name}{suffix}")

    print(f"Processing {label} file: {input_path}")
    success = handler(input_path, password, output_path)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()