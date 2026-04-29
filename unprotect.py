#!/usr/bin/env python3
import sys
import os
import argparse
import shutil
import zipfile

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
            sys.exit(1)
        try:
            office_file.load_key(password=password)
            with open(tmp_path, "wb") as out:
                office_file.decrypt(out)
        except Exception as e:
            print(f"Error: Couldn't decrypt file - {e}")
            sys.exit(1)
    return True


def _cleanup(path: str):
    if path and os.path.exists(path):
        os.remove(path)


def _rewrite_zip(zip_path: str, filename_in_zip: str, new_content: bytes):
    """Replace a single file inside a zip archive in-place
    preserving compression type and metadata fopr every other entry"""
    tmp_zip = zip_path + ".zip.tmp"
    with zipfile.ZipFile(zip_path, "r") as zin, \
        zipfile.ZipFile(tmp_zip, "w") as zout:
        for item in zin.infolist():
            if item.filename == filename_in_zip:
                # Write replacement with same compression as original
                zout.writestr(item, new_content, compress_type=item.compress_type)
            else:
                # Copy verbatim, preserving all metadata
                zout.writestr(item, zin.read(item.filename), compress_type=item.compress_type)
    os.replace(tmp_zip, zip_path)

# PDF Functions

def unprotect_pdf(input_path: str, password: str, output_path: str) -> bool:
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        print("Missing dependency! Please run: pip install pypdf")
        return False
    
    reader = PdfReader(input_path)

    if reader.is_encrypted:
        if not password:
            print("Error: PDF is encrypted but no password was given.")
            return False
        result = reader.decrypt(password)
        if result == 0:
            print("Error: Wrong password!")
            return False

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"PDF has been unprotected: {output_path}")
    return True

    # Excel

def unprotect_excel(input_path: str, password: str, output_path: str) -> bool:
    tmp_path = output_path + ".tmp.xlsx"
    was_encrypted = _msoffcrypto_decrypt(input_path, password, tmp_path)
    work_path = tmp_path if was_encrypted else input_path

    try:
        from openpyxl import load_workbook
        wb = load_workbook(work_path)

        # Remove workbook-level protection
        if wb.security and wb.security.workbookPassword:
            wb.security.workbookPassword = None
            wb.security.lockStructure = False
            wb.security.lockWindows = False

        # Remove sheet-level protection from every sheet
        for sheet in wb.worksheets:
            if sheet.protection.sheet:
                sheet.protection.sheet = False
                sheet.protection.password = None

        wb.save(output_path)
    except ImportError:
        print("Missing dependency! Please run: pip install openpyxl")
        _cleanup(tmp_path)
        return False
    except Exception as e:
        print(f"Error removing sheet protection: {e}")
        _cleanup(tmp_path)
        return False

    _cleanup(tmp_path)
    print(f"Excel file has been unprotected: {output_path}")
    return True

# Word
def unprotect_word(input_path: str, password: str, output_path: str) -> bool:
    tmp_path = output_path + ".tmp.docx"
    was_encrypted = _msoffcrypto_decrypt(input_path, password, tmp_path)
    work_path = tmp_path if was_encrypted else input_path

    shutil.copy2(work_path, output_path)
    _cleanup(tmp_path)

    try:
        import lxml.etree as etree # Pyright has an aneurysm on etree or else i would have used "from lxml import etree"

        with zipfile.ZipFile(output_path, "r") as z:
            settings_name = next((n for n in z.namelist() if n.endswith("settings.xml")), None)
            if settings_name is None:
                print(f"Word file is unprotected (no settings have been found): {output_path}")
                return True
            settings_xml = z.read(settings_name)

        root = etree.fromstring(settings_xml)
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        for tag in ["documentProtection", "writeProtection"]:
            for el in root.findall(f"{{{ns}}}{tag}"):
                root.remove(el)

        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF8", standalone=True)
        _rewrite_zip(output_path, settings_name, new_xml)

    except ImportError:
        print("Missing dependency! Please run: pip install lxml")
        return False
    except Exception as e:
        print(f"Error trying to remove Word protection: {e}")
        return False

    print(f"Word file has been unprotected: {output_path}")
    return True

# Powerpoint

def unprotect_powerpoint(input_path: str, password: str, output_path: str) -> bool:
    tmp_path = output_path + ".tmp.pptx"
    was_encrypted = _msoffcrypto_decrypt(input_path, password, tmp_path)
    work_path = tmp_path if was_encrypted else input_path
 
    shutil.copy2(work_path, output_path)
    _cleanup(tmp_path)
 
    try:
        import lxml.etree as etree
 
        with zipfile.ZipFile(output_path, "r") as z:
            prs_name = next((n for n in z.namelist() if n.endswith("presentation.xml")), None)
            if prs_name is None:
                print(f"✓ PowerPoint file unprotected (no presentation.xml): {output_path}")
                return True
            prs_xml = z.read(prs_name)
 
        root = etree.fromstring(prs_xml)
        ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
 
        for tag in ["modifyVerifier", "writeProtection"]:
            for el in root.findall(f"{{{ns}}}{tag}"):
                root.remove(el)
 
        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        _rewrite_zip(output_path, prs_name, new_xml)
 
    except ImportError:
        print("Missing dependency. Run: pip install lxml")
        return False
    except Exception as e:
        print(f"Error removing PowerPoint protection: {e}")
        return False
 
    print(f"✓ PowerPoint file unprotected: {output_path}")
    return True

# Main

SUPPORTED = {
    ".pdf":  ("PDF",        unprotect_pdf),
    ".xlsx": ("Excel",      unprotect_excel),
    ".xlsm": ("Excel",      unprotect_excel),
    ".xls":  ("Excel",      unprotect_excel),
    ".docx": ("Word",       unprotect_word),
    ".doc":  ("Word",       unprotect_word),
    ".pptx": ("PowerPoint", unprotect_powerpoint),
    ".ppt":  ("PowerPoint", unprotect_powerpoint),
}

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