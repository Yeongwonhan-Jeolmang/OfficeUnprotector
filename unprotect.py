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
    """Replace a single file inside a zip archive in-place"""
    tmp_zip = zip_path + ".zip.tmp"
    with zipfile.ZipFile(zip_path, "r") as zin, \
        zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_content if item.filename == filename_in_zip else zin.read(item.filename)
            zout.writestr(item, data)
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