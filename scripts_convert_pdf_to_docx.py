#!/usr/bin/env python3
# Usage: python3 convert_pdf_to_docx.py /path/to/input.pdf /path/to/output_dir
# Requires: ocrmypdf, pdf2docx installed and tesseract with Persian (fas)

import sys
import os
import subprocess
from pdf2docx import Converter

def run_cmd(cmd):
    print("RUN:", " ".join(cmd))
    subprocess.run(cmd, check=True)

def ocr_pdf(input_pdf, out_pdf):
    # Best-effort OCR settings for Persian
    cmd = [
        "ocrmypdf",
        "-l", "fas",
        "--deskew",
        "--image-dpi", "300",
        "--jobs", "2",
        "--optimize", "1",
        input_pdf, out_pdf
    ]
    run_cmd(cmd)

def pdf_to_docx(searchable_pdf, out_docx):
    cv = Converter(searchable_pdf)
    try:
        cv.convert(out_docx)  # convert whole document
    finally:
        cv.close()

def fallback_libreoffice(searchable_pdf, out_docx):
    # If pdf2docx fails, try LibreOffice conversion (may be better/worse depending on PDF)
    cmd = ["soffice", "--headless", "--convert-to", "docx", "--outdir", os.path.dirname(out_docx), searchable_pdf]
    run_cmd(cmd)
    # soffice names output with same base name; move if necessary
    base = os.path.splitext(os.path.basename(searchable_pdf))[0] + ".docx"
    produced = os.path.join(os.path.dirname(out_docx), base)
    if os.path.exists(produced):
        if produced != out_docx:
            os.replace(produced, out_docx)

def main():
    if len(sys.argv) < 3:
        print("Usage: convert_pdf_to_docx.py input.pdf output_dir")
        sys.exit(2)
    input_pdf = sys.argv[1]
    out_dir = sys.argv[2]
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_pdf))[0]
    searchable = os.path.join(out_dir, base + "_searchable.pdf")
    out_docx = os.path.join(out_dir, base + ".docx")

    try:
        print("Step 1: OCR PDF -> searchable PDF")
        ocr_pdf(input_pdf, searchable)
    except subprocess.CalledProcessError as e:
        print("OCR failed:", e)
        sys.exit(1)

    try:
        print("Step 2: Convert searchable PDF -> DOCX (pdf2docx)")
        pdf_to_docx(searchable, out_docx)
    except Exception as e:
        print("pdf2docx failed:", e)
        print("Trying LibreOffice conversion as fallback...")
        try:
            fallback_libreoffice(searchable, out_docx)
        except Exception as e2:
            print("LibreOffice fallback failed:", e2)
            sys.exit(1)

    print("Done. Output DOCX:", out_docx)

if __name__ == "__main__":
    main()