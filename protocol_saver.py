#!/usr/bin/env python3
"""Python port of protocol_saver.vba (ConvertPDFsToDocx).

Converts every PDF in a folder to DOCX (same name, same folder) using pdf2docx.
Layout reconstruction differs from Word's PDF reflow, so results may not be identical.
"""
import os
import argparse
from pdf2docx import Converter


def convert_pdfs_to_docx(directory_path):
    for file in sorted(os.listdir(directory_path)):
        if not file.lower().endswith(".pdf"):
            continue
        pdf_path = os.path.join(directory_path, file)
        docx_path = os.path.join(directory_path, file[:-4] + ".docx")
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
        except Exception as e:
            print(f"Failed to convert {pdf_path}: {e}")

    print("PDF to DOCX conversion complete.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert all PDF files in a folder to DOCX.")
    parser.add_argument("input_dir", help="Directory containing PDF files")
    args = parser.parse_args()
    convert_pdfs_to_docx(args.input_dir)
