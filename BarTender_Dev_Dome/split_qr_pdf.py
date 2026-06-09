#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Split PDF pages containing multiple QR code labels into individual PDFs.
Each label is detected by analyzing text block positions.
Output format: page_{global_index}.pdf
"""

import os
import sys

import fitz  # PyMuPDF


def detect_label_count(page):
    """Detect how many QR labels are on this page by analyzing text blocks."""
    blocks = page.get_text("blocks")
    if not blocks:
        return 0

    label_indices = set()
    for block in blocks:
        y0 = block[1]
        label_idx = int(y0 // 172)
        label_indices.add(label_idx)

    return len(label_indices)


def split_pdf_by_labels(input_path, output_dir):
    """Split PDF so each QR label becomes its own PDF file."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    doc = fitz.open(input_path)
    output_files = []
    label_counter = 0

    for page_num in range(len(doc)):
        page = doc[page_num]
        label_count = detect_label_count(page)

        for i in range(label_count):
            rect = fitz.Rect(0, i * 172, 99, i * 172 + 170)

            new_doc = fitz.open()
            new_page = new_doc.new_page(width=rect.width, height=rect.height)
            new_page.show_pdf_page(new_page.rect, doc, page_num, clip=rect)

            output_path = os.path.join(output_dir, f"page_{label_counter + 1}.pdf")
            new_doc.save(output_path)
            new_doc.close()

            output_files.append(output_path)
            label_counter += 1

    doc.close()
    return output_files


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python split_qr_pdf.py <input_pdf> <output_dir>")
        sys.exit(1)

    input_pdf = sys.argv[1]
    output_dir = sys.argv[2]

    if not os.path.exists(input_pdf):
        print(f"ERROR: Input file not found: {input_pdf}")
        sys.exit(1)

    files = split_pdf_by_labels(input_pdf, output_dir)
    print(f"Split into {len(files)} individual QR PDFs:")
    for file_path in files:
        print(f"  {file_path}")

