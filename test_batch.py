#!/usr/bin/env python3
import os
import sys
from collections import Counter

# Ensure we can import the src package
sys.path.append('.')

from src.pdf_processor import PDFProcessor
from src.data_extractor import DataExtractor
from src.data_writer import DataWriter

def main():
    input_dir = "Cupones"
    output_path = os.path.join(input_dir, "resultado.xlsx")

    pdfs = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    pdfs.sort()
    if not pdfs:
        print("No PDFs found in Cupones/")
        return

    processor = PDFProcessor()
    extractor = DataExtractor()
    writer = DataWriter()

    results = []
    field_missing_counts = Counter()
    per_file_missing = {}

    print(f"Processing {len(pdfs)} PDFs...")
    for i, pdf in enumerate(pdfs, 1):
        try:
            text = processor.extract_text_from_pdf(pdf)
            data = extractor.extract_data(text or "", os.path.basename(pdf))
            results.append(data)

            # Count "No encontrado" fields
            missing_fields = [k for k, v in data.items() if v == "No encontrado"]
            for mf in missing_fields:
                field_missing_counts[mf] += 1
            if missing_fields:
                per_file_missing[os.path.basename(pdf)] = missing_fields

            print(f"[{i}/{len(pdfs)}] {os.path.basename(pdf)} -> missing: {', '.join(missing_fields) if missing_fields else 'none'}")
        except Exception as e:
            print(f"[{i}/{len(pdfs)}] ERROR processing {os.path.basename(pdf)}: {e}")

    # Write output
    total = writer.write_data(results, output_path)
    print(f"\nWrote {total} records to {output_path}")

    # Summary of missing fields
    if field_missing_counts:
        print("\nFields still missing (counts):")
        for field, cnt in field_missing_counts.items():
            print(f"  {field}: {cnt}")
        # Show up to 10 example files with missing fields
        print("\nSample files with missing fields:")
        shown = 0
        for fname, mlist in per_file_missing.items():
            print(f"  {fname}: {', '.join(mlist)}")
            shown += 1
            if shown >= 10:
                break
    else:
        print("\nAll fields extracted for all files.")

if __name__ == "__main__":
    main()