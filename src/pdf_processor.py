import os
import re
import pdfplumber
from .logger import logger
import json
from .ocr_processor import OCRProcessor


class PDFProcessor:
    def __init__(self, config_path='config/config.json'):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        self.ocr_processor = OCRProcessor(config_path)



    def extract_text_from_pdf(self, pdf_path):
        """Extract text from a single PDF, using OCR if it's a scan."""
        try:
            if not os.path.exists(pdf_path):
                logger.error(f"PDF file not found: {pdf_path}")
                return ""

            if not os.access(pdf_path, os.R_OK):
                logger.error(f"Cannot read PDF file (permission denied): {pdf_path}")
                return ""

            with pdfplumber.open(pdf_path) as pdf:
                if not pdf.pages:
                    logger.error(f"No pages in PDF {pdf_path}")
                    return ""

                # Try to extract text with layout preservation
                try:
                    text = pdf.pages[0].extract_text(layout=True) or ""
                except Exception as e:
                    logger.warning(f"Layout extraction failed, trying simple extraction: {e}")
                    text = pdf.pages[0].extract_text() or ""

                # Check if the extracted text is meaningful (not a scan)
                if self._is_scan(text):
                    logger.info(f"PDF {pdf_path} appears to be a scan, using OCR")
                    # Use OCR for scanned PDFs
                    text = self.ocr_processor.extract_text_from_scan(pdf_path)
                    if not text:
                        return "SCAN"  # Special marker for scans that couldn't be processed
                    return text
                else:
                    return text

        except PermissionError as e:
            logger.error(f"Permission denied accessing {pdf_path}: {e}")
            return ""
        except Exception as e:
            logger.error(f"Error extracting text from {pdf_path}: {e}")
            return ""

    def _is_scan(self, text):
        """Determine if the PDF is a scanned image based on extracted text."""
        if not text or text.strip() == "":
            return True

        # Count meaningful words (words longer than 2 characters)
        words = re.findall(r'\b\w{3,}\b', text)
        if len(words) < 10:  # Very few words suggest it's a scan
            return True

        # Check for common PDF text patterns
        # If text contains mostly symbols or very short fragments, it's likely a scan
        lines = text.split('\n')
        meaningful_lines = 0
        for line in lines:
            line = line.strip()
            if len(line) > 5 and re.search(r'[a-zA-Z]{3,}', line):  # At least 3 letters
                meaningful_lines += 1

        # If less than 30% of lines are meaningful, consider it a scan
        if meaningful_lines / len(lines) < 0.3 if lines else True:
            return True

        return False

    def process_pdfs_concurrent(self, pdf_paths, callback=None, save_ocr_text=False, ocr_text_dir=None):
        """Process multiple PDFs concurrently with better error handling."""
        from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError

        max_workers = min(self.config['settings']['max_workers'], len(pdf_paths))
        results = {}

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_path = {
                executor.submit(self.extract_text_from_pdf, path): path
                for path in pdf_paths
            }

            for future in as_completed(future_to_path):
                path = future_to_path[future]
                filename = os.path.basename(path)

                try:
                    # Get result with timeout
                    text = future.result(timeout=180)  # 3 minutes per file
                    results[path] = text

                    if callback:
                        callback(filename, text)

                except TimeoutError:
                    logger.error(f"Processing timeout for {filename}")
                    results[path] = ""
                    if callback:
                        callback(filename, "")

                except Exception as e:
                    logger.error(f"Failed to process {filename}: {e}")
                    results[path] = ""
                    if callback:
                        callback(filename, "")

        return results