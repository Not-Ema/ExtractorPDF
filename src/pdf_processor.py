import os
import re
import pdfplumber
from .logger import logger
import json


class PDFProcessor:
    def __init__(self, config_path='config/config.json'):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)



    def extract_text_from_pdf(self, pdf_path):
        """Extract text from a single PDF."""
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
                

                return text

        except PermissionError as e:
            logger.error(f"Permission denied accessing {pdf_path}: {e}")
            return ""
        except Exception as e:
            logger.error(f"Error extracting text from {pdf_path}: {e}")
            return ""

    def process_pdfs_concurrent(self, pdf_paths, callback=None):
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