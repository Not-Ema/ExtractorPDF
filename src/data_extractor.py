import re
import json
from .logger import logger
from .ocr_processor import OCRProcessor

class DataExtractor:
    def __init__(self, config_path='config/config.json'):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        self.patterns = self.config['extraction_patterns']
        self.fields = self.config['fields']
        self.ocr_processor = OCRProcessor(config_path)

    def extract_data(self, text, filename):
        """Extract data from text using configured patterns."""
        data = {field: 'No encontrado' for field in self.fields}
        data['Archivo Origen'] = filename

        try:
            # Check if this is a scan that couldn't be processed
            if text == "SCAN":
                logger.warning(f"File {filename} is a scan that couldn't be processed with OCR")
                data['error'] = 'Archivo es un scan y no se pudo procesar con OCR'
                return data

            # Check if text looks like OCR output (has OCR artifacts)
            if self._is_ocr_text(text):
                logger.info(f"Using OCR-specific extraction for {filename}")
                # Use OCR extraction logic directly for scanned PDFs
                ocr_data = self.ocr_processor.extract_fields_from_ocr_text(text)
                # Map OCR field names to standard field names
                field_mapping = {
                    'Cliente': 'Cliente',
                    'Identificacion': 'Identificación',
                    'Contrato': 'Contrato',
                    'DirCliente': 'Dirección',
                    'NoSolicitud': 'No. Solicitud',
                    'NoRefPago': 'No. Rel. Pago',
                    'TipoCupon': 'Tipo de Cupón',
                    'ValidoHasta': 'Valido hasta',
                    'ValorAPagar': 'Valor a Pagar',
                    'CodigoBarraLimpio': 'Codigo de Barras Limpio',
                    'CodigoBarraRaw': None  # This field is not in the standard output
                }
                for ocr_field, std_field in field_mapping.items():
                    if ocr_data.get(ocr_field) is not None and ocr_data.get(ocr_field) != '':
                        data[std_field] = ocr_data[ocr_field]
                return data
            else:
                # Use standard extraction for digital PDFs
                text = self._normalize_text(text)
                extracted = self._extract_standard(text)
                data.update(extracted)
                return data

            # Additional fallbacks for common edge cases
            if data.get('Contrato') == 'No encontrado':
                m = re.search(r'Contrato\s*([0-9]{3,})', text, re.IGNORECASE)
                if m:
                    data['Contrato'] = m.group(1)

            if data.get('No. Solicitud') == 'No encontrado':
                m = re.search(r'No\W*Solic(?:itud)?\W*:?\s*([0-9]{6,})', text, re.IGNORECASE | re.DOTALL)
                if m:
                    data['No. Solicitud'] = m.group(1)

            if data.get('No. Rel. Pago') == 'No encontrado':
                m = re.search(r'No\W*Ref\W*\.?\W*(?:Pago)?\W*:?\s*([0-9]{6,})', text, re.IGNORECASE | re.DOTALL)
                if m:
                    data['No. Rel. Pago'] = m.group(1)

            if data.get('Dirección') == 'No encontrado':
                m = re.search(r'Dir\W*\.?\W*Cliente\W*:\s*(.+?)(?:\s+(?:FAX|L[íi]nea|Valor|Tipo|No\.)|$)', text, re.IGNORECASE | re.DOTALL)
                if m:
                    data['Dirección'] = re.sub(r'\s+', ' ', m.group(1).strip())

            if data.get('Valor a Pagar') == 'No encontrado':
                m = re.search(r'Valor\W*a\W*pagar\W*:\s*\$?([\d\.,]+)', text, re.IGNORECASE)
                if m:
                    data['Valor a Pagar'] = m.group(1)

            # Fallbacks and enhancements
            if data.get('Tipo de Cupón', 'No encontrado') == 'No encontrado':
                tipo = self._search_first(['tipo_cupon', 'tipo_cupon_alt1', 'tipo_cupon_alt2'], text, flags=re.DOTALL)
                if tipo != 'No encontrado':
                    data['Tipo de Cupón'] = tipo

            # Fallback for Cliente if still missing
            if data.get('Cliente', 'No encontrado') == 'No encontrado':
                m = re.search(r'Cliente[:\s]*([A-ZÑÁÉÍÓÚ\-\.\s]{3,}?)\s+Identificaci[óo]n', text, re.IGNORECASE | re.DOTALL)
                if m:
                    data['Cliente'] = re.sub(r'\s+', ' ', m.group(1).strip())

            # Fallback for Identificación if missing (search in full text)
            if data.get('Identificación', 'No encontrado') == 'No encontrado':
                m = re.search(self.patterns.get('identificacion', r'Identificaci\W*n\s*:?\s*(\d{6,})'), text, re.IGNORECASE | re.DOTALL)
                if m:
                    data['Identificación'] = m.group(1)

            # Fallback for No. Rel. Pago from barcode (8020) group
            if data.get('No. Rel. Pago', 'No encontrado') == 'No encontrado':
                m = re.search(r'\(8020\)\s*0*?(\d{6,})', text, re.IGNORECASE)
                if m:
                    data['No. Rel. Pago'] = m.group(1)

            # Fallback for No. Solicitud with tolerant pattern
            if data.get('No. Solicitud', 'No encontrado') == 'No encontrado':
                m = re.search(r'No\W*Solici(?:tud)?\W*:?\s*([0-9]{6,})', text, re.IGNORECASE | re.DOTALL)
                if m:
                    data['No. Solicitud'] = m.group(1)

            # If standard extraction failed for some fields, try coupon parsing
            if any(data[field] == 'No encontrado' for field in ['Cliente', 'Identificación', 'Dirección']):
                coupon_data = self._extract_coupon_data(text)
                for key, value in coupon_data.items():
                    if data.get(key, 'No encontrado') == 'No encontrado' and value:
                        data[key] = value

            logger.info(f"Data extracted successfully for {filename}")
        except Exception as e:
            logger.error(f"Error extracting data from {filename}: {e}")

        return data

    def _extract_standard(self, text):
        """Standard extraction using regex patterns."""
        data = {}
        # Extract blocks
        bloque_cliente_match = re.search(self.patterns['bloque_cliente'], text, re.DOTALL | re.IGNORECASE)
        bloque_contrato_match = re.search(self.patterns['bloque_contrato'], text, re.DOTALL | re.IGNORECASE)

        if bloque_cliente_match:
            bloque_cliente = bloque_cliente_match.group(1)
            data['Cliente'] = self._search_pattern(self.patterns['cliente'], bloque_cliente)
            data['Identificación'] = self._search_pattern(self.patterns['identificacion'], bloque_cliente)

        if bloque_contrato_match:
            bloque_contrato = bloque_contrato_match.group(1)
            data['Contrato'] = self._search_pattern(self.patterns['contrato'], bloque_contrato)
            data['Dirección'] = self._search_pattern(self.patterns['direccion'], bloque_contrato, flags=re.DOTALL)

        # General searches
        data['No. Solicitud'] = self._search_pattern(self.patterns['no_solicitud'], text, flags=re.DOTALL)
        data['No. Rel. Pago'] = self._search_pattern(self.patterns['no_rel_pago'], text, flags=re.DOTALL)
        # Try multiple variants for Tipo de Cupón
        data['Tipo de Cupón'] = self._search_first(['tipo_cupon', 'tipo_cupon_alt1', 'tipo_cupon_alt2'], text, flags=re.DOTALL)
        data['Valido hasta'] = self._search_pattern(self.patterns['valido_hasta'], text, flags=re.DOTALL)
        data['Valor a Pagar'] = self._search_pattern(self.patterns['valor_pagar'], text, flags=re.DOTALL)
        # Robust barcode extraction (first well-formed occurrence only)
        data['Codigo de Barras Limpio'] = self._extract_barcode(text)

        # Harden No. Solicitud with very tolerant fallback (handles odd encodings/spaces)
        if data.get('No. Solicitud', 'No encontrado') == 'No encontrado':
            m = re.search(r'Solici[^\d]{0,20}(\d{6,})', text, re.IGNORECASE | re.DOTALL)
            if m:
                data['No. Solicitud'] = m.group(1)

        # Normalize Tipo (e.g., 'N G' -> 'NG')
        if data.get('Tipo de Cupón') not in (None, 'No encontrado'):
            data['Tipo de Cupón'] = re.sub(r'\s+', '', data['Tipo de Cupón']).upper()

        return data

    def _extract_coupon_data(self, text):
        """Extract data from coupon-style text with key-value pairs."""
        data = {}
        lines = text.split('\n')
        current_key = None

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Check if line starts with a potential key (uppercase Spanish words)
            if re.match(r'^[A-ZÑÁÉÍÓÚ\s]{3,}:?', line):
                # Remove colon if present
                key = re.sub(r':$', '', line).strip()
                current_key = self._map_coupon_key(key)
            elif current_key and line:
                # This is the value for the previous key
                data[current_key] = line
                current_key = None

        return data

    def _map_coupon_key(self, key):
        """Map coupon keys to standard fields."""
        key_mapping = {
            'NOMBRE': 'Cliente',
            'NOMBRE COMPLETO': 'Cliente',
            'CLIENTE': 'Cliente',
            'IDENTIFICACIÓN': 'Identificación',
            'ID': 'Identificación',
            'CÉDULA': 'Identificación',
            'DIRECCIÓN': 'Dirección',
            'DIR': 'Dirección',
            'CONTRATO': 'Contrato',
            'No. CONTRATO': 'Contrato',
            'SOLICITUD': 'No. Solicitud',
            'No. SOLICITUD': 'No. Solicitud',
            'REL. PAGO': 'No. Rel. Pago',
            'No. REL. PAGO': 'No. Rel. Pago',
            'TIPO DE CUPÓN': 'Tipo de Cupón',
            'VÁLIDO HASTA': 'Valido hasta',
            'VALOR A PAGAR': 'Valor a Pagar',
            'CÓDIGO DE BARRAS': 'Codigo de Barras Limpio'
        }
        return key_mapping.get(key.upper(), None)

    def _search_pattern(self, pattern, text, flags=0):
        """Search for pattern and return cleaned match."""
        match = re.search(pattern, text, re.IGNORECASE | flags)
        if match:
            cleaned = match.group(1).strip().replace('\n', ' ')
            return re.sub(r'\s+', ' ', cleaned)
        return "No encontrado"

    def _clean_barcode(self, barcode_text):
        """Clean barcode by removing non-digits."""
        if barcode_text and barcode_text != "No encontrado":
            return re.sub(r'\D', '', barcode_text)
        return "No encontrado"

    def _normalize_text(self, text):
        """Normalize artifacts and whitespace from extracted text."""
        if not text:
            return ""
        # Replace common PDF artifact sequences
        text = text.replace('(cid:13)(cid:10)', ' ')
        # Normalize non-breaking and special spaces
        text = text.replace('\u00A0', ' ').replace('\u202F', ' ').replace('\u2007', ' ')
        # Fix common mis-encoded Spanish tokens seen in PDFs
        replacements = {
            'Identificacin': 'Identificación',
            'Cupn': 'Cupón',
            'Lnea': 'Línea',
        }
        for k, v in replacements.items():
            text = text.replace(k, v)
        # Normalize ordinal indicators to 'No.' without touching plain 'N'
        text = text.replace('Nº', 'No.').replace('N°', 'No.')
        # Collapse excessive whitespace
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\s+\n', '\n', text)
        text = re.sub(r'\n\s+', '\n', text)
        return text


    def _search_first(self, keys, text, flags=0):
        """Try multiple pattern keys and return the first successful match."""
        for key in keys:
            pattern = self.patterns.get(key)
            if not pattern:
                continue
            val = self._search_pattern(pattern, text, flags=flags)
            if val != "No encontrado":
                return val
        return "No encontrado"

    def _extract_barcode(self, text):
        """Extract the first well-formed barcode block and clean it to digits."""
        try:
            pattern = self.patterns.get('codigo_barras')
            if not pattern:
                return "No encontrado"
            matches = re.findall(pattern, text, flags=re.IGNORECASE | re.DOTALL)
            if not matches:
                return "No encontrado"
            raw = matches[0][0] if isinstance(matches[0], tuple) else matches[0]
            # Keep only digits
            digits = re.sub(r'\D', '', raw)
            # Basic sanity check: typical length around 44-64 digits
            if len(digits) < 30:
                return "No encontrado"
            return digits
        except Exception:
            return "No encontrado"

    def _is_ocr_text(self, text):
        """Check if text appears to be OCR output (has OCR artifacts)."""
        if not text:
            return False

        # OCR text often has:
        # - Many short lines
        # - Inconsistent spacing
        # - Character recognition errors (like 0 instead of O)
        # - Lower case letters mixed with upper case in unexpected ways

        lines = text.split('\n')
        short_lines = sum(1 for line in lines if len(line.strip()) < 10)
        short_line_ratio = short_lines / len(lines) if lines else 0

        # Check for OCR common errors
        ocr_errors = re.findall(r'\b\d+[A-Z]+\d*\b', text)  # Numbers mixed with letters
        error_ratio = len(ocr_errors) / len(text.split()) if text.split() else 0

        # If many short lines or OCR-like errors, likely OCR text
        return short_line_ratio > 0.5 or error_ratio > 0.1