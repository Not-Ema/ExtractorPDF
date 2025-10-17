import os
import re
import threading
import queue
import platform
import shutil
import json
from pdf2image import convert_from_path
import pytesseract
from PIL import Image, ImageFilter, ImageOps
import pandas as pd
from .logger import logger

# ---------- Default CONFIG ----------
DEFAULT_DPI = 600
DEFAULT_LANG = "spa"

# ---------- Find Tesseract executable ----------
def find_tesseract():
    """Find Tesseract executable in common locations."""
    possible_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        shutil.which('tesseract'),  # Check PATH
    ]
    for path in possible_paths:
        if path and os.path.exists(path):
            return path
    return None

# ---------- OCR + extracción (adaptado) ----------
def image_preprocess(img: Image.Image, upscale_if_small=True) -> Image.Image:
    # Convertir a escala de grises
    img = img.convert("L")

    # Upscaling para mejorar resolución de OCR
    w, h = img.size
    if upscale_if_small and w < 2000:  # aumenta más el umbral
        img = img.resize((int(w*2), int(h*2)), Image.LANCZOS)

    # Aumentar contraste
    img = ImageOps.autocontrast(img)

    # Binarización (convertir a blanco y negro puro)
    img = img.point(lambda x: 0 if x < 160 else 255, '1')

    # Filtrar ruido
    img = img.filter(ImageFilter.MedianFilter(size=3))

    # Convertir de nuevo a L para el OCR
    img = img.convert("L")

    return img

def clean_barcode(s: str) -> str:
    if not s:
        return None
    cleaned = re.sub(r'[^0-9A-Z]', '', s.upper())
    return cleaned or None

PATTERNS = {
    "Cliente": r"Cliente[:\s\-]*([A-ZÁÉÍÓÚÑa-záéíóúñ\.\,\s]{3,100})",
    "Contrato": r"Contrato[:\s\-]*([0-9]{3,20})",
    "NoSolicitud": r"No\.?\s*Solicit(?:ud|ion)[:\s\-]*([0-9]{4,20})",
    "TipoCupon": r"Tipo\s+de\s+Cup[oó]n[:\s\-]*([A-Z0-9\-\s]{1,40})",
    "ValorAPagar": r"Valor\s+a\s+pagar[:\s\-]*\$?\s*([0-9\.,]{1,20})",
    "Identificacion": r"Identificaci[oó]n[:\s\-]*([0-9\-\s]{6,25})",
    "DirCliente": r"Dir(?:\.|eccion|\.? Cliente)[:\s\-]*([A-Z0-9\-\.,#\s]{5,150})",
    "NoRefPago": r"No\.?\s*Ref\.?\s*\.?Pago[:\s\-]*([0-9]{5,30})",
    "ValidoHasta": r"Valido\s+hasta[:\s\-]*([0-9]{1,2}[-/][A-Z]{3,}[-/][0-9]{4}|[0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4}|[0-9]{4}[-/][0-9]{2}[-/][0-9]{2})"
}

def find_first(pattern, text, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    if m:
        return m.group(1).strip()
    return None

def clean_digits(s: str) -> str:
    """Devuelve solo los dígitos de una cadena; si hay pocos dígitos, devuelve None."""
    if not s:
        return None
    d = re.sub(r'\D', '', s)
    return d if len(d) >= 1 else None

# ---------- Normalización y heurísticas para montos OCR corruptos ----------
def normalize_ocr_amount_text(s: str, context=""):
    """
    Normaliza una cadena OCR que pretende ser un monto.
    - Reemplaza caracteres confundidos por dígitos.
    - Extrae la secuencia de dígitos más plausible (preferencia >=4 dígitos).
    - Devuelve string solo con dígitos o None si no encuentra nada razonable.
    """
    if not s:
        return None
    raw = str(s).strip()

    # 1) Mapeo conservador de confusiones comunes (mayúsculas/minúsculas)
    #    Ajusta según lo que veas: O->0, o->0, I->1, l->1, i->1, S->5, Z->2, B->8
    #    K/X a veces aparece antes de números => si está al inicio lo quitamos
    replacements = {
        'O': '0', 'o': '0',
        'I': '1', 'l': '1', 'i': '1',
        'S': '5', 's': '5',
        'Z': '2', 'z': '2',
        'B': '8', 'b': '8',
        # algunos ruídos comunes: remove non-sense letters near digits
        'K': '', 'k': '', 'X': '', 'x': '', 'Q': '0'
    }

    # apply mapping character by character but keep ., and ,
    mapped_chars = []
    for ch in raw:
        if ch in replacements:
            mapped_chars.append(replacements[ch])
        else:
            mapped_chars.append(ch)
    mapped = ''.join(mapped_chars)

    # 2) Remove stray characters except digits and . ,
    cleaned = re.sub(r"[^0-9\.,]", "", mapped)

    # 3) If we have thousands separators like '530,000' or '53.000', remove separators and return digits
    # Prefer the longest digit group >=3
    digit_groups = re.findall(r"\d{3,}", cleaned)
    if digit_groups:
        # choose longest group (most likely full amount)
        best = max(digit_groups, key=len)
        return re.sub(r'[^0-9]', '', best)

    # 4) else, try to extract ANY group of >=2 digits (relajado a 3 si quieres)
    dg2 = re.findall(r"\d{2,}", cleaned)
    if dg2:
        best = max(dg2, key=len)
        # only accept if at least 3 digits (to avoid 0/01 noise); adjust to 4 if needed
        if len(best) >= 3:
            return best

    return None

# ---------- Improved ocr_amount_from_image (slightly adjusted) ----------
def ocr_amount_from_image(pil_image: Image.Image, keyword="VALOR", expand_px=120):
    """
    Re-ocr de una región alrededor de 'keyword'. Devuelve cadena normalizada de monto (solo dígitos)
    o None.
    """
    try:
        from pytesseract import Output
        data = pytesseract.image_to_data(pil_image, lang='spa', output_type=Output.DICT)
    except Exception:
        return None

    # buscar ocurrencias del keyword (insensible a mayúsculas)
    idxs = [i for i, t in enumerate(data['text']) if t and keyword.upper() in t.upper()]
    if not idxs:
        return None

    i = idxs[0]
    x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
    x0 = max(0, x - expand_px)
    y0 = max(0, y - expand_px)
    x1 = min(pil_image.width, x + w + expand_px)
    y1 = min(pil_image.height, y + h + expand_px)
    crop = pil_image.crop((x0, y0, x1, y1))

    # OCR restringido (whitelist)
    custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789.,$'
    try:
        text = pytesseract.image_to_string(crop, lang='spa', config=custom_config)
        return normalize_ocr_amount_text(text)
    except Exception:
        return None

def extract_fields_from_text(text: str) -> dict:
    """
    Heurística mejorada:
    - Cliente (MAYÚSCULAS)
    - Identificacion, Contrato, DirCliente, NoSolicitud, NoRefPago, TipoCupon,
      ValidoHasta, ValorAPagar, CodigoBarraRaw, CodigoBarraLimpio
    """
    data = {
        "Cliente": None, "Identificacion": None, "Contrato": None, "DirCliente": None,
        "NoSolicitud": None, "NoRefPago": None, "TipoCupon": None, "ValidoHasta": None,
        "ValorAPagar": None, "CodigoBarraRaw": None, "CodigoBarraLimpio": None
    }

    txt = text.replace("\r", "\n")

    # --- Cliente: asume que está en MAYÚSCULAS y tras "Cliente:"
    m = re.search(r"Cliente[:\s]*([A-ZÁÉÍÓÚÑ0-9\.\s]{2,200})", txt)
    if m:
        raw_name = m.group(1).strip()
        # elimina sufijos que empiecen por ' a <digits>' o números largos (tel), o tokens de dirección
        cliente = re.sub(r"\s+(?:a\s*\d{3,}|\d{6,}|\bKR\b|\bCL\b|\bAV\b|PBX|FAX|Fax|Contrato|No\.)\b[\s\S]*$", "", raw_name, flags=re.IGNORECASE).strip()
        cliente = re.sub(r"[,\-\s]+$", "", cliente)
        # asegurar mayúsculas (según pediste)
        data["Cliente"] = cliente.upper() if cliente else None

    # --- Identificación: etiqueta o número en la misma línea que cliente
    m = re.search(r"Identificaci[oó]n[:\s]*([\d\-\s]{6,20})", txt, flags=re.IGNORECASE)
    if not m:
        m2 = re.search(r"Cliente[:\s].*?(\d{6,12})", txt)
        if m2:
            m = m2
    if m:
        data["Identificacion"] = re.sub(r"\D", "", m.group(1))

    # --- Contrato
    m = re.search(r"Contrato[:\s]*([0-9]{3,20})", txt, flags=re.IGNORECASE)
    if m:
        data["Contrato"] = m.group(1)

    # --- DirCliente (intenta etiqueta o patrón 'KR/CL/AV')
    m = re.search(r"Dir(?:\.|eccion)?(?:\.|:)?\s*Cliente[:\s]*([A-Z0-9ÁÉÍÓÚÑ\-\.,#\s]{3,200})", txt, flags=re.IGNORECASE)
    if not m:
        matches = re.findall(r"((?:KR|CL|AV|C[^\n]{1,30}|[A-Z]{2,5}\s*\d{1,3})[^\n]{0,150})", txt, flags=re.IGNORECASE)
        for match in matches:
            if "cliente" not in match.lower():
                data["DirCliente"] = match.strip().split("\n")[0].strip()
                break

    # --- NoRefPago
    m = re.search(r"No\.?\s*Ref\.?\s*[:\s]*Pago[:\s]*([0-9]{5,30})", txt, flags=re.IGNORECASE)
    if not m:
        m = re.search(r"No\.?\s*Ref\.?\s*[:\s]*([0-9]{5,30})", txt, flags=re.IGNORECASE)
    if m:
        data["NoRefPago"] = m.group(1)

    # --- TipoCupon
    m = re.search(r"Tipo\s*(?:de)?\s*Cup[oó]n[:\s]*([A-Z0-9\-]{1,20})", txt, flags=re.IGNORECASE)
    if not m:
        m = re.search(r"Tipo(?:\s+de)?[:\s]*([A-Z]{1,6})", txt, flags=re.IGNORECASE)
    if m:
        data["TipoCupon"] = m.group(1).strip().upper()

    # --- ValidoHasta (fecha dd-MMM-YYYY, corrige 0 por O en los meses)
    m = re.search(r"([0-9]{1,2}[-/][A-Z0-9]{3,}[-/][0-9]{4})", txt, flags=re.IGNORECASE)
    if m:
        fecha = m.group(1).upper()
        # Corrige confusiones del OCR: 0 → O en el mes
        # (pero solo si no parece una fecha numérica tipo 01-10-2025)
        partes = fecha.split("-")
        if len(partes) == 3 and not partes[1].isdigit():
            partes[1] = partes[1].replace("0", "O")
            fecha = "-".join(partes)
        data["ValidoHasta"] = fecha

    # ------------------ REEMPLAZAR POR ESTE BLOQUE ------------------
    # 9) Valor a pagar: priorizamos en este orden:
    #  1) montos con signo $ (ej: $20,000.00)
    #  2) texto explícito "Valor a pagar" (aunque no tenga $) — extraer números cercanos
    #  3) etiquetas Total / Total Efectivo cercanas
    #  4) secuencia numérica razonable que NO sea una fecha (evitar dd-MMM-YYYY)
    amount = None

    # helper: comprueba si la cadena es un formato de fecha (dd-MMM-YYYY o dd/mm/yyyy)
    def is_date_like(s):
        if not s:
            return False
        s = str(s).strip()
        if re.search(r"\b\d{1,2}[-/][A-Z]{3}[-/]\d{4}\b", s, flags=re.IGNORECASE):
            return True
        if re.search(r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b", s):
            return True
        return False

    # 1) buscar montos con $ (mejor precisión)
    m = re.search(r"\$\s*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)", txt)
    if m:
        amount = re.sub(r'[^0-9]', '', m.group(1))

    # 2) si no hay $, buscar "Valor a pagar" y extraer el primer número que no sea fecha
    if not amount:
        m = re.search(r"Valor\s*a\s*pagar[:\s]*([^\n\r]{1,60})", txt, flags=re.IGNORECASE)
        if m:
            candidate = m.group(1)
            # buscar dentro del candidate cualquier secuencia de dígitos/grupos de miles
            m2 = re.search(r"([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)", candidate)
            if m2:
                cand = m2.group(1)
                if not is_date_like(cand):
                    amount = re.sub(r'[^0-9]', '', cand)

    # 3) fallback: buscar cerca de 'Total' o 'Total Efectivo'
    if not amount:
        for kw in ["Total Efectivo", "Total Efectivo:", "Total Cheques", "Total", "TOTAL"]:
            pos = txt.find(kw)
            if pos != -1:
                window = txt[max(0, pos-40): pos+80]
                m3 = re.search(r"([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)", window)
                if m3 and not is_date_like(m3.group(1)):
                    amount = re.sub(r'[^0-9]', '', m3.group(1))
                    break

    # 4) último recurso: buscar cualquier número largo razonable que NO sea fecha y no coincida con año o contrato/ref
    if not amount:
        cand_nums = re.findall(r"\d{3,30}", txt)
        # eliminar fechas y valores que son claramente año/día con 4 dígitos
        cand_nums = [c for c in cand_nums if not re.search(r"^\d{4}$", c)]  # descarta solo-año 2025
        # eliminar secuencias que coincidan con la fecha literal ddmmYYYY si están en el texto de la fecha
        cand_nums = [c for c in cand_nums if not re.search(r"\d{2}0\d{2,}", c)]
        # elegir la primera candidata que no sea una fecha y que tenga al menos 3 dígitos
        chosen = None
        for c in cand_nums:
            if len(c) >= 3 and not is_date_like(c):
                # evitar capturar el mismo número que es Contrato/NoRef/Identificación si ya lo conocemos:
                if data.get("Contrato") and c == data.get("Contrato"):
                    continue
                if data.get("NoRefPago") and c == data.get("NoRefPago"):
                    continue
                if data.get("Identificacion") and c == data.get("Identificacion"):
                    continue
                chosen = c
                break
        if chosen:
            amount = chosen

    if amount:
        # normalizar: dejar solo dígitos, sin decimales (si quieres decimales los puedes mantener)
        data["ValorAPagar"] = re.sub(r'[^0-9]', '', str(amount))
    else:
        data["ValorAPagar"] = None
# ------------------ FIN BLOQUE ------------------

    # --- Código de barra: última línea con paréntesis/dígitos largos
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    barcode_line = None
    for ln in reversed(lines):
        if re.search(r"\(\d{2,}\)", ln) or (len(re.sub(r"\D", "", ln)) >= 20):
            barcode_line = ln
            break
    if barcode_line:
        data["CodigoBarraRaw"] = barcode_line
        data["CodigoBarraLimpio"] = clean_barcode(barcode_line)

    if not data["ValorAPagar"] and data["CodigoBarraLimpio"]:
        m = re.search(r'(\d{5})96', data["CodigoBarraLimpio"])
        if m:
            data["ValorAPagar"] = m.group(1)

    # ---------------- REEMPLAZAR LA SECCIÓN NoSolicitud POR ESTE BLOQUE ----------------
    data["NoSolicitud"] = None

    # 1) buscar líneas que contengan la palabra "Solicitud" (varias variantes)
    solicitud_lines = []
    for ln in txt.splitlines():
        if re.search(r"\bSolicit(?:ud|ion)\b", ln, flags=re.IGNORECASE):
            solicitud_lines.append(ln.strip())

    found = None
    # intentar extraer número directamente en esas líneas
    for ln in solicitud_lines:
        m = re.search(r"([0-9]{5,30})", ln)
        if m:
            found = m.group(1)
            break

    # 2) si no hay, encontrar la posición de la primera ocurrencia de "Solicitud" y tomar el número más cercano en el texto
    if not found:
        mpos = re.search(r"\bSolicit(?:ud|ion)\b", txt, flags=re.IGNORECASE)
        if mpos:
            pos = mpos.start()
            all_nums = [(mo.group(0), mo.start()) for mo in re.finditer(r"[0-9]{5,30}", txt)]
            if all_nums:
                best = min(all_nums, key=lambda x: abs(x[1] - pos))
                found = best[0]

    # 3) fallback: buscar "No. Solicitud" explícito en cualquier parte
    if not found:
        m = re.search(r"No\.?\s*Solicit(?:ud|ion)[:\s\-]*([0-9]{5,30})", txt, flags=re.IGNORECASE)
        if m:
            found = m.group(1)

    if found:
        data["NoSolicitud"] = re.sub(r"\D", "", found)
    else:
        data["NoSolicitud"] = None
# --------------- FIN BLOQUE NoSolicitud ------------------

    return data

def extract_text_from_pdf(pdf_path, dpi=DEFAULT_DPI, lang=DEFAULT_LANG, tesseract_config="--psm 6", save_ocr_text=False, ocr_text_dir=None, logger=None):
    pages = convert_from_path(pdf_path, dpi=dpi)
    texts = []
    for _i, page in enumerate(pages):
        img = image_preprocess(page)
        text = pytesseract.image_to_string(img, lang=lang, config=tesseract_config)
        texts.append(text)
    full_text = "\n\n".join(texts)
    if save_ocr_text and ocr_text_dir:
        os.makedirs(ocr_text_dir, exist_ok=True)
        fn = os.path.splitext(os.path.basename(pdf_path))[0] + ".txt"
        with open(os.path.join(ocr_text_dir, fn), "w", encoding="utf-8") as f:
            f.write(full_text)
    return full_text

class OCRProcessor:
    def __init__(self, config_path='config/config.json'):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        self.tesseract_cmd = None
        self._setup_tesseract()

    def _setup_tesseract(self):
        """Setup Tesseract executable."""
        # Auto-detect Tesseract
        tesseract_exe = find_tesseract()
        if tesseract_exe:
            pytesseract.pytesseract.tesseract_cmd = tesseract_exe
            self.tesseract_cmd = tesseract_exe
        else:
            logger.warning("Tesseract executable not found. OCR functionality will be limited.")

        # Set TESSDATA_PREFIX if possible
        if self.tesseract_cmd:
            tessdata_dir = os.path.join(os.path.dirname(self.tesseract_cmd), 'tessdata')
            if os.path.exists(tessdata_dir):
                os.environ['TESSDATA_PREFIX'] = tessdata_dir

    def extract_text_from_scan(self, pdf_path, dpi=DEFAULT_DPI, lang=DEFAULT_LANG, save_ocr_text=False, ocr_text_dir=None):
        """Extract text from scanned PDF using OCR."""
        try:
            return extract_text_from_pdf(
                pdf_path, dpi=dpi, lang=lang,
                tesseract_config="--psm 6",
                save_ocr_text=self.save_ocr_text if hasattr(self, 'save_ocr_text') else save_ocr_text,
                ocr_text_dir=self.ocr_text_dir if hasattr(self, 'ocr_text_dir') else ocr_text_dir,
                logger=logger
            )
        except Exception as e:
            logger.error(f"OCR extraction failed for {pdf_path}: {e}")
            return ""

    def extract_fields_from_ocr_text(self, text):
        """Extract fields from OCR text."""
        return extract_fields_from_text(text)