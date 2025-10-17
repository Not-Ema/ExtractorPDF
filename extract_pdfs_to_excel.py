"""
pdf_ocr_gui.py
GUI (Tkinter) para extraer campos de PDFs por OCR y guardarlos en Excel.
Incluye ventanas Toplevel personalizadas para:
 - seleccionar carpeta con PDFs
 - seleccionar carpeta y nombre de archivo para guardar el Excel
"""

import os
import re
import threading
import queue
import platform
import shutil
from tkinter import Tk, StringVar, IntVar, BooleanVar, Toplevel, filedialog, messagebox, ttk, scrolledtext, Label, Button, Entry, Checkbutton
from pdf2image import convert_from_path
import pytesseract
from PIL import Image, ImageFilter, ImageOps
import pandas as pd

# ---------- Default CONFIG ----------
DEFAULT_DPI = 600
DEFAULT_LANG = "spa"
# ------------------------------------

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
# ---------- OCR + extracción (adaptado y mejorado) ----------
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

import re

import re

def clean_barcode(s: str) -> str:
    if not s:
        return None
    return re.sub(r'[^0-9A-Z]', '', s.upper())

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
        cliente = re.sub(r"\s+(?:a\s*\d{3,}|\d{6,}|\bKR\b|\bCL\b|\bAV\b|PBX|FAX|Fax|Contrato|No\.?)\b[\s\S]*$", "", raw_name, flags=re.IGNORECASE).strip()
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


def extract_text_from_pdf(pdf_path, dpi=300, lang='spa', tesseract_config="--psm 6", save_ocr_text=False, ocr_text_dir=None, logger=None):
    pages = convert_from_path(pdf_path, dpi=dpi)
    texts = []
    for _i, page in enumerate(pages):
        img = image_preprocess(page)
        text = pytesseract.image_to_string(img, lang=lang, config=tesseract_config)
        texts.append(text)
    full_text = "\n\n".join(texts)
    if save_ocr_text and ocr_text_dir:
        fn = os.path.splitext(os.path.basename(pdf_path))[0] + ".txt"
        with open(os.path.join(ocr_text_dir, fn), "w", encoding="utf-8") as f:
            f.write(full_text)
    return full_text

# ---------- Worker: procesa una carpeta ----------
def process_all_pdfs(input_folder, output_excel, dpi, lang, tesseract_cmd, save_ocr_text, ocr_text_dir, progress_queue, log_queue, stop_event):
    try:
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
        else:
            # Auto-detect Tesseract if not specified
            tesseract_exe = find_tesseract()
            if tesseract_exe:
                pytesseract.pytesseract.tesseract_cmd = tesseract_exe
            else:
                log_queue.put("Error: Tesseract executable not found. Please install Tesseract OCR from https://github.com/UB-Mannheim/tesseract/wiki or specify the correct path in the GUI.")
                progress_queue.put(("done", 0, 0))
                return

        # Set TESSDATA_PREFIX to the tessdata directory
        tesseract_exe = pytesseract.pytesseract.tesseract_cmd
        if not os.path.exists(tesseract_exe):
            log_queue.put(f"Error: Tesseract executable not found at {tesseract_exe}. Please install Tesseract OCR or specify the correct path in the GUI.")
            progress_queue.put(("done", 0, 0))
            return
        tessdata_dir = os.path.join(os.path.dirname(tesseract_exe), 'tessdata')
        os.environ['TESSDATA_PREFIX'] = tessdata_dir

        # Check if language data files exist
        lang_files = lang.split('+')
        for l in lang_files:
            l = l.strip()
            if not os.path.exists(os.path.join(tessdata_dir, f'{l}.traineddata')):
                log_queue.put(f"Error: Language data file '{l}.traineddata' not found in {tessdata_dir}. Please download it from https://github.com/tesseract-ocr/tessdata and place it in the tessdata directory.")
                progress_queue.put(("done", 0, 0))
                return

        files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
        total = len(files)
        if total == 0:
            log_queue.put("No se encontraron archivos PDF en la carpeta seleccionada.")
            progress_queue.put(("done", 0, 0))
            return

        if save_ocr_text:
            os.makedirs(ocr_text_dir, exist_ok=True)

        rows = []
        for idx, pdf in enumerate(files, start=1):
            if stop_event.is_set():
                log_queue.put("Proceso cancelado por el usuario.")
                break
            log_queue.put(f"Procesando: {os.path.basename(pdf)} ({idx}/{total}) ...")
            try:
                text = extract_text_from_pdf(pdf, dpi=dpi, lang=lang, tesseract_config="--psm 6", save_ocr_text=save_ocr_text, ocr_text_dir=ocr_text_dir)
                fields = extract_fields_from_text(text)
                fields["_file"] = os.path.basename(pdf)
                rows.append(fields)
                log_queue.put(f"  -> OK")
            except Exception as e:
                log_queue.put(f"  -> ERROR: {e}")
                rows.append({"_file": os.path.basename(pdf), "error": str(e)})
            progress_queue.put(("progress", idx, total))

        # Guardar Excel
        df = pd.DataFrame(rows)
        # ordenar columnas
        cols_order = ["_file", "Cliente", "Contrato", "Identificacion", "NoSolicitud",
                      "TipoCupon", "ValorAPagar", "NoRefPago", "DirCliente", "ValidoHasta",
                      "CodigoBarraRaw", "CodigoBarraLimpio", "error"]
        cols = [c for c in cols_order if c in df.columns] + [c for c in df.columns if c not in cols_order]
        df = df[cols]
        df.to_excel(output_excel, index=False)
        log_queue.put(f"Excel guardado en: {output_excel}")
        progress_queue.put(("done", total, total))
    except Exception as e:
        log_queue.put(f"Fallo inesperado: {e}")
        progress_queue.put(("done", 0, 0))

# ---------- Folder browser utilities (Toplevel) ----------
def get_roots():
    system = platform.system().lower()
    if system == "windows":
        # Mostrar todas las letras de disco disponibles
        drives = []
        for letter in range(67):  # A..Z approx
            drive = f"{chr(65 + letter)}:\\"
            if os.path.exists(drive):
                drives.append(drive)
        return drives if drives else [os.path.expanduser("~")]
    else:
        # linux/mac -> root
        return ["/"]

def list_dir(path):
    try:
        with os.scandir(path) as it:
            dirs = [entry.name for entry in it if entry.is_dir()]
        dirs.sort()
        return dirs
    except PermissionError:
        return []
    except FileNotFoundError:
        return []

# ---------- GUI ----------
class OCRGui:
    def __init__(self, root):
        self.root = root
        root.title("OCR PDF -> Excel")
        root.geometry("820x540")

        # Variables
        self.input_folder = StringVar()
        self.output_file = StringVar()
        self.dpi = IntVar(value=DEFAULT_DPI)
        self.lang = StringVar(value=DEFAULT_LANG)
        self.tesseract_cmd = StringVar(value="")
        self.save_ocr_text = BooleanVar(value=False)
        self.ocr_text_dir = StringVar(value=os.path.join(os.getcwd(), "ocr_texts"))

        # Auto-detect Tesseract on startup
        detected = find_tesseract()
        if detected:
            self.tesseract_cmd.set(detected)

        # Queues and thread control
        self.progress_queue = queue.Queue()
        self.log_queue = queue.Queue()
        self.worker_thread = None
        self.stop_event = threading.Event()

        # Layout
        padx = 8
        pady = 6
        row = 0
        Label(root, text="Carpeta con PDFs:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
        Entry(root, textvariable=self.input_folder, width=68).grid(row=row, column=1, columnspan=3, sticky="w", padx=padx, pady=pady)
        Button(root, text="Seleccionar (ventana)", command=self.select_input_folder).grid(row=row, column=4, padx=padx, pady=pady)
        Button(root, text="Seleccionar (nativo)", command=self.select_input_folder_native).grid(row=row, column=5, padx=2, pady=pady)

        row += 1
        Label(root, text="Guardar Excel como:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
        Entry(root, textvariable=self.output_file, width=68).grid(row=row, column=1, columnspan=3, sticky="w", padx=padx, pady=pady)
        Button(root, text="Seleccionar (ventana)", command=self.select_output_file).grid(row=row, column=4, padx=padx, pady=pady)
        Button(root, text="Seleccionar (nativo)", command=self.select_output_file_native).grid(row=row, column=5, padx=2, pady=pady)

        row += 1
        Label(root, text="DPI (OCR):").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
        Entry(root, textvariable=self.dpi, width=8).grid(row=row, column=1, sticky="w", padx=padx, pady=pady)
        Label(root, text="Idioma Tesseract (ej: spa, eng, spa+eng):").grid(row=row, column=2, sticky="w", padx=padx, pady=pady)
        Entry(root, textvariable=self.lang, width=18).grid(row=row, column=3, sticky="w", padx=padx, pady=pady)

        row += 1
        Label(root, text="Ruta Tesseract (opcional):").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
        Entry(root, textvariable=self.tesseract_cmd, width=68).grid(row=row, column=1, columnspan=3, sticky="w", padx=padx, pady=pady)
        Button(root, text="Examinar", command=self.select_tesseract).grid(row=row, column=4, padx=padx, pady=pady)

        row += 1
        Checkbutton(root, text="Guardar OCR .txt por PDF", variable=self.save_ocr_text).grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
        Entry(root, textvariable=self.ocr_text_dir, width=56).grid(row=row, column=1, columnspan=3, sticky="w", padx=padx, pady=pady)
        Button(root, text="Carpeta OCR", command=self.select_ocr_text_dir).grid(row=row, column=4, padx=padx, pady=pady)

        row += 1
        self.start_btn = Button(root, text="Iniciar", command=self.start_processing, width=14)
        self.start_btn.grid(row=row, column=1, pady=12)
        self.cancel_btn = Button(root, text="Cancelar", command=self.cancel_processing, width=14, state="disabled")
        self.cancel_btn.grid(row=row, column=2, pady=12)

        row += 1
        Label(root, text="Progreso:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
        self.progress = ttk.Progressbar(root, orient="horizontal", length=620, mode="determinate")
        self.progress.grid(row=row, column=1, columnspan=4, sticky="w", padx=padx, pady=pady)

        row += 1
        Label(root, text="Logs:").grid(row=row, column=0, sticky="nw", padx=padx, pady=pady)
        self.logbox = scrolledtext.ScrolledText(root, width=100, height=14, wrap="word")
        self.logbox.grid(row=row, column=1, columnspan=5, padx=padx, pady=pady)

        # Poll queues
        root.after(200, self._poll_queues)

    # ---------- Custom folder selection windows ----------
    def select_input_folder(self):
        # open custom Toplevel folder browser (modal)
        start = self.input_folder.get().strip() or os.path.expanduser("~")
        folder = open_folder_browser(self.root, title="Seleccionar carpeta con PDFs", start_path=start, only_directories=True)
        if folder:
            self.input_folder.set(folder)

    def select_output_file(self):
        # open custom Toplevel save browser (choose folder + filename)
        start_folder = os.path.dirname(self.output_file.get()) if self.output_file.get() else os.path.expanduser("~")
        suggested = os.path.basename(self.output_file.get()) or "extracted_data.xlsx"
        folder, filename = open_save_browser(self.root, title="Guardar Excel como", start_path=start_folder, suggested_name=suggested)
        if folder and filename:
            self.output_file.set(os.path.join(folder, filename))

    # fallback native dialogs (por si prefieres)
    def select_input_folder_native(self):
        d = filedialog.askdirectory(title="Seleccionar carpeta con PDFs")
        if d:
            self.input_folder.set(d)

    def select_output_file_native(self):
        f = filedialog.asksaveasfilename(title="Guardar Excel", defaultextension=".xlsx",
                                         filetypes=[("Excel files","*.xlsx"), ("All files","*.*")])
        if f:
            self.output_file.set(f)

    def select_tesseract(self):
        f = filedialog.askopenfilename(title="Seleccionar ejecutable Tesseract (si aplica)")
        if f:
            self.tesseract_cmd.set(f)

    def select_ocr_text_dir(self):
        d = filedialog.askdirectory(title="Carpeta para guardar OCR (.txt)")
        if d:
            self.ocr_text_dir.set(d)

    # ---------- Processing control ----------
    def start_processing(self):
        input_folder = self.input_folder.get().strip()
        output_file = self.output_file.get().strip()
        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Selecciona una carpeta válida con PDFs.")
            return
        if not output_file:
            messagebox.showerror("Error", "Selecciona dónde guardar el archivo Excel.")
            return

        dpi = int(self.dpi.get())
        lang = self.lang.get().strip() or DEFAULT_LANG
        tesseract_cmd = self.tesseract_cmd.get().strip() or None
        save_ocr_text = bool(self.save_ocr_text.get())
        ocr_text_dir = self.ocr_text_dir.get().strip() if save_ocr_text else None

        # disable start, enable cancel
        self.start_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.logbox.delete("1.0", "end")
        self.progress['value'] = 0

        # prepare thread
        self.stop_event.clear()
        self.worker_thread = threading.Thread(
            target=process_all_pdfs,
            args=(input_folder, output_file, dpi, lang, tesseract_cmd, save_ocr_text, ocr_text_dir, self.progress_queue, self.log_queue, self.stop_event),
            daemon=True
        )
        self.worker_thread.start()
        self.logbox.insert("end", "Iniciado procesamiento...\n")
        self.logbox.see("end")

    def cancel_processing(self):
        if messagebox.askyesno("Confirmar", "¿Deseas cancelar el proceso en curso?"):
            self.stop_event.set()
            self.cancel_btn.config(state="disabled")
            self.logbox.insert("end", "Cancelando... por favor espera.\n")
            self.logbox.see("end")

    # Polling for updates from worker
    def _poll_queues(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                self.logbox.insert("end", item + "\n")
                self.logbox.see("end")
        except queue.Empty:
            pass

        try:
            while True:
                evt = self.progress_queue.get_nowait()
                if evt[0] == "progress":
                    idx, total = evt[1], evt[2]
                    pct = int((idx/total)*100) if total else 0
                    self.progress['maximum'] = total
                    self.progress['value'] = idx
                elif evt[0] == "done":
                    self.progress['value'] = evt[1]
                    self.start_btn.config(state="normal")
                    self.cancel_btn.config(state="disabled")
                    self.logbox.insert("end", "Proceso terminado.\n")
                    self.logbox.see("end")
        except queue.Empty:
            pass

        # re-schedule
        self.root.after(200, self._poll_queues)

# ---------- Toplevel folder browser implementations ----------
def open_folder_browser(parent, title="Seleccionar carpeta", start_path=None, only_directories=True):
    """
    Muestra una Toplevel con Treeview para explorar carpetas.
    Devuelve la ruta seleccionada o None.
    """
    sel = {"path": None}
    if not start_path or not os.path.exists(start_path):
        start_path = os.path.expanduser("~")

    top = Toplevel(parent)
    top.title(title)
    top.geometry("640x420")
    top.transient(parent)
    top.grab_set()

    # Treeview
    tree = ttk.Treeview(top)
    tree.heading("#0", text="Carpetas", anchor="w")
    tree.pack(fill="both", expand=True, padx=8, pady=8)

    # Helper to insert child nodes lazily
    def insert_node(parent_node, fullpath):
        # add a dummy child if directory has subdirs
        node = tree.insert(parent_node, "end", text=os.path.basename(fullpath) or fullpath, open=False, values=[fullpath])
        try:
            # detect subdirectories
            with os.scandir(fullpath) as it:
                has_sub = any(entry.is_dir() for entry in it)
            if has_sub:
                tree.insert(node, "end", text="(loading)", values=["__dummy__"])
        except Exception:
            pass
        return node

    # populate roots
    roots = get_roots()
    for r in roots:
        insert_node("", r)

    # expand on demand
    def on_open(event):
        node = tree.focus()
        children = tree.get_children(node)
        # if first child is dummy, replace with actual
        if children:
            first = children[0]
            vals = tree.item(first, "values")
            if vals and vals[0] == "__dummy__":
                # remove dummy
                tree.delete(first)
                # get fullpath of node
                parts = []
                cur = node
                while cur:
                    txt = tree.item(cur, "text")
                    parts.insert(0, txt)
                    cur = tree.parent(cur)
                # reconstruct fullpath from values stored earlier (we stored fullpath in values for root only)
                # Instead, we can walk down: compute fullpath by reading 'values' of node if present
                # We'll retrieve fullpath via a helper that extracts concatenated text up to root
                # Simpler: store full path in 'values' when inserting - modify insert_node to set values accordingly.
                # We'll modify approach: when inserting, set values=[fullpath]; reconstruct from that.
                node_path = tree.set(node, "#1") if "#1" in tree.set(node) else None
                # but many ttk implementations vary; safer approach: compute path by walking and joining names
                cur = node
                comps = []
                while cur:
                    txt = tree.item(cur, "text")
                    comps.insert(0, txt)
                    cur = tree.parent(cur)
                # find first root that matches one of get_roots()
                possible = None
                for r in roots:
                    if comps and (comps[0] == os.path.basename(r) or comps[0] == r):
                        possible = r
                        break
                if possible:
                    # build full path:
                    full = possible
                    for seg in comps[1:]:
                        full = os.path.join(full, seg)
                else:
                    # fallback: reconstruct via iterating parent values (if available)
                    full = ""
                    cur = node
                    stack = []
                    while cur:
                        stack.insert(0, tree.item(cur, "text"))
                        cur = tree.parent(cur)
                    # join carefully
                    if platform.system().lower() == "windows" and stack:
                        # first component may be 'C:\' root with backslash
                        full = stack[0]
                        for s in stack[1:]:
                            full = os.path.join(full, s)
                    else:
                        full = os.path.sep.join(stack)
                # list subdirs
                try:
                    subs = list_dir(full)
                    for s in subs:
                        insert_node(node, os.path.join(full, s))
                except Exception:
                    pass

    # A more reliable variant: we store full path in node 'values' when inserting.
    # To keep code simpler, re-build the tree with stored fullpath explicitly:
    tree.delete(*tree.get_children())
    def populate_with_fullpaths():
        for r in roots:
            node = tree.insert("", "end", text=r, values=(r,))
            _populate_children(node, r)
    def _populate_children(parent_node, fullpath):
        try:
            subdirs = list_dir(fullpath)
            for sd in subdirs:
                sd_full = os.path.join(fullpath, sd)
                child = tree.insert(parent_node, "end", text=sd, values=(sd_full,))
                # check if child has children
                if list_dir(sd_full):
                    tree.insert(child, "end", text="(dummy)", values=("__dummy__",))
        except Exception:
            pass
    populate_with_fullpaths()

    def on_expand(event):
        node = tree.focus()
        # check for dummy child
        children = tree.get_children(node)
        for ch in children:
            vals = tree.item(ch, "values")
            if vals and vals[0] == "__dummy__":
                # remove dummy and load real children
                tree.delete(ch)
                parent_full = tree.item(node, "values")[0]
                try:
                    for sd in list_dir(parent_full):
                        sd_full = os.path.join(parent_full, sd)
                        child = tree.insert(node, "end", text=sd, values=(sd_full,))
                        if list_dir(sd_full):
                            tree.insert(child, "end", text="(dummy)", values=("__dummy__",))
                except Exception:
                    pass

    tree.bind("<<TreeviewOpen>>", on_expand)

    # Buttons
    btn_frame = ttk.Frame(top)
    btn_frame.pack(fill="x", padx=8, pady=6)
    selected_label = Label(btn_frame, text="Seleccionado: ")
    selected_label.pack(side="left", padx=6)

    def update_selected_label(evt=None):
        node = tree.focus()
        if node:
            vals = tree.item(node, "values")
            if vals:
                selected_label.config(text=f"Seleccionado: {vals[0]}")
            else:
                selected_label.config(text="Seleccionado: ")
    tree.bind("<<TreeviewSelect>>", update_selected_label)

    def choose_folder():
        node = tree.focus()
        if not node:
            messagebox.showwarning("Atención", "Selecciona una carpeta en el árbol.")
            return
        vals = tree.item(node, "values")
        if not vals:
            messagebox.showwarning("Atención", "Ruta no válida.")
            return
        sel["path"] = vals[0]
        top.destroy()

    def go_up():
        node = tree.focus()
        if not node:
            return
        parent = tree.parent(node)
        if parent:
            tree.selection_set(parent)
            tree.focus(parent)
            tree.see(parent)

    Button(btn_frame, text="Subir (Up)", command=go_up).pack(side="right", padx=4)
    Button(btn_frame, text="Seleccionar carpeta", command=choose_folder).pack(side="right", padx=4)
    Button(btn_frame, text="Cancelar", command=top.destroy).pack(side="right", padx=4)

    # wait modal
    parent.wait_window(top)
    return sel["path"]

def open_save_browser(parent, title="Guardar archivo", start_path=None, suggested_name="extracted_data.xlsx"):
    """
    Muestra una Toplevel que permite seleccionar carpeta via árbol + ingresar nombre de archivo.
    Devuelve (folder, filename) o (None, None).
    """
    res = {"folder": None, "filename": None}
    if not start_path or not os.path.exists(start_path):
        start_path = os.path.expanduser("~")

    top = Toplevel(parent)
    top.title(title)
    top.geometry("700x480")
    top.transient(parent)
    top.grab_set()

    # Treeview for folders (left)
    left_frame = ttk.Frame(top)
    left_frame.pack(side="left", fill="both", expand=True, padx=6, pady=6)
    tree = ttk.Treeview(left_frame)
    tree.heading("#0", text="Carpetas", anchor="w")
    tree.pack(fill="both", expand=True)

    roots = get_roots()
    def _populate():
        for r in roots:
            n = tree.insert("", "end", text=r, values=(r,))
            try:
                for sd in list_dir(r):
                    sd_full = os.path.join(r, sd)
                    child = tree.insert(n, "end", text=sd, values=(sd_full,))
                    if list_dir(sd_full):
                        tree.insert(child, "end", text="(dummy)", values=("__dummy__",))
            except Exception:
                pass
    _populate()

    def on_expand(event):
        node = tree.focus()
        for ch in tree.get_children(node):
            vals = tree.item(ch, "values")
            if vals and vals[0] == "__dummy__":
                tree.delete(ch)
                parent_full = tree.item(node, "values")[0]
                try:
                    for sd in list_dir(parent_full):
                        sd_full = os.path.join(parent_full, sd)
                        child = tree.insert(node, "end", text=sd, values=(sd_full,))
                        if list_dir(sd_full):
                            tree.insert(child, "end", text="(dummy)", values=("__dummy__",))
                except Exception:
                    pass
    tree.bind("<<TreeviewOpen>>", on_expand)

    right_frame = ttk.Frame(top)
    right_frame.pack(side="right", fill="y", padx=6, pady=6)

    Label(right_frame, text="Carpeta seleccionada:").pack(anchor="w")
    selected_var = StringVar(value=start_path)
    Label(right_frame, textvariable=selected_var, wraplength=250).pack(anchor="w", pady=(0,8))

    def update_selected(evt=None):
        node = tree.focus()
        if node:
            vals = tree.item(node, "values")
            if vals:
                selected_var.set(vals[0])

    tree.bind("<<TreeviewSelect>>", update_selected)

    Label(right_frame, text="Nombre archivo:").pack(anchor="w", pady=(6,0))
    filename_var = StringVar(value=suggested_name)
    Entry(right_frame, textvariable=filename_var, width=40).pack(anchor="w", pady=(0,8))

    def do_save():
        folder = selected_var.get().strip()
        filename = filename_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Selecciona una carpeta válida para guardar.")
            return
        if not filename:
            messagebox.showerror("Error", "Escribe un nombre de archivo válido.")
            return
        # ensure extension
        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"
        res["folder"] = folder
        res["filename"] = filename
        top.destroy()

    btns = ttk.Frame(right_frame)
    btns.pack(anchor="e", pady=12)
    Button(btns, text="Guardar", command=do_save).pack(side="right", padx=6)
    Button(btns, text="Cancelar", command=top.destroy).pack(side="right", padx=6)

    parent.wait_window(top)
    return res["folder"], res["filename"]

# ---------- Main ----------
def main():
    root = Tk()
    app = OCRGui(root)
    root.mainloop()

if __name__ == "__main__":
    main()
