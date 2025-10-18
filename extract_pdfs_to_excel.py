import os, sys, platform, subprocess, tkinter.messagebox as mb

# 1. Evitar ventana negra al ejecutar Tesseract (Windows)
if platform.system() == "Windows":
    _orig = subprocess.Popen
    # use creationflags=CREATE_NO_WINDOW para evitar consola
    CREATE_NO_WINDOW = 0x08000000
    subprocess.Popen = lambda *a, **k: _orig(*a, **{**k, "creationflags": k.get("creationflags", CREATE_NO_WINDOW)})

# ---------- Determinar carpeta base (donde está el .exe o el .py) ----------
def get_base_dir():
    """
    Devuelve la carpeta donde se encuentra el ejecutable en runtime:
     - si está 'frozen' (compilado) -> carpeta del ejecutable
     - si está en desarrollo -> carpeta del script
    """
    if getattr(sys, "frozen", False):
        # .exe o carpeta standalone
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

BASE = get_base_dir()

# ---------- Buscar tesseract.exe en ubicaciones razonables ----------
def find_tesseract_exe(base):
    candidates = [
        os.path.join(base, "tesseract", "tesseract.exe"),   # ./tesseract/tesseract.exe
        os.path.join(base, "tesseract.exe"),               # ./tesseract.exe
        os.path.join(base, "bin", "tesseract.exe"),        # ./bin/tesseract.exe (por si usas estructura distinta)
    ]

    # si está dentro de un bundle temporal (onefile extract), también revisar sys._MEIPASS si existe
    if hasattr(sys, "_MEIPASS"):
        meipass = getattr(sys, "_MEIPASS")
        candidates.extend([
            os.path.join(meipass, "tesseract", "tesseract.exe"),
            os.path.join(meipass, "tesseract.exe")
        ])

    # comprobar PATH como último recurso
    for c in candidates:
        c = os.path.normpath(c)
        if os.path.isfile(c):
            return c

    # buscar en PATH
    for path_dir in os.environ.get("PATH", "").split(os.pathsep):
        p = os.path.join(path_dir, "tesseract.exe")
        if os.path.isfile(p):
            return p

    return None

TESSERACT_EXE = find_tesseract_exe(BASE)

if not TESSERACT_EXE:
    mb.showerror("Falta Tesseract", f"No se ha encontrado tesseract.exe en ninguna ubicación esperada.\n\nBuscado en:\n• {BASE}\\tesseract\\tesseract.exe\n• {BASE}\\tesseract.exe\n• PATH\n\nColoca la carpeta 'tesseract' junto al .exe o instala tesseract en el PATH.")
    sys.exit(1)

# ------------- Asignar comando a pytesseract (después de importar pytesseract) -------------
import pytesseract
pytesseract.pytesseract.tesseract_cmd = os.path.normpath(TESSERACT_EXE)


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
from tkinter import Tk, StringVar, IntVar, BooleanVar, Toplevel, filedialog, messagebox, ttk, scrolledtext, Label, Button, Entry, Checkbutton
import tkinter as tk
from datetime import datetime
from pdf2image import convert_from_path
import pytesseract
from PIL import Image, ImageFilter, ImageOps
import pandas as pd

# ---------- Default CONFIG ----------
DEFAULT_DPI = 600
DEFAULT_LANG = "spa"
# ------------------------------------
# Ruta relativa al ejecutable portable
import subprocess
import pytesseract
import os
import platform

# Ruta al tesseract portable

# ---------- Helper: limpiar nombre de cliente ----------
# ---------- Helper: limpiar nombre de cliente (mejorado) ----------
def clean_client_name(name: str) -> str:
    """
    Limpia un nombre extraído por OCR:
      - elimina contenido entre paréntesis/ corchetes
      - quita sufijos artefacto cortos (RI, IO, IA, AI, TAI, OI, etc.)
      - elimina tokens de 1 letra (ej. 'N', 'A') y tokens de 2-3 letras sospechosos,
        salvo partículas válidas (DE, DEL, LA, SAN, MC, Y, etc.)
      - elimina palabras tipo LEADER/REP y números pegados
    """
    if not name:
        return None

    s = str(name)

    # quitar contenido entre corchetes o paréntesis y guiones largos
    s = re.sub(r"[\[\(].*?[\]\)]", " ", s)
    s = re.sub(r"[—–\-]+", " ", s)

    # normalizar espacios y pasar a mayúsculas
    s = re.sub(r'\s+', ' ', s).strip().upper()

    # eliminar patrones "ROLE + número" (ej: LEADER 1051823433)
    s = re.sub(r'\b(?:LEADER|REP|REPRESENTANTE|CONTACTO|CONTACT|AGENTE|OPERADOR|TELEFONO|TEL|CEL|CELULAR|MOVIL|MOV)\b\s*\d{3,}\b', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(?:LEADER|REP|REPRESENTANTE|CONTACTO|AGENTE|OPERADOR)\b', ' ', s, flags=re.IGNORECASE)

    # tokenizar solo letras (evitamos arrastrar dígitos)
    tokens = [t for t in re.findall(r"[A-ZÁÉÍÓÚÑ]+", s)]
    if not tokens:
        return None

    # partículas a conservar
    KEEP = {"DE", "DEL", "LA", "LAS", "LOS", "SAN", "SANTA", "MC", "VON", "Y", "DA", "DI", "ST", "SANTO"}

    # blacklist corta (artefactos comunes). Añadí TAI, OI y variantes que mencionaste.
    BLACKLIST_SHORT = {
        "IT","IM","II","I","R","M","S","T","OT","XT","IV","VI",
        "RI","IO","IA","AI","IN","ON","EN","AN","NA","N","A",
        "TAI","OI","IO","OI","AI","IA","RI","RI"  # repetidos por seguridad (no dañan)
    }

    cleaned = []
    for t in tokens:
        # conservar partículas válidas
        if t in KEEP:
            cleaned.append(t)
            continue
        # eliminar tokens de longitud 1
        if len(t) == 1:
            continue
        # eliminar tokens cortos que están en la blacklist
        if 2 <= len(t) <= 3 and t in BLACKLIST_SHORT:
            continue
        # heurística extra: si son 2 letras y ambas consonantes poco probables en nombre => eliminar
        if len(t) == 2 and t not in KEEP:
            if re.match(r'^[BCDFGHJKLMNPQRSTVWXYZ]{2}$', t):
                continue
            # si no está en blacklist y no es doble consonante, lo dejamos (ej: 'LU' podría ser parte de nombre)
        # para tokens >= 4 ó 3 que no están en blacklist, los conservamos
        cleaned.append(t)

    # quitar sufijos finales cortos que hayan quedado
    while cleaned and len(cleaned[-1]) <= 2 and cleaned[-1] not in KEEP:
        cleaned.pop()

    result = " ".join(cleaned).strip()
    return result if result else None



# ---------- OCR + extracción (adaptado) ----------
# ---------- OCR + extracción (adaptado y mejorado) ----------
def image_preprocess(img: Image.Image, upscale_if_small=True) -> Image.Image:
    """
    Mejora la calidad de imagen antes del OCR:
      - Convierte a escala de grises.
      - Aumenta la resolución si es pequeña.
      - Aplica contraste y nitidez.
      - Binariza (umbral adaptativo).
    """
    # Convertir a escala de grises
    img = img.convert("L")

    # Aumentar tamaño si la imagen es pequeña (para evitar letras rotas)
    w, h = img.size
    if upscale_if_small and w < 4000:
        factor = 4000 / w
        img = img.resize((int(w * factor), int(h * factor)), Image.LANCZOS)

    # Aumentar contraste y claridad
    img = ImageOps.autocontrast(img)
    from PIL import ImageEnhance
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(3.0)

    enhancer = ImageEnhance.Sharpness(img)
    img = enhancer.enhance(3.0)

    # Aumentar contraste y claridad
    # img = ImageOps.autocontrast(img)
    # Filtro de suavizado mediano (elimina puntos de ruido)
    img = img.filter(ImageFilter.MedianFilter(size=3))

    # Aplicar umbral binario para texto nítido (binarización manual)
    # Aumenta la separación texto-fondo
    img = img.point(lambda x: 0 if x < 120 else 255, '1')

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

        # --- Cliente: asume que está en MAYÚSCULAS y tras "Cliente:" (mejor extracción multilinea)
        # --- Cliente: extracción multiline mejorada y limpieza
        # --- Cliente: extracción multilínea mejorada y limpieza robusta ---
    data["Cliente"] = None

    # Buscar "Cliente:" en el texto
    m = re.search(r"Cliente[:\s]*(.+)", txt, flags=re.IGNORECASE)
    if m:
        # Tomar la parte después de "Cliente:" (solo la primera línea)
                # Tomar la parte después de "Cliente:" (solo la primera línea)
        base_line = m.group(1).splitlines()[0].strip()

        # Eliminar todo desde "IDENTIFICACION" en adelante (misma línea)
        base_line = re.split(r'\bIDENTIFICACI[oó]N\b', base_line, flags=re.IGNORECASE)[0].strip()

        # Extraer tokens de letras
        tokens = [t.upper() for t in re.findall(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]+", base_line)]

        # Palabras que no deben considerarse parte del nombre
        stop_keywords = {
            "CONTRATO", "PBX", "FAX", "DIR", "DIRECCION", "NO", "NO.",
            "REF", "REFERENCIA", "PAGO", "LÍNEA", "LINEA", "TIPO",
            "TOTAL", "AV", "KR", "CL", "C"
        }

        # Filtrar tokens iniciales (solo letras, sin palabras de control)
        tokens = [t for t in tokens if t not in stop_keywords and len(t) > 1 and t not in {"N", "A"}]
        tokens = [t for t in tokens if t not in {"RI", "IO", "IA", "AI", "IN", "AN", "NA"}]

        # Leer líneas siguientes (máximo 6) por si el nombre continúa en otra línea
        following = txt[m.end():].splitlines()
        look = 0
        for ln in following:
            if look >= 6:
                break
            look += 1
            ln_strip = ln.strip()
            if not ln_strip:
                continue
            up = ln_strip.upper()

            # Si detecta otra sección o campo, detener
            if any(up.startswith(k) for k in (
                "CONTRATO", "PBX", "FAX", "NO.", "NO ", "PAGO",
                "LÍNEA", "LINEA", "DIR", "DIREC", "REFERENCIA",
                "TIPO", "TOTAL"
            )):
                break

            # Si la línea contiene rol (leader, representante, etc.), cortar antes
            if re.search(r'\b(LEADER|REP|REPRESENTANTE|CONTACTO|AGENTE|OPERADOR)\b', up, flags=re.IGNORECASE):
                before_role = re.split(r'\b(LEADER|REP|REPRESENTANTE|CONTACTO|AGENTE|OPERADOR)\b', up, flags=re.IGNORECASE)[0]
                ln_name_tokens = [t.upper() for t in re.findall(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]{3,}", before_role)]
                ln_name_tokens = [t for t in ln_name_tokens if t not in stop_keywords]
                tokens.extend(ln_name_tokens)
                break

            # Si contiene un número largo (teléfono o identificación), cortar
            if re.search(r'\d{5,}', ln_strip):
                ln_name_tokens = [t.upper() for t in re.findall(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]{3,}", ln_strip)]
                ln_name_tokens = [t for t in ln_name_tokens if t not in stop_keywords]
                tokens.extend(ln_name_tokens)
                break

            # Si parece una línea de nombre (predomina texto alfabético)
            ln_tokens = [t.upper() for t in re.findall(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]+", ln_strip)]
            if not ln_tokens:
                continue
            ln_filtered = [t for t in ln_tokens if t not in stop_keywords and len(t) > 1 and t not in {"N", "A"}]
            ln_filtered = [t for t in ln_filtered if t not in {"RI", "IO", "IA", "AI", "IN", "AN", "NA"}]
            tokens.extend(ln_filtered)

        # Unir tokens para formar el nombre bruto
        raw_joined = " ".join(tokens).strip()

        # --- Limpieza final del nombre (roles, cortos, etc.) ---
        def clean_client_name(name: str) -> str:
            if not name:
                return None

            s = str(name)
            s = re.sub(r"[\[\(].*?[\]\)]", " ", s)
            s = re.sub(r"[—–\-]+", " ", s)
            s = re.sub(r'\s+', ' ', s).strip().upper()

            # Eliminar patrones "ROL + número"
            s = re.sub(
                r'\b(?:LEADER|REP|REPRESENTANTE|CONTACTO|AGENTE|OPERADOR|TELEFONO|TEL|CEL|CELULAR|MOVIL|MOV)\b\s*\d{3,}\b',
                ' ', s, flags=re.IGNORECASE
            )
            s = re.sub(
                r'\b(?:LEADER|REP|REPRESENTANTE|CONTACTO|AGENTE|OPERADOR)\b',
                ' ', s, flags=re.IGNORECASE
            )

            # Tokenizar
            tokens = [t for t in re.findall(r"[A-ZÁÉÍÓÚÑ]+", s)]
            if not tokens:
                return None

            KEEP = {"DE", "DEL", "LA", "LAS", "LOS", "SAN", "SANTA", "MC", "VON", "Y", "DA", "DI", "ST", "SANTO"}
            BLACKLIST_SHORT = {
                "IT","IM","II","I","R","M","S","T","OT","XT","IV","VI",
                "RI","IO","IA","AI","IN","ON","EN","AN","NA","N","A",
                "TAI","OI","IO","OI","AI","IA","RI","RI"  # repetidos por seguridad (no dañan)
                }

            cleaned = []
            for t in tokens:
                if t in KEEP:
                    cleaned.append(t)
                    continue
                if len(t) == 1 or t in BLACKLIST_SHORT:
                    continue
                if 2 <= len(t) <= 3 and t in BLACKLIST_SHORT:
                    continue
                if len(t) == 2 and t not in KEEP:
                    if re.match(r'^[BCDFGHJKLMNPQRSTVWXYZ]{2}$', t):
                        continue
                cleaned.append(t)

            while cleaned and len(cleaned[-1]) <= 2 and cleaned[-1] not in KEEP:
                cleaned.pop()

            result = " ".join(cleaned).strip()
            return result if result else None

        # Aplicar limpieza
        cleaned = clean_client_name(raw_joined) if raw_joined else None
        data["Cliente"] = cleaned if cleaned else (raw_joined if raw_joined else None)



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
        m = re.search(r"((?:KR|CL|AV|C[^\n]{1,30}|[A-Z]{2,5}\s*\d{1,3})[^\n]{0,60})", txt, flags=re.IGNORECASE)
    if m:
        data["DirCliente"] = m.group(1).strip().split("\n")[0].strip()

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

        # Si el tipo de cupón es "CA", dejar NoSolicitud en blanco
        if data["TipoCupon"] == "CA":
            data["NoSolicitud"] = None

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
    # --- Código de barra: solo líneas con formato GS1-128 válido
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    barcode_line = None
    for ln in reversed(lines):
        # Patrón GS1-128: (AI)valor repetido, con AI de 2+ dígitos
        if re.fullmatch(r'(\(\d{2,4}\)\d+)+', ln):
            barcode_line = ln
            break
    if barcode_line:
        data["CodigoBarraRaw"] = barcode_line
        data["CodigoBarraLimpio"] = clean_barcode(barcode_line)
     # --- Importe desde AI 3900
    m3900 = re.search(r'\(3900\)(\d{4,12})', barcode_line)
    if m3900:
        # sin decimales (entero)
        data["ValorAPagar"] = str(int(m3900.group(1)))

    # ---------------- REEMPLAZAR LA SECCIÓN NoSolicitud POR ESTE BLOQUE ----------------
    # ------------------ NoSolicitud (siguiendo tu regla: línea No. Ref -> 1º FAX, 2º NoSolicitud) ------------------
    data["NoSolicitud"] = None
    found = None

    # Buscar la línea que contiene "No. Ref" (variantes)
    m_ref_line = re.search(r"^(.*No\.?\s*Ref\.?.*)$", txt, flags=re.IGNORECASE | re.MULTILINE)
    if m_ref_line:
        line = m_ref_line.group(1)
        # extraer todos los números largos (6+ dígitos) de esa línea, en el orden que aparecen
        nums = [mo.group(0) for mo in re.finditer(r"[0-9]{6,15}", line)]
        if len(nums) >= 2:
            # según tu regla: el primero es FAX, el segundo es NoSolicitud
            candidate = nums[1]
            # evitar devolver la misma que la identificación
            if data.get("Identificacion") and re.sub(r"\D","",candidate) == re.sub(r"\D","",str(data["Identificacion"])):
                # si coincide con la identificación, intentar tomar el tercero (si existe)
                if len(nums) >= 3:
                    candidate = nums[2]
                else:
                    candidate = None
            found = candidate
        elif len(nums) == 1:
            # Solo hay un número largo en la línea. Intentar buscar contexto cercano
            #  a) si en la misma línea aparece 'FAX' justo antes del número, ese será fax -> buscar número siguiente en ventana del texto
            # (Tomamos ±200 caracteres alrededor de la posición de la línea en el texto)
            pos = m_ref_line.start(1)
            window = txt[max(0, pos-200): pos+200]
            all_nums_window = [mo.group(0) for mo in re.finditer(r"[0-9]{6,15}", window)]
            # si hay al menos 2 en la ventana, preferimos el que no sea el primero (asumiendo fax primero)
            if len(all_nums_window) >= 2:
                candidate = all_nums_window[1]
                if data.get("Identificacion") and re.sub(r"\D","",candidate) == re.sub(r"\D","",str(data["Identificacion"])):
                    if len(all_nums_window) >= 3:
                        candidate = all_nums_window[2]
                    else:
                        candidate = None
                found = candidate
            else:
                # fallback: no se pudo determinar, dejamos None (mejor vacío que equivocarnos)
                found = None
    else:
        # Si no hay línea "No. Ref", fallback conservador:
        # buscar la palabra "Solicitud" y tomar el número más cercano a la derecha que no sea la ident.
        mpos = re.search(r"\bSolicit(?:ud|ion)\b", txt, flags=re.IGNORECASE)
        if mpos:
            pos = mpos.start()
            # buscar números en un rango a la derecha
            window = txt[pos: pos+300]
            nums = [mo.group(0) for mo in re.finditer(r"[0-9]{6,15}", window)]
            # preferir el primer número que no coincida con Identificacion
            candidate = None
            for n in nums:
                if data.get("Identificacion") and re.sub(r"\D","",n) == re.sub(r"\D","",str(data["Identificacion"])):
                    continue
                candidate = n
                break
            found = candidate

    # asegurarse de no devolver la identificación por error
    if found and data.get("Identificacion") and re.sub(r"\D","",found) == re.sub(r"\D","",str(data["Identificacion"])):
        found = None

    data["NoSolicitud"] = re.sub(r"\D", "", found) if found else None

    return data


import pdfplumber  # colocarlo al inicio junto con tus imports

def extract_text_from_pdf(pdf_path, dpi=600, lang='spa', tesseract_config="--psm 6",
                          save_ocr_text=False, ocr_text_dir=None, logger=None,
                          selectable_text_min_chars=50):
    """
    Extrae texto de un PDF intentando primero obtener texto seleccionable (pdfplumber).
    Si no se detecta texto suficiente (menos de selectable_text_min_chars), hace OCR
    usando pdf2image + pytesseract y devuelve ese texto.
    Parámetros:
      - pdf_path: ruta al PDF
      - dpi: resolución para convertir páginas a imagen (si OCR requerido)
      - lang: idiomas para tesseract (ej: 'spa')
      - tesseract_config: configuración de tesseract (ej: "--psm 6")
      - save_ocr_text: si True guarda .txt con el texto OCR
      - ocr_text_dir: carpeta donde guardar .txt
      - selectable_text_min_chars: mínimo de caracteres para considerar "texto seleccionable útil"
    """
    # 1) Intentar texto seleccionable con pdfplumber
    try:
        text_pages = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # extrae texto de la página; strip() para eliminar espacios extra
                t = page.extract_text()
                if t:
                    text_pages.append(t)
        selectable_text = "\n\n".join(text_pages).strip()
        # si hay texto seleccionable suficiente, devolverlo directamente
        if selectable_text and len(re.sub(r'\s+', '', selectable_text)) >= selectable_text_min_chars:
            if logger:
                logger(f"Usando texto seleccionable de: {os.path.basename(pdf_path)}")
            # opcional: normalizar saltos de línea múltiples
            return selectable_text
    except Exception as e:
        # si falla pdfplumber (archivo raro), seguimos a OCR sin interrumpir
        if logger:
            logger(f"pdfplumber fallo para {os.path.basename(pdf_path)}: {e}. Se intentará OCR.")

    # 2) Si no hay texto seleccionable suficiente -> usar OCR (imagen)
    # Convertir páginas a imágenes
    pages = convert_from_path(pdf_path, dpi=dpi)
    texts = []
    for _i, page in enumerate(pages):
        # aplicar preprocesado (tu función image_preprocess)
        try:
            img = image_preprocess(page)
            text = pytesseract.image_to_string(img, lang=lang, config=tesseract_config)
            texts.append(text)
        except Exception as e:
            # si falla en una página, seguir con las demás
            if logger:
                logger(f"OCR fallo en página {_i+1} de {os.path.basename(pdf_path)}: {e}")
    full_text = "\n\n".join(texts)

    # 3) Guardar .txt si se solicita
    if save_ocr_text and ocr_text_dir:
        try:
            os.makedirs(ocr_text_dir, exist_ok=True)
            fn = os.path.splitext(os.path.basename(pdf_path))[0] + ".txt"
            with open(os.path.join(ocr_text_dir, fn), "w", encoding="utf-8") as f:
                f.write(full_text)
        except Exception as e:
            if logger:
                logger(f"No se pudo guardar OCR .txt para {pdf_path}: {e}")

    return full_text


# ---------- Worker: procesa una carpeta ----------
def process_all_pdfs(input_folder, output_excel, dpi, lang, tesseract_cmd, save_ocr_text, ocr_text_dir, progress_queue, log_queue, stop_event):
    try:
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

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

        # Guardar Excel (append if exists)
        df = pd.DataFrame(rows)
        # ordenar columnas
        cols_order = ["_file", "Cliente", "Contrato", "Identificacion", "NoSolicitud",
                      "TipoCupon", "ValorAPagar", "NoRefPago", "DirCliente", "ValidoHasta",
                      "CodigoBarraRaw", "CodigoBarraLimpio", "error"]
        cols = [c for c in cols_order if c in df.columns] + [c for c in df.columns if c not in cols_order]
        df = df[cols]

        if os.path.exists(output_excel):
            # Load existing data and append
            existing_df = pd.read_excel(output_excel)
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            combined_df.to_excel(output_excel, index=False)
            log_queue.put(f"Datos agregados al Excel existente: {output_excel}")
        else:
            df.to_excel(output_excel, index=False)
            log_queue.put(f"Excel creado en: {output_excel}")

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
        root.title("Extractor de Datos PDF → Excel v0.1.0")
        root.geometry("800x650")
        root.configure(bg="#f8f9fa")

        # Variables
        self.input_folder = StringVar()
        self.output_file = StringVar()
        self.dpi = IntVar(value=DEFAULT_DPI)
        self.lang = StringVar(value=DEFAULT_LANG)
        self.tesseract_cmd = StringVar(value="")
        self.is_processing = False

        # Queues and thread control
        self.progress_queue = queue.Queue()
        self.log_queue = queue.Queue()
        self.worker_thread = None
        self.stop_event = threading.Event()

        # Setup UI
        self.setup_styles()
        self.create_widgets()

        # Welcome message
        self.log_message("👋 ¡Bienvenido al Extractor de Datos PDF!", "info")
        self.log_message("💡 Sigue los pasos numerados para comenzar", "info")

        # Poll queues
        root.after(200, self._poll_queues)
    def check_for_updates(self):
        self.log_message("🔍 Buscando actualizaciones...", "info")
        updater = GitHubUpdater(logger=lambda msg: self.log_message(msg, "info"))
        threading.Thread(target=updater.check, args=(self.on_update_check_complete,), daemon=True).start()

    def on_update_check_complete(self, ok, msg):
        self.root.after(0, lambda: self.log_message(msg, "success" if ok else "warning"))
        if ok:
            resp = messagebox.askyesno("🔄 Actualización disponible", f"{msg}\n\n¿Deseas aplicarla ahora?")
            if resp:
                self.log_message("⬇️ Aplicando actualización...", "info")
            # La actualización ya fue aplicada por el updater, solo reinicia
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Helvetica', 12, 'bold'), background='#f8f9fa')
        style.configure('Success.TButton', background='#28a745', foreground='white')
        style.configure('Primary.TButton', background='#007bff', foreground='white')
        style.configure('Warning.TButton', background='#ffc107', foreground='black')
        style.configure('Error.TButton', background='#dc3545', foreground='white')

    def create_input_section(self, parent):
        input_frame = ttk.LabelFrame(parent, text="📁 Paso 1: Seleccionar Carpeta de PDFs", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        self.input_label = ttk.Label(input_frame, text="📍 Ninguna carpeta seleccionada", foreground="gray", font=('Helvetica', 9))
        self.input_label.pack(anchor="w", pady=(0, 5))

        input_btn = ttk.Button(input_frame, text="🗂️ Examinar Carpeta", command=self.select_input_folder, style='Primary.TButton')
        input_btn.pack(anchor="w")

    def create_output_section(self, parent):
        output_frame = ttk.LabelFrame(parent, text="💾 Paso 2: Guardar Archivo Excel", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        self.output_label = ttk.Label(output_frame, text="📍 Ningún archivo seleccionado", foreground="gray", font=('Helvetica', 9))
        self.output_label.pack(anchor="w", pady=(0, 5))

        output_btn = ttk.Button(output_frame, text="💾 Guardar Como...", command=self.select_output_file, style='Primary.TButton')
        output_btn.pack(anchor="w")

    def create_control_section(self, parent):
        control_frame = ttk.LabelFrame(parent, text="🚀 Paso 3: Procesar Archivos", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # Botón de inicio
        self.start_button = ttk.Button(control_frame, text="▶️ Iniciar Extracción",
                                    command=self.start_processing,
                                    style='Success.TButton')
        self.start_button.pack(fill=tk.X, pady=(0, 10))

        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(control_frame, variable=self.progress_var,
                                            maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        # Frame para botones secundarios
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X)

        # Botón de actualización
        update_btn = ttk.Button(button_frame, text="🔄 Actualizar",
                                command=self.check_for_updates,
                                style='Primary.TButton')
        update_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Botones de limpiar log, ayuda y acerca de
        clear_btn = ttk.Button(button_frame, text="🧹 Limpiar Log",
                            command=self.clear_log,
                            style='Warning.TButton')
        clear_btn.pack(side=tk.LEFT, padx=(0, 5))

        help_btn = ttk.Button(button_frame, text="❓ Ayuda",
                            command=self.show_help)
        help_btn.pack(side=tk.LEFT, padx=(0, 5))

        about_btn = ttk.Button(button_frame, text="ℹ️ Acerca de",
                            command=self.show_about)
        about_btn.pack(side=tk.LEFT)
        

    def create_log_section(self, parent):
        log_frame = ttk.LabelFrame(parent, text="📋 Registro de Actividad", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.log_area = scrolledtext.ScrolledText(text_frame, height=12, wrap=tk.WORD, state='disabled', font=('Consolas', 9))
        self.log_area.pack(fill=tk.BOTH, expand=True)

        self.log_area.tag_config("info", foreground="#333333")
        self.log_area.tag_config("success", foreground="#28a745", font=('Consolas', 9, 'bold'))
        self.log_area.tag_config("error", foreground="#dc3545", font=('Consolas', 9, 'bold'))
        self.log_area.tag_config("warning", foreground="#fd7e14", font=('Consolas', 9, 'bold'))

    def create_footer(self, parent):
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill=tk.X, pady=(10, 0))

        footer_label = ttk.Label(footer_frame, text="v0.1.0 - Extractor Modular Robusto | Desarrollado con ❤️", font=('Helvetica', 8), foreground="gray")
        footer_label.pack(anchor="center")

    def select_input_folder(self):
        folder = filedialog.askdirectory(title="📁 Selecciona la carpeta con los archivos PDF")
        if folder:
            self.input_folder.set(folder)
            folder_name = os.path.basename(folder)
            pdf_count = len([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
            self.input_label.config(text=f"📁 {folder_name} ({pdf_count} PDFs)", foreground="black")
            self.log_message(f"📁 Carpeta seleccionada: {folder_name}", "success")
            self.check_ready_to_process()

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="💾 Guardar archivo como...",
            defaultextension='.xlsx',
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.output_file.set(file_path)
            file_name = os.path.basename(file_path)
            self.output_label.config(text=f"💾 {file_name}", foreground="black")
            self.log_message(f"💾 Archivo de salida: {file_name}", "success")
            self.check_ready_to_process()


    def check_ready_to_process(self):
        if self.input_folder and self.output_file:
            self.log_message("✅ ¡Listo para procesar! Haz clic en 'Iniciar Extracción'", "success")

    def select_tesseract(self):
        f = filedialog.askopenfilename(title="Seleccionar ejecutable Tesseract (si aplica)")
        if f:
            self.tesseract_cmd.set(f)

    def select_ocr_text_dir(self):
        d = filedialog.askdirectory(title="📁 Seleccionar carpeta para guardar archivos OCR (.txt)")
        if d:
            self.ocr_text_dir.set(d)


    def start_processing(self):
        input_folder = self.input_folder.get().strip()
        output_file = self.output_file.get().strip()
        if not input_folder or not output_file:
            messagebox.showwarning("⚠️ Campos incompletos", "Por favor selecciona:\n• Carpeta con PDFs\n• Archivo Excel de salida")
            return

        if self.is_processing:
            return

        self.is_processing = True
        self.start_button.config(text="⏳ Procesando...", state="disabled")
        self.progress_var.set(0)

        self.clear_log()
        self.log_message("🚀 Iniciando procesamiento...", "info")
        self.log_message(f"📂 Carpeta: {os.path.basename(input_folder)}", "info")
        self.log_message(f"📄 Archivo: {os.path.basename(output_file)}", "info")

        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()

    def process_files(self):
        try:
            input_folder = self.input_folder.get().strip()
            output_file = self.output_file.get().strip()
            pdf_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
            if not pdf_files:
                self.root.after(0, lambda: self.log_message("⚠️ No se encontraron archivos PDF", "warning"))
                self.root.after(0, lambda: self._reset_ui())
                return

            total_files = len(pdf_files)
            self.root.after(0, lambda: self.log_message(f"📄 Se encontraron {total_files} archivos PDF", "info"))

            processed_count = 0
            data_list = []
            errors_count = 0
            scan_count = 0

            def progress_callback(filename, text):
                nonlocal processed_count, errors_count, scan_count
                processed_count += 1
                progress = (processed_count / total_files) * 100

                # Extract data in the worker thread
                if text and text != "SCAN":
                    try:
                        data = extract_fields_from_text(text)
                        data["_file"] = filename
                        data_list.append(data)
                        # Schedule GUI update for success
                        self.root.after(0, lambda: self._update_progress(filename, text, progress, processed_count, total_files))
                    except Exception as e:
                        errors_count += 1
                        self.root.after(0, lambda fn=filename, p=progress, pc=processed_count, tf=total_files:
                                      self._update_progress_error(fn, str(e), p, pc, tf))
                elif text == "SCAN":
                    scan_count += 1
                    # Schedule GUI update for scan
                    self.root.after(0, lambda fn=filename, p=progress, pc=processed_count, tf=total_files:
                                  self._update_progress_scan(fn, p, pc, tf))
                else:
                    errors_count += 1
                    # Schedule GUI update for empty text
                    self.root.after(0, lambda fn=filename, p=progress, pc=processed_count, tf=total_files:
                                  self._update_progress(fn, None, p, pc, tf))

            # Process PDFs sequentially
            for pdf in pdf_files:
                if self.stop_event.is_set():
                    self.root.after(0, lambda: self.log_message("Proceso cancelado por el usuario.", "warning"))
                    break
                try:
                    text = extract_text_from_pdf(pdf, dpi=int(self.dpi.get()), lang=self.lang.get().strip() or DEFAULT_LANG,
                                               save_ocr_text=False, ocr_text_dir=None)
                    progress_callback(os.path.basename(pdf), text)
                except Exception as e:
                    progress_callback(os.path.basename(pdf), None)

            # Schedule final UI updates on main thread
            self.root.after(0, lambda: self._finalize_processing(data_list, errors_count, scan_count, total_files, output_file))

        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"❌ Error crítico: {str(e)}", "error"))
            self.root.after(0, lambda: messagebox.showerror("❌ Error", f"Error durante el procesamiento:\n\n{str(e)}"))
            self.root.after(0, lambda: self._reset_ui())

    def _finalize_processing(self, data_list, errors_count, scan_count, total_files, output_file):
        """Finalize processing and update UI on main thread."""
        try:
            if data_list:
                # Save Excel
                df = pd.DataFrame(data_list)
                cols_order = ["_file", "Cliente", "Contrato", "Identificacion", "NoSolicitud",
                              "TipoCupon", "ValorAPagar", "NoRefPago", "DirCliente", "ValidoHasta",
                              "CodigoBarraRaw", "CodigoBarraLimpio", "error"]
                cols = [c for c in cols_order if c in df.columns] + [c for c in df.columns if c not in cols_order]
                df = df[cols]

                if os.path.exists(output_file):
                    existing_df = pd.read_excel(output_file)
                    combined_df = pd.concat([existing_df, df], ignore_index=True)
                    combined_df.to_excel(output_file, index=False)
                    self.log_message(f"🎉 ¡Datos agregados exitosamente!", "success")
                else:
                    df.to_excel(output_file, index=False)
                    self.log_message(f"🎉 ¡Proceso completado exitosamente!", "success")

                self.log_message(f"📊 Total de registros procesados: {len(data_list)}", "success")

                if errors_count > 0 or scan_count > 0:
                    if errors_count > 0:
                        self.log_message(f"⚠️ Archivos con errores: {errors_count}", "warning")
                    if scan_count > 0:
                        self.log_message(f"📄 Archivos escaneados (omitidos): {scan_count}", "warning")

                success_rate = ((total_files - errors_count - scan_count) / total_files) * 100

                result = messagebox.askyesno("✅ Proceso Exitoso",
                                           f"¡Proceso completado!\n\n"
                                           f"📊 Registros procesados: {len(data_list)}\n"
                                           f"✅ Archivos exitosos: {total_files - errors_count - scan_count}\n"
                                           f"❌ Archivos con errores: {errors_count}\n"
                                           f"📄 Archivos escaneados (omitidos): {scan_count}\n"
                                           f"📈 Tasa de éxito: {success_rate:.1f}%\n"
                                           f"📄 Archivo: {os.path.basename(output_file)}\n\n"
                                           f"¿Deseas abrir la carpeta del archivo?")
                if result:
                    self.open_output_folder()
            else:
                self.log_message("❌ No se procesaron archivos exitosamente", "error")
                messagebox.showwarning("⚠️ Sin Datos",
                                     "No se pudo extraer datos de ningún archivo.\n\n"
                                     "Verifica que los PDFs contengan el formato esperado.")
        except Exception as e:
            self.log_message(f"❌ Error al guardar resultados: {str(e)}", "error")
            messagebox.showerror("❌ Error", f"Error al guardar los resultados:\n\n{str(e)}")
        finally:
            self._reset_ui()

    def _reset_ui(self):
        """Reset UI elements after processing."""
        self.is_processing = False
        self.start_button.config(text="▶️ Iniciar Extracción", state="normal")
        self.progress_var.set(0)

    def _update_progress_error(self, filename, error, progress, processed_count, total_files):
        """Update progress bar and log for errors."""
        self.progress_var.set(progress)
        self.log_message(f"📖 ({processed_count}/{total_files}) Procesando: {filename}", "info")
        self.log_message(f"   ❌ Error: {error}", "error")
        self.root.update_idletasks()

    def _update_progress_scan(self, filename, progress, processed_count, total_files):
        """Update progress bar and log for scans."""
        self.progress_var.set(progress)
        self.log_message(f"📖 ({processed_count}/{total_files}) Procesando: {filename}", "info")
        self.log_message("   📄 Archivo es un scan - procesando con OCR", "warning")
        self.root.update_idletasks()

    def _update_progress(self, filename, text, progress, processed_count, total_files):
        """Update progress bar and log from main thread."""
        self.progress_var.set(progress)
        self.log_message(f"📖 ({processed_count}/{total_files}) Procesando: {filename}", "info")
        if text:
            self.log_message("   ✅ Datos extraídos", "success")
        else:
            self.log_message("   ❌ Error extrayendo texto", "error")
        self.root.update_idletasks()

    def open_output_folder(self):
        try:
            output_file = self.output_file.get().strip()
            folder_path = os.path.dirname(output_file)
            if os.name == 'nt':
                os.startfile(folder_path)
            elif os.name == 'posix':
                os.system(f'open "{folder_path}"')
        except Exception as e:
            self.log_message(f"❌ No se pudo abrir la carpeta: {e}", "error")

    def log_message(self, message, tipo="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"

        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, formatted_message, tipo)
        self.log_area.config(state='disabled')
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_area.config(state='normal')
        self.log_area.delete('1.0', tk.END)
        self.log_area.config(state='disabled')

    def show_help(self):
        help_text = """
🔧 GUÍA DE USO:

1️⃣ Selecciona la carpeta que contiene los archivos PDF
2️⃣ Elige dónde guardar el archivo Excel de salida
3️⃣ Opcional: Activa guardar archivos OCR .txt para PDFs escaneados
4️⃣ Haz clic en 'Iniciar Extracción' y espera

📋 CAMPOS EXTRAÍDOS:
• Cliente
• Identificación
• Contrato
• Dirección
• Valor a Pagar
• No. Solicitud
• No. Rel. Pago
• Tipo de Cupón
• Válido hasta
• Código de Barras

💡 CONSEJOS:
• Los PDFs pueden ser digitales o escaneados (OCR automático)
• Se pueden procesar múltiples archivos a la vez
• Los datos se agregan al Excel/CSV existente
• Para PDFs escaneados, el OCR puede tardar más tiempo
        """
        messagebox.showinfo("❓ Ayuda", help_text)

    def select_ocr_text_dir(self):
        d = filedialog.askdirectory(title="📁 Seleccionar carpeta para guardar archivos OCR (.txt)")
        if d:
            self.ocr_text_dir.set(d)

    def show_about(self):
        about_text = """
📊 Extractor de Datos PDF → Excel v0.1.0

🎯 CARACTERÍSTICAS:
• Extracción automática de datos de PDFs
• Soporte OCR para PDFs escaneados
• Interfaz intuitiva y amigable
• Procesamiento concurrente para escalabilidad
• Barra de progreso en tiempo real
• Registro detallado de actividades

🛠️ TECNOLOGÍAS:
• Python 3.x
• pdfplumber (extracción de texto)
• pytesseract + Tesseract OCR (PDFs escaneados)
• pandas (manejo de datos)
• tkinter (interfaz gráfica)

        """
        messagebox.showinfo("ℹ️ Acerca de", about_text)

    def cancel_processing(self):
        if messagebox.askyesno("Confirmar", "¿Deseas cancelar el proceso en curso?"):
            self.stop_event.set()
            self.log_message("Cancelando... por favor espera.", "warning")

    def show_completion_dialog(self):
        output_file = self.output_file.get().strip()
        input_folder = self.input_folder.get().strip()

        top = Toplevel(self.root)
        top.title("Proceso Completado")
        top.geometry("400x150")
        top.transient(self.root)
        top.grab_set()

        Label(top, text="¿Qué desea hacer?", font=("Arial", 12)).pack(pady=10)

        def open_folder():
            import subprocess
            import platform
            try:
                if platform.system() == "Windows":
                    subprocess.run(["explorer", input_folder])
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", input_folder])
                else:  # Linux
                    subprocess.run(["xdg-open", input_folder])
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")
            top.destroy()

        def open_file():
            import subprocess
            import platform
            try:
                if platform.system() == "Windows":
                    subprocess.run(["start", output_file], shell=True)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", output_file])
                else:  # Linux
                    subprocess.run(["xdg-open", output_file])
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")
            top.destroy()

        btn_frame = ttk.Frame(top)
        btn_frame.pack(pady=20)
        Button(btn_frame, text="Abrir carpeta de PDFs", command=open_folder).pack(side="left", padx=10)
        Button(btn_frame, text="Abrir archivo Excel", command=open_file).pack(side="left", padx=10)
        Button(btn_frame, text="Cerrar", command=top.destroy).pack(side="left", padx=10)

        self.root.wait_window(top)

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
                    # Show post-processing dialog
                    self.show_completion_dialog()
        except queue.Empty:
            pass

        # re-schedule
        self.root.after(200, self._poll_queues)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="📊 Extractor de Datos PDF → Excel", font=('Helvetica', 16, 'bold'), style='Title.TLabel')
        title_label.pack(pady=(0, 20))

        self.create_input_section(main_frame)
        self.create_output_section(main_frame)
        self.create_control_section(main_frame)
        self.create_log_section(main_frame)
        self.create_footer(main_frame)

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

# ------------------------------------------------------------------
#  Mini-GitHub updater  (public domain)
# ------------------------------------------------------------------
import shutil
from pathlib import Path
import zipfile
import tempfile
import subprocess
import sys

class GitHubUpdater:
    """
    Checks & applies new releases from
    https://github.com/OWNER/REPO/releases/latest
    """
    API_URL   = "https://github.com/{owner}/{repo}/releases/latest"
    OWNER     = "Not-Ema"   # <── change here
    REPO      = "ExtractorPDF"     # <── change here
    VERSION   = "v0.1.0"             # <── current version string

    def __init__(self, logger=None):
        self.logger = logger or print

    # --------------- public API -----------------
    def check(self, on_complete=lambda ok, msg: None):
        """Run in thread.  on_complete(True/False, message)"""
        try:
            latest = self._latest_release()
            latest_tag = latest["tag_name"]
            if latest_tag.lstrip("v") <= self.VERSION.lstrip("v"):
                on_complete(False, f"Ya estás en la última versión ({self.VERSION})")
                return

            asset = self._choose_asset(latest)
            if not asset:
                on_complete(False, "No hay ZIP portable para esta plataforma")
                return

            on_complete(True, f"Nueva versión {latest_tag} disponible. Descargando…")
            self._download_and_apply(asset["browser_download_url"], latest_tag)
            on_complete(True, "Actualización aplicada. Reiniciando…")
            self._restart()

        except Exception as e:
            on_complete(False, f"Error al buscar actualización: {e}")

    # --------------- internals ------------------
    def _latest_release(self):
        url = self.API_URL.format(owner=self.OWNER, repo=self.REPO)
        import requests
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.json()

    def _choose_asset(self, release):
        """Return first ZIP asset (you can filter by name)."""
        for a in release.get("assets", []):
            if a["name"].lower().endswith(".zip"):
                return a
        return None

    def _download_and_apply(self, zip_url, new_tag):
        import requests
        base = Path(sys.executable if getattr(sys, 'frozen', False) else __file__).resolve().parent
        backup = base / "backup_before_update"
        if backup.exists():
            shutil.rmtree(backup)

        # Descargar ZIP temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
            self.logger(f"Descargando {zip_url}")
            with requests.get(zip_url, stream=True, timeout=30) as r:
                r.raise_for_status()
                for chunk in r.iter_content(chunk_size=8192):
                    tmp.write(chunk)
            tmp.flush()

            # Backup de la carpeta actual (sin el ZIP)
            self.logger("Haciendo backup…")
            shutil.copytree(base, backup, ignore=shutil.ignore_patterns("backup_before_update"))

            # Extraer encima
            self.logger("Extrayendo actualización…")
            with zipfile.ZipFile(tmp.name, 'r') as zf:
                zf.extractall(base)

        # Guardar versión nueva
        (base / "version.txt").write_text(new_tag, encoding="utf8")
        self.logger("Actualización lista.")

    def _restart(self):
        """Re-launch the new EXE and exit current."""
        exe = sys.executable
        self.logger(f"Reiniciando {exe}")
        subprocess.Popen([exe], cwd=Path(exe).parent)
        os._exit(0)

# ---------- Main ----------
def main():
    root = Tk()
    app = OCRGui(root)
    root.mainloop()

if __name__ == "__main__":
    main()
