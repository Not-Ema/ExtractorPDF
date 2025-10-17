import os
import subprocess
import shutil
import zipfile

APP_NAME = "PDFExtractor"
MAIN_SCRIPT = "extract_pdfs_to_excel.py"
OUTPUT_DIR = "dist"
TESSERACT_DIR = "tesseract"

# Limpiar carpeta de salida
if os.path.exists(OUTPUT_DIR):
    shutil.rmtree(OUTPUT_DIR)

# Comando de Nuitka
cmd = [
    "python", "-m", "nuitka",
    "--standalone",
    "--onefile",
    "--windows-disable-console",
    "--include-data-dir=tesseract=tesseract",
    f"--output-dir={OUTPUT_DIR}",
    f"--output-filename={APP_NAME}.exe",
    MAIN_SCRIPT
]

# Ejecutar
subprocess.run(cmd, check=True)

# Crear ZIP para GitHub Releases
zip_path = f"{APP_NAME}_portable.zip"
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(OUTPUT_DIR):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = os.path.relpath(file_path, OUTPUT_DIR)
            zf.write(file_path, arcname)

print(f"✅ ZIP creado: {zip_path}")