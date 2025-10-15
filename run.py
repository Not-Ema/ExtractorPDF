#!/usr/bin/env python3
"""
Entry point for the PDF Extractor application.
"""
import tkinter as tk
import sys
import json
import os
import traceback

# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def check_dependencies():
    """Check if all required dependencies are installed."""
    missing_deps = []
    required_packages = {
        'pdfplumber': 'pdfplumber',
        'pandas': 'pandas',
        'openpyxl': 'openpyxl',
        'numpy': 'numpy',
        'PIL': 'Pillow'
    }
    
    for module, package in required_packages.items():
        try:
            __import__(module)
        except ImportError:
            missing_deps.append(package)
    
    if missing_deps:
        root = tk.Tk()
        root.withdraw()
        from tkinter import messagebox
        deps_str = ' '.join(missing_deps)
        messagebox.showerror(
            "Librerías Faltantes", 
            f"Dependencias no encontradas:\n{', '.join(missing_deps)}\n\n"
            f"Ejecuta:\npip install {deps_str}"
        )
        sys.exit(1)

def load_config():
    """Load configuration file."""
    config_path = 'config/config.json'
    if not os.path.exists(config_path):
        # Create default config if it doesn't exist
        os.makedirs('config', exist_ok=True)
        default_config = {
            "extraction_patterns": {
                "bloque_cliente": "Cliente:(.*?)Contrato",
                "bloque_contrato": "Contrato(.*?)No\\.\\s*Solicitud",
                "cliente": "([A-ZÑÁÉÍÓÚ\\.\\-\\s]{5,}?)(?=\\s*Identificaci[\\s\\S]{0,3}?n)",
                "identificacion": "Identificaci[\\s\\S]{0,3}?n\\s*:?\\s*(\\d{6,})",
                "contrato": "Contrato\\s*(\\d{3,})",
                "direccion": "Dir\\.?\\s*Cliente:\\s*(.+?)\\s*(?:FAX:|L[íi]nea|Valor\\s*a\\s*pagar|Tipo\\s*de|No\\.|$)",
                "no_solicitud": "No\\.?\\s*Solic(?:itud)?\\s*:?\\s*(\\d{6,})",
                "no_rel_pago": "No\\.?\\s*Ref\\.?[\\s\\S]{0,20}?(\\d{6,})",
                "tipo_cupon": "Tipo\\s*de\\s*([A-Z]{1,4})",
                "tipo_cupon_alt1": "Cup[óo]n:\\s*([A-Z]{1,4})",
                "tipo_cupon_alt2": "Tipo\\s*de\\s*([A-Z]{1,4})\\s*Valido",
                "valido_hasta": "Valido hasta:\\s*([0-9A-Z\\-]+)",
                "valor_pagar": "Valor a pagar:\\s*\\$?([\\d,\\.]+)",
                "codigo_barras": "(\\(415\\)\\d+\\(8020\\)\\d+\\(3900\\)\\d+\\(96\\)\\d{8})"
            },
            "fields": [
                "Cliente", "Identificación", "Contrato", "Dirección",
                "No. Solicitud", "No. Rel. Pago", "Tipo de Cupón",
                "Valido hasta", "Valor a Pagar", "Codigo de Barras Limpio",
                "Archivo Origen"
            ],
            "settings": {
                "max_workers": 4,
                "output_format": "xlsx",
                "log_level": "INFO",
                "log_file": "extractor.log"
            }
        }
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=2, ensure_ascii=False)
        print("✅ Archivo de configuración creado")
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print("✅ Configuración cargada")
        return config
    except Exception as e:
        print(f"❌ Error cargando configuración: {e}")
        sys.exit(1)

def main():
    """Main entry point."""
    try:
        # Check dependencies first
        check_dependencies()

        # Load configuration
        config = load_config()
        
        # Import and run GUI
        from src.gui import ExtractorApp
        
        root = tk.Tk()
        app = ExtractorApp(root)
        root.mainloop()
        
    except ImportError as e:
        root = tk.Tk()
        root.withdraw()
        from tkinter import messagebox
        messagebox.showerror(
            "Error de Importación",
            f"No se pudo importar un módulo necesario:\n\n{str(e)}\n\n"
            f"Traceback:\n{traceback.format_exc()}"
        )
        sys.exit(1)
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        from tkinter import messagebox
        messagebox.showerror(
            "Error Inesperado",
            f"Ocurrió un error inesperado:\n\n{str(e)}\n\n"
            f"Traceback:\n{traceback.format_exc()}"
        )
        sys.exit(1)

if __name__ == "__main__":
    main()