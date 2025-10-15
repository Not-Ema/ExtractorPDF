import tkinter as tk
import sys
import json


# Load config to check settings
try:
    with open('config/config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
except FileNotFoundError:
    print("❌ config.json no encontrado. Asegúrate de que esté en el directorio actual.")
    sys.exit(1)

# Import and run GUI
from src.gui import ExtractorApp

if __name__ == "__main__":
    root = tk.Tk()
    app = ExtractorApp(root)
    root.mainloop()