import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from datetime import datetime
import json
from PIL import Image, ImageTk
from .pdf_processor import PDFProcessor
from .data_extractor import DataExtractor
from .data_writer import DataWriter
from .logger import logger

class ExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Datos PDF → Excel/CSV v0.1.0")
        self.root.geometry("800x650")
        self.root.configure(bg="#f8f9fa")

        # Load logo
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'logo.png')
        self.logo_image = ImageTk.PhotoImage(Image.open(logo_path).resize((32, 32), Image.LANCZOS))
        self.root.iconphoto(True, self.logo_image)

        # Load config
        with open('config/config.json', 'r', encoding='utf-8') as f:
            self.config = json.load(f)

        # Initialize modules
        self.processor = PDFProcessor()
        self.extractor = DataExtractor()
        self.writer = DataWriter()

        # Variables
        self.input_folder = ""
        self.output_file = ""
        self.is_processing = False
        self.scan_count = 0

        # Setup UI
        self.setup_styles()
        self.create_widgets()

        # Welcome message
        self.log_message("👋 ¡Bienvenido al Extractor de Datos PDF!", "info")
        self.log_message("💡 Sigue los pasos numerados para comenzar", "info")

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Helvetica', 12, 'bold'), background='#f8f9fa')
        style.configure('Success.TButton', background='#28a745', foreground='white')
        style.configure('Primary.TButton', background='#007bff', foreground='white')
        style.configure('Warning.TButton', background='#ffc107', foreground='black')
        style.configure('Error.TButton', background='#dc3545', foreground='white')

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="📊 Extractor de Datos PDF → Excel/CSV", font=('Helvetica', 16, 'bold'), style='Title.TLabel')
        title_label.pack(pady=(0, 20))

        self.create_input_section(main_frame)
        self.create_output_section(main_frame)
        self.create_control_section(main_frame)
        self.create_log_section(main_frame)
        self.create_footer(main_frame)

    def create_input_section(self, parent):
        input_frame = ttk.LabelFrame(parent, text="📁 Paso 1: Seleccionar Carpeta de PDFs", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        self.input_label = ttk.Label(input_frame, text="📍 Ninguna carpeta seleccionada", foreground="gray", font=('Helvetica', 9))
        self.input_label.pack(anchor="w", pady=(0, 5))

        input_btn = ttk.Button(input_frame, text="🗂️ Examinar Carpeta", command=self.select_input_folder, style='Primary.TButton')
        input_btn.pack(anchor="w")

    def create_output_section(self, parent):
        output_frame = ttk.LabelFrame(parent, text="💾 Paso 2: Guardar Archivo Excel/CSV", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        self.output_label = ttk.Label(output_frame, text="📍 Ningún archivo seleccionado", foreground="gray", font=('Helvetica', 9))
        self.output_label.pack(anchor="w", pady=(0, 5))

        output_btn = ttk.Button(output_frame, text="💾 Guardar Como...", command=self.select_output_file, style='Primary.TButton')
        output_btn.pack(anchor="w")

    def create_control_section(self, parent):
        control_frame = ttk.LabelFrame(parent, text="🚀 Paso 3: Procesar Archivos", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_button = ttk.Button(control_frame, text="▶️ Iniciar Extracción", command=self.start_processing, style='Success.TButton')
        self.start_button.pack(fill=tk.X, pady=(0, 10))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(control_frame, variable=self.progress_var, maximum=100, length=300)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X)

        clear_btn = ttk.Button(button_frame, text="🧹 Limpiar Log", command=self.clear_log, style='Warning.TButton')
        clear_btn.pack(side=tk.LEFT, padx=(0, 5))

        help_btn = ttk.Button(button_frame, text="❓ Ayuda", command=self.show_help)
        help_btn.pack(side=tk.LEFT, padx=(0, 5))

        about_btn = ttk.Button(button_frame, text="ℹ️ Acerca de", command=self.show_about)
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
            self.input_folder = folder
            folder_name = os.path.basename(folder)
            pdf_count = len([f for f in os.listdir(folder) if f.lower().endswith('.pdf')])
            self.input_label.config(text=f"📁 {folder_name} ({pdf_count} PDFs)", foreground="black")
            self.log_message(f"📁 Carpeta seleccionada: {folder_name}", "success")
            self.check_ready_to_process()

    def select_output_file(self):
        format_ext = '.xlsx' if self.config['settings']['output_format'] == 'xlsx' else '.csv'
        file_path = filedialog.asksaveasfilename(
            title="💾 Guardar archivo como...",
            defaultextension=format_ext,
            filetypes=[("Archivos Excel", "*.xlsx"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.output_file = file_path
            file_name = os.path.basename(file_path)
            self.output_label.config(text=f"💾 {file_name}", foreground="black")
            self.log_message(f"💾 Archivo de salida: {file_name}", "success")
            self.check_ready_to_process()

    def check_ready_to_process(self):
        if self.input_folder and self.output_file:
            self.log_message("✅ ¡Listo para procesar! Haz clic en 'Iniciar Extracción'", "success")

    def start_processing(self):
        if not self.input_folder or not self.output_file:
            messagebox.showwarning("⚠️ Campos incompletos", "Por favor selecciona:\n• Carpeta con PDFs\n• Archivo Excel/CSV de salida")
            return

        if self.is_processing:
            return

        self.is_processing = True
        self.start_button.config(text="⏳ Procesando...", state="disabled")
        self.progress_var.set(0)

        self.clear_log()
        self.log_message("🚀 Iniciando procesamiento...", "info")
        self.log_message(f"📂 Carpeta: {os.path.basename(self.input_folder)}", "info")
        self.log_message(f"📄 Archivo: {os.path.basename(self.output_file)}", "info")

        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()

    def process_files(self):
        try:
            pdf_files = [os.path.join(self.input_folder, f) for f in os.listdir(self.input_folder) if f.lower().endswith('.pdf')]
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
                        data = self.extractor.extract_data(text, filename)
                        data_list.append(data)
                        # Schedule GUI update for success
                        self.root.after(0, lambda: self._update_progress(filename, text, progress, processed_count, total_files))
                    except Exception as e:
                        errors_count += 1
                        logger.error(f"Error extracting data from {filename}: {e}")
                        # Schedule GUI update for error
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

            # Process PDFs concurrently
            self.processor.process_pdfs_concurrent(pdf_files, progress_callback)

            # Schedule final UI updates on main thread
            self.root.after(0, lambda: self._finalize_processing(data_list, errors_count, scan_count, total_files))

        except PermissionError as e:
            logger.error(f"Permission error: {e}")
            self.root.after(0, lambda: self.log_message(f"❌ Error de permisos: {str(e)}", "error"))
            self.root.after(0, lambda: messagebox.showerror("❌ Error de Permisos",
                                                           f"No se puede acceder a los archivos:\n\n{str(e)}\n\n"
                                                           "Verifica que los archivos no estén abiertos en otro programa."))
            self.root.after(0, lambda: self._reset_ui())
        except Exception as e:
            logger.error(f"Critical error: {e}", exc_info=True)
            self.root.after(0, lambda: self.log_message(f"❌ Error crítico: {str(e)}", "error"))
            self.root.after(0, lambda: messagebox.showerror("❌ Error", f"Error durante el procesamiento:\n\n{str(e)}"))
            self.root.after(0, lambda: self._reset_ui())

    def _finalize_processing(self, data_list, errors_count, scan_count, total_files):
        """Finalize processing and update UI on main thread."""
        try:
            if data_list:
                records = self.writer.write_data(data_list, self.output_file)
                self.log_message(f"🎉 ¡Proceso completado exitosamente!", "success")
                self.log_message(f"📊 Total de registros procesados: {records}", "success")

                if errors_count > 0 or scan_count > 0:
                    if errors_count > 0:
                        self.log_message(f"⚠️ Archivos con errores: {errors_count}", "warning")
                    if scan_count > 0:
                        self.log_message(f"📄 Archivos escaneados (omitidos): {scan_count}", "warning")

                success_rate = ((total_files - errors_count - scan_count) / total_files) * 100

                result = messagebox.askyesno("✅ Proceso Exitoso",
                                           f"¡Proceso completado!\n\n"
                                           f"📊 Registros procesados: {records}\n"
                                           f"✅ Archivos exitosos: {total_files - errors_count - scan_count}\n"
                                           f"❌ Archivos con errores: {errors_count}\n"
                                           f"📄 Archivos escaneados (omitidos): {scan_count}\n"
                                           f"📈 Tasa de éxito: {success_rate:.1f}%\n"
                                           f"📄 Archivo: {os.path.basename(self.output_file)}\n\n"
                                           f"¿Deseas abrir la carpeta del archivo?")
                if result:
                    self.open_output_folder()
            else:
                self.log_message("❌ No se procesaron archivos exitosamente", "error")
                messagebox.showwarning("⚠️ Sin Datos",
                                     "No se pudo extraer datos de ningún archivo.\n\n"
                                     "Verifica que los PDFs contengan el formato esperado.")
        except Exception as e:
            logger.error(f"Error in finalization: {e}")
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
        self.log_message("   📄 Archivo es un scan y no se procesa", "warning")
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
            folder_path = os.path.dirname(self.output_file)
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
2️⃣ Elige dónde guardar el archivo Excel/CSV de salida
3️⃣ Haz clic en 'Iniciar Extracción' y espera

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
• Los PDFs pueden ser digitales o escaneados
• Se pueden procesar múltiples archivos a la vez
• Los datos se agregan al Excel/CSV existente
        """
        messagebox.showinfo("❓ Ayuda", help_text)

    def show_about(self):
        about_text = """
📊 Extractor de Datos PDF → Excel/CSV v0.1.0

🎯 CARACTERÍSTICAS:
• Extracción automática de datos de PDFs
• Interfaz intuitiva y amigable
• Procesamiento concurrente para escalabilidad
• Barra de progreso en tiempo real
• Registro detallado de actividades

🛠️ TECNOLOGÍAS:
• Python 3.x
• pdfplumber (extracción de texto)
• pandas (manejo de datos)
• tkinter (interfaz gráfica)

        """
        messagebox.showinfo("ℹ️ Acerca de", about_text)