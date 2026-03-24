import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import csv
import os
import threading
import requests

# Intentamos importar PyMuPDF o pypdf (librerías ligeras para PDFs)
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

class CURPExtractorLiteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de CURPs (Versión Steam Deck/Ligera)")
        self.root.minsize(900, 600)
        
        self.extracted_curps = []
        
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Extractor de CURPs (Sin Tesseract/Selenium)", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Archivos
        upload_frame = ttk.LabelFrame(main_frame, text="Subir archivos", padding="10")
        upload_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        upload_frame.columnconfigure(1, weight=1)
        
        upload_btn = ttk.Button(upload_frame, text="Subir PDFs y Fotos de CURP", command=self.upload_files)
        upload_btn.grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        
        self.progress = ttk.Progressbar(upload_frame, mode='indeterminate')
        self.progress.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.status_label = ttk.Label(upload_frame, text="Preparado para procesar archivos")
        self.status_label.grid(row=0, column=2, sticky=tk.E)
        
        # Entrada manual
        input_frame = ttk.LabelFrame(main_frame, text="Entrada manual/masiva", padding="10")
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="Ingresa CURP:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.curp_entry = ttk.Entry(input_frame, font=("Courier", 11), width=20)
        self.curp_entry.grid(row=0, column=1, padx=(0, 10), sticky=(tk.W, tk.E))
        self.curp_entry.bind('<Return>', self.add_manual_curp)
        self.curp_entry.bind('<KeyRelease>', self.validate_curp_input)
        
        self.add_curp_btn = ttk.Button(input_frame, text="Agregar", command=self.add_manual_curp)
        self.add_curp_btn.grid(row=0, column=2, padx=(0, 10))
        
        self.validation_label = ttk.Label(input_frame, text="", font=("Arial", 9))
        self.validation_label.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        
        # Tabla de resultados
        results_frame = ttk.LabelFrame(main_frame, text="CURPs Extraídas", padding="10")
        results_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)
        
        columns = ('Origen', 'CURP', 'Estatus')
        self.tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=12)
        
        self.tree.heading('Origen', text='Origen')
        self.tree.heading('CURP', text='CURP')
        self.tree.heading('Estatus', text='Estatus')
        
        self.tree.column('Origen', width=200, minwidth=150)
        self.tree.column('CURP', width=200, minwidth=180)
        self.tree.column('Estatus', width=100, minwidth=80)
        
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        self.tree.bind('<Button-3>', self.show_context_menu)
        
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Copiar CURP", command=self.copy_single_curp)
        self.context_menu.add_command(label="Quitar CURP", command=self.remove_single_curp)

        # Botones de acción
        actions_frame = ttk.Frame(main_frame)
        actions_frame.grid(row=4, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        copy_btn = ttk.Button(actions_frame, text="Copiar todas las CURPs", command=self.copy_to_clipboard)
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        download_btn = ttk.Button(actions_frame, text="Exportar a CSV", command=self.download_csv)
        download_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_btn = ttk.Button(actions_frame, text="Limpiar tabla", command=self.clear_results)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        instructions = ("Nota de esta versión:\n"
                       "- Ya no usa Selenium ni Tesseract local para correr bien en el Steam Deck.\n"
                       "- Los PDFs se extraen con texto nativo.\n"
                       "- Las Imágenes usan una API de OCR gratuita (requiere internet).")
        
        instructions_label = ttk.Label(main_frame, text=instructions, font=("Arial", 9), 
                                     foreground="gray", justify=tk.LEFT)
        instructions_label.grid(row=5, column=0, columnspan=3, pady=(20, 0), sticky=tk.W)
        
    def validate_curp_input(self, event=None):
        curp = self.curp_entry.get().upper()
        self.curp_entry.delete(0, tk.END)
        self.curp_entry.insert(0, curp)
        if not curp:
            self.validation_label.config(text="", foreground="black")
            return
        if len(curp) < 18:
            self.validation_label.config(text=f"Longitud: {len(curp)}/18 caracteres", foreground="orange")
        elif len(curp) == 18:
            if self.is_valid_curp_format(curp):
                self.validation_label.config(text="✓ Formato Valido", foreground="green")
            else:
                self.validation_label.config(text="✗ Formato Inválido", foreground="red")
        else:
            self.validation_label.config(text="✗ Demasiados caracteres", foreground="red")

    def is_valid_curp_format(self, curp):
        pattern = r'^[A-Z]{4}[0-9]{6}[HMX][A-Z]{5}[A-Z0-9][0-9]$'
        return bool(re.match(pattern, curp))

    def add_manual_curp(self, event=None):
        curp = self.curp_entry.get().strip().upper()
        if not curp: return
        
        if len(curp) != 18:
            messagebox.showwarning("CURP Incorrecta", "Debe tener 18 caracteres.")
            return

        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            if len(values) >= 2 and values[1] == curp:
                messagebox.showwarning("CURP Duplicada", f"La CURP '{curp}' ya está agregada.")
                return

        self.add_result("Ingreso manual", curp, "Agregada")
        self.extracted_curps.append(curp)
        self.curp_entry.delete(0, tk.END)
        self.validation_label.config(text="")
        self.status_label.config(text=f"CURP agregada: {curp}")

    def remove_single_curp(self):
        if hasattr(self, 'context_menu_item'):
            item = self.context_menu_item
            values = self.tree.item(item)['values']
            if len(values) >= 2:
                curp = values[1]
                self.tree.delete(item)
                try:
                    self.extracted_curps.remove(curp)
                except ValueError:
                    pass
                self.status_label.config(text=f"CURP eliminada: {curp}")

    def upload_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Seleccionar Archivos",
            filetypes=[
                ("Todos los soportados", "*.pdf *.jpg *.jpeg *.png *.bmp *.webp"),
                ("Archivos PDF", "*.pdf"), 
                ("Imágenes", "*.jpg *.jpeg *.png *.bmp *.webp")
            ]
        )
        if file_paths:
            threading.Thread(target=self.process_files, args=(file_paths,), daemon=True).start()

    def process_files(self, file_paths):
        self.root.after(0, lambda: self.progress.start())
        self.root.after(0, lambda: self.status_label.config(text="Procesando archivos..."))
        
        for file_path in file_paths:
            try:
                filename = os.path.basename(file_path)
                self.root.after(0, lambda f=filename: self.status_label.config(text=f"Procesando: {f}"))
                
                ext = os.path.splitext(file_path)[1].lower()
                if ext == '.pdf':
                    curp = self.extract_curp_from_pdf(file_path)
                else:
                    curp = self.extract_curp_from_image(file_path)
                    
                if curp:
                    self.root.after(0, lambda f=filename, c=curp: self.add_result(f, c, "Éxito"))
                    self.extracted_curps.append(curp)
                else:
                    self.root.after(0, lambda f=filename: self.add_result(f, "Falla", "CURP no encontrada"))
                    
            except Exception as e:
                filename = os.path.basename(file_path)
                err = str(e)[:40]
                self.root.after(0, lambda f=filename, err=err: self.add_result(f, "Error", f"Error: {err}"))
        
        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.status_label.config(text="Procesamiento completo."))

    def extract_curp_from_image(self, image_path):
        try:
            # Usando la API gratuita de OCR.space (apikey 'helloworld' o uno de prueba)
            payload = {
                'isOverlayRequired': False,
                'apikey': 'helloworld',
                'language': 'spa',
                'OCREngine': 2 # Engine 2 es mejor para caracteres alfanuméricos mezclados y comprobantes
            }
            with open(image_path, 'rb') as f:
                r = requests.post('https://api.ocr.space/parse/image',
                                  files={os.path.basename(image_path): f},
                                  data=payload,
                                  timeout=30)
            
            result = r.json()
            if result.get("IsErroredOnProcessing"):
                print("Error de API OCR:", result.get("ErrorMessage"))
                return None
                
            parsed_results = result.get("ParsedResults", [])
            if not parsed_results:
                return None
                
            text = parsed_results[0].get("ParsedText", "")
            return self.find_curp_in_text(text)
            
        except Exception as e:
            print("Excepción llamando a la API de OCR:", e)
            return None

    def extract_curp_from_pdf(self, pdf_path):
        text = ""
        # Usamos fitz si está disponible, si no, intentamos pypdf
        if HAS_FITZ:
            try:
                doc = fitz.open(pdf_path)
                for page in doc:
                    text += page.get_text()
                doc.close()
            except:
                pass
        elif HAS_PYPDF:
            try:
                reader = PdfReader(pdf_path)
                for page in reader.pages:
                    text += page.extract_text() or ""
            except:
                pass
        else:
            print("No se encontró `PyMuPDF` ni `pypdf`. Se usará lectura pura de texto (menos confiable).")
            # Método muy rústico pero funciona para algunos PDFs textuales
            try:
                with open(pdf_path, 'rb') as f:
                    content = f.read().decode('utf-8', errors='ignore')
                    text = content
            except:
                pass

        return self.find_curp_in_text(text)

    def find_curp_in_text(self, text):
        if not text:
            return None
        
        # Patrones para buscar la CURP
        pattern = r'Clave:\s*([A-Za-z0-9]{18})'
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).upper()
            
        pattern2 = r'Clave\s*([A-Za-z0-9]{18})'
        match2 = re.search(pattern2, text, re.IGNORECASE)
        if match2:
            return match2.group(1).upper()
            
        # Buscar "aleatoriamente" una CURP si no está etiquetada como Clave:
        pattern3 = r'\b([A-Z]{4}[0-9]{6}[HMX][A-Z]{5}[A-Z0-9][0-9])\b'
        match3 = re.search(pattern3, text, re.IGNORECASE)
        if match3:
            return match3.group(1).upper()

        return None

    def add_result(self, source, curp, status):
        item_id = self.tree.insert('', 'end', values=(source, curp, status))
        return item_id

    def copy_to_clipboard(self):
        valid_curps = [c for c in self.extracted_curps if c and len(c) == 18]
        if not valid_curps: return
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(valid_curps))
        self.root.update()
        messagebox.showinfo("Copiadas", f"¡Se copiaron {len(valid_curps)} CURPs al portapapeles!")

    def download_csv(self):
        if not self.tree.get_children(): return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if file_path:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['Origen', 'CURP', 'Estatus'])
                for item in self.tree.get_children():
                    writer.writerow(self.tree.item(item)['values'])
            messagebox.showinfo("Éxito", "Archivo CSV guardado.")

    def clear_results(self):
        self.tree.delete(*self.tree.get_children())
        self.extracted_curps.clear()

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu_item = item
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def copy_single_curp(self):
        if hasattr(self, 'context_menu_item'):
            values = self.tree.item(self.context_menu_item)['values']
            if len(values) >= 2:
                curp = values[1]
                self.root.clipboard_clear()
                self.root.clipboard_append(str(curp))
                self.root.update()

def main():
    if not HAS_FITZ and not HAS_PYPDF:
        print("Advertencia: No cuentas con PyMuPDF (fitz) ni pypdf instalados.")
        print("Para mejor extracción de PDFs, instala una de las dos:")
        print("  pip install PyMuPDF")
        print("o pip install pypdf")
        
    root = tk.Tk()
    app = CURPExtractorLiteApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
