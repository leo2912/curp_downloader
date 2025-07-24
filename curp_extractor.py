import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pytesseract
from PIL import Image, ImageTk
import re
import csv
import os
from pathlib import Path
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import time
import fitz  # PyMuPDF for PDF processing

class CURPExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor y Descargador de CURPs")
        self.root.geometry("900x700")
        
        # Store extracted CURPs
        self.extracted_curps = []
        
        # Web driver instance
        self.driver = None
        self.download_folder = None
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Extractor y Descargador de CURPs", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Upload button
        #upload_btn = ttk.Button(main_frame, text="Upload Files (Images & PDFs)", command=self.upload_files)
        #upload_btn.grid(row=1, column=0, padx=(0, 10), pady=(0, 10), sticky=tk.W)
     
     # File upload section
        upload_frame = ttk.LabelFrame(main_frame, text="Subir archivos", padding="10")
        upload_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        upload_frame.columnconfigure(1, weight=1)
        upload_btn = ttk.Button(upload_frame, text="Subir archivos (Imágenes & PDFs)", command=self.upload_files)
        upload_btn.grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        
        # Progress bar
        # self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        # self.progress.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(0, 10))
        self.progress = ttk.Progressbar(upload_frame, mode='indeterminate')
        self.progress.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(upload_frame, text="Preparado para procesar archivos")
        #self.status_label.grid(row=1, column=2, pady=(0, 10), sticky=tk.E)
        self.status_label.grid(row=0, column=2, sticky=tk.E)
        
        # Manual input section

        input_frame = ttk.LabelFrame(main_frame, text="Entrada manual de CURPs", padding="10")
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
       

        # CURP input field
        ttk.Label(input_frame, text="Ingresa CURP:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.curp_entry = ttk.Entry(input_frame, font=("Courier", 11), width=20)
        self.curp_entry.grid(row=0, column=1, padx=(0, 10), sticky=(tk.W, tk.E))
        self.curp_entry.bind('<Return>', self.add_manual_curp)
        self.curp_entry.bind('<KeyRelease>', self.validate_curp_input)

        

        # Add CURP button
        self.add_curp_btn = ttk.Button(input_frame, text="Agrega CURP", command=self.add_manual_curp)
        self.add_curp_btn.grid(row=0, column=2, padx=(0, 10))

        # Validation label
        self.validation_label = ttk.Label(input_frame, text="", font=("Arial", 9))
        self.validation_label.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))

        # Bulk input section
        bulk_frame = ttk.Frame(input_frame)
        bulk_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        bulk_frame.columnconfigure(0, weight=1)
        ttk.Label(bulk_frame, text="Múltiples CURPs (Una CURP por línea):").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        # Text area for bulk input
        text_frame = ttk.Frame(bulk_frame)
        text_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        text_frame.columnconfigure(0, weight=1)
        self.bulk_text = tk.Text(text_frame, height=4, width=50, font=("Courier", 10))
        self.bulk_text.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Scrollbar for text area
        bulk_scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.bulk_text.yview)
        self.bulk_text.configure(yscrollcommand=bulk_scrollbar.set)
        bulk_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Bulk add button
        bulk_add_btn = ttk.Button(bulk_frame, text="Agregar todas las CURPs", command=self.add_bulk_curps)
        bulk_add_btn.grid(row=2, column=0, pady=(5, 0), sticky=tk.W)

        # Clear bulk button
        clear_bulk_btn = ttk.Button(bulk_frame, text="Borrar", command=self.clear_bulk_text)
        clear_bulk_btn.grid(row=2, column=1, pady=(5, 0), padx=(10, 0), sticky=tk.W)

        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="CURPs Extraídas", padding="10")
        results_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)
        
        # Selection controls frame
        selection_frame = ttk.Frame(results_frame)
        selection_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Select all/none buttons
        select_all_btn = ttk.Button(selection_frame, text="Seleccionar todo", command=self.select_all_items)
        select_all_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        select_none_btn = ttk.Button(selection_frame, text="Deseleccionar Todo", command=self.select_none_items)
        select_none_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download selected button
        download_selected_btn = ttk.Button(selection_frame, text="Descargar PDFs seleccionados", 
                                         command=self.download_selected_pdfs)
        download_selected_btn.pack(side=tk.LEFT)
        
        # Treeview for results
        #columns = ('Select', 'File', 'CURP', 'Status')
        columns = ('Seleccionar', 'Origen', 'CURP', 'Estatus')
        self.tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=12)
        
        # Define headings
        self.tree.heading('Seleccionar', text='☐')
        #self.tree.heading('File', text='File Name')
        self.tree.heading('Origen', text='Origen')
        self.tree.heading('CURP', text='CURP')
        self.tree.heading('Estatus', text='Estatus')
        
        # Configure column widths
        self.tree.column('Seleccionar', width=50, minwidth=50)
        self.tree.column('Origen', width=200, minwidth=150)
        #self.tree.column('File', width=200, minwidth=150)
        self.tree.column('CURP', width=200, minwidth=180)
        self.tree.column('Estatus', width=100, minwidth=80)
        
        # Bind events for checkbox functionality and right-click menu
        self.tree.bind('<Button-1>', self.on_treeview_click)
        self.tree.bind('<Button-3>', self.show_context_menu)  # Right-click
        
        # Store selection state for each item
        self.selected_items = set()
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Create context menu
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Descargar el PDF de esta CURP", command=self.download_single_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Copiar CURP", command=self.copy_single_curp)
        self.context_menu.add_command(label="Alternar Selección", command=self.toggle_single_selection)
        self.context_menu.add_command(label="Quitar CURP", command=self.remove_single_curp)

        # Action buttons frame
        actions_frame = ttk.Frame(main_frame)
        actions_frame.grid(row=5, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # Copy button
        copy_btn = ttk.Button(actions_frame, text="Copiar CURPs", command=self.copy_to_clipboard)
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download CSV button
        download_btn = ttk.Button(actions_frame, text="Descargar como CSV", command=self.download_csv)
        download_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear button
        clear_btn = ttk.Button(actions_frame, text="Limpiar", command=self.clear_results)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download PDFs button
        download_pdfs_btn = ttk.Button(actions_frame, text="Descargar CURP PDFs", command=self.download_curp_pdfs)
        download_pdfs_btn.pack(side=tk.LEFT)
        
        # Instructions label
        # instructions = ("Instructions:\n"
        #                "1. Click 'Upload Files' to select images (JPG, PNG, etc.) or PDF files\n"
        #                "2. The app will extract 18-character CURP codes after 'Clave:'\n"
        #                "3. View results in the table below\n"
        #                "4. Use checkboxes to select CURPs, or right-click for individual actions\n"
        #                "5. Copy codes to clipboard, download as CSV, or download official PDFs")

        # Instructions label

        instructions = ("Instructions:\n"
                       "1. Upload files: Click 'Upload Files' to select images (JPG, PNG, etc.) or PDF files\n"
                       "2. Manual input: Type CURPs directly in the input field or use bulk input\n"
                       "3. The app will validate and display all CURPs in the table below\n"
                       "4. Use checkboxes to select CURPs, or right-click for individual actions\n"
                       "5. Copy codes to clipboard, download as CSV, or download official PDFs")
        
        instructions_label = ttk.Label(main_frame, text=instructions, font=("Arial", 9), 
                                     foreground="gray", justify=tk.LEFT)
        instructions_label.grid(row=6, column=0, columnspan=3, pady=(20, 0), sticky=tk.W)
    def validate_curp_input(self, event=None):
        """Validate CURP input in real-time"""
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
                self.validation_label.config(text="✓ Formato Valido CURP", foreground="green")

            else:
                self.validation_label.config(text="✗ Formato Invalido CURP", foreground="red")

        else:
            self.validation_label.config(text="✗ CURP Incorrecta (no tiene 18 caracteres)", foreground="red")

    def is_valid_curp_format(self, curp):
        """Validate CURP format using regex"""
        # Basic CURP format: 4 letters + 6 digits + 1 letter + 5 alphanumeric + 2 digits
        pattern = r'^[A-Z]{4}[0-9]{6}[HM][A-Z]{5}[0-9]{2}$'
        return bool(re.match(pattern, curp))
   

    def add_manual_curp(self, event=None):
        """Add manually entered CURP to results"""
        curp = self.curp_entry.get().strip().upper()
  
        if not curp:
            messagebox.showwarning("Campo vacío", "Escribe una CURP.")
            return

        if len(curp) != 18:
            messagebox.showwarning("CURP Incorrecta", "LA CURP debe tener 18 caracteres.")
            return

        if not self.is_valid_curp_format(curp):
            result = messagebox.askyesno("Formato Inválido", 
                                       f"La CURP '{curp}' no cumple por el formato estándar.\n\n"
                                       "¿Quieres agregarla de todas formas?")
            if not result:
                return
 
        # Check if CURP already exists
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']

            if len(values) >= 3 and values[2] == curp:
                messagebox.showwarning("CURP Duplicada", f"CURP '{curp}' existente.")
                return

        # Add to results
        self.add_result("Ingreso manual", curp, "agregada")
        self.extracted_curps.append(curp)
  
        # Clear input field
        self.curp_entry.delete(0, tk.END)
        self.validation_label.config(text="")
  
        # Show success message
        self.status_label.config(text=f"CURP agregada: {curp}")
   
    def add_bulk_curps(self):
        """Add multiple CURPs from bulk text input"""
        text = self.bulk_text.get("1.0", tk.END).strip()

        if not text:
            messagebox.showwarning("Campo vacío", "Por favor escribe las CURPs en el cuadro de texto.")
            return

        # Split by lines and clean up
        lines = [line.strip().upper() for line in text.split('\n') if line.strip()]
        added_count = 0
        skipped_count = 0
        invalid_count = 0

        for line in lines:
            # Skip if not 18 characters
            if len(line) != 18:
                invalid_count += 1
                continue
  
           # Check if already exists
            exists = False

            for item in self.tree.get_children():
                values = self.tree.item(item)['values']

                if len(values) >= 3 and values[2] == line:
                    exists = True
                    break

            if exists:
                skipped_count += 1
                continue
           
            # Add to results
            status = "Added"

            if not self.is_valid_curp_format(line):
                status = "Agregada (Formato Invalido)"

            self.add_result("Bulk Input", line, status)
            self.extracted_curps.append(line)
            added_count += 1
  
        # Show summary
        summary = f"Importación masiva completada:\n"
        summary += f"CURPs agregadas: {added_count}\n"
        if skipped_count > 0:
            summary += f"CURPs duplicadas no agregadas: {skipped_count}\n"

        if invalid_count > 0:
            summary += f"CURPs invalidas (no tienen 18 caracteres): {invalid_count}"
    
        messagebox.showinfo("Resultados importación masiva", summary)
        self.status_label.config(text=f"Importación masiva: {added_count} fueron agregadas CURPs")
 
    def clear_bulk_text(self):
        """Clear the bulk input text area"""
        self.bulk_text.delete("1.0", tk.END)

    def remove_single_curp(self):
        """Remove the right-clicked CURP from results"""
        if hasattr(self, 'context_menu_item'):
            item = self.context_menu_item
            values = self.tree.item(item)['values']

            if len(values) >= 3:
                curp = values[2]
                result = messagebox.askyesno("Confirmar eliminación", f"¿Quieres quitar la '{curp}' de la lista?")
  
                if result:
                    # Remove from tree
                    self.tree.delete(item)
                    
                    # Remove from selected items if it was selected
                    if item in self.selected_items:
                        self.selected_items.remove(item)
                    
                    # Remove from extracted_curps list
                    try:
                        self.extracted_curps.remove(curp)
                    except ValueError:
                        pass  # CURP not in list
 
                    self.status_label.config(text=f"CURP eliminada: {curp}")

        
    def upload_files(self):
        """Handle file upload (both images and PDFs)"""
        file_paths = filedialog.askopenfilenames(
            title="Select Files",
            filetypes=[
                ("Archivos compatibles", "*.jpg *.jpeg *.png *.bmp *.tiff *.gif *.pdf"),
                ("Imágenes", "*.jpg *.jpeg *.png *.bmp *.tiff *.gif"),
                ("PDFs", "*.pdf"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        if file_paths:
            # Process files in a separate thread to avoid blocking the UI
            threading.Thread(target=self.process_files, args=(file_paths,), daemon=True).start()
    
    def upload_images(self):
        """Handle image upload (kept for backward compatibility)"""
        self.upload_files()
    
    def process_files(self, file_paths):
        """Process files (both images and PDFs) and extract CURPs"""
        self.root.after(0, lambda: self.progress.start())
        self.root.after(0, lambda: self.status_label.config(text="Procesando archivos..."))
        
        for i, file_path in enumerate(file_paths):
            try:
                # Update status
                filename = os.path.basename(file_path)
                self.root.after(0, lambda f=filename: self.status_label.config(text=f"Procesando: {f}"))
                
                # Check file type
                file_extension = os.path.splitext(file_path)[1].lower()
                
                if file_extension == '.pdf':
                    # Process PDF file
                    curp = self.extract_curp_from_pdf(file_path)
                else:
                    # Process image file
                    curp = self.extract_curp_from_image(file_path)
                
                if curp:
                    # Add to results
                    self.root.after(0, lambda f=filename, c=curp: self.add_result(f, c, "Éxito"))
                    self.extracted_curps.append(curp)
                else:
                    self.root.after(0, lambda f=filename: self.add_result(f, "Falla", "CURP no encontrada"))
                    self.extracted_curps.append(None)
                    
            except Exception as e:
                filename = os.path.basename(file_path)
                error_msg = str(e)[:50] + "..." if len(str(e)) > 50 else str(e)
                self.root.after(0, lambda f=filename, err=error_msg: self.add_result(f, f"Error: {err}", "Error"))
        
        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.status_label.config(text=f"Extracción completa. Se encontrarón {len([c for c in self.extracted_curps if c])} CURPs"))
    
    def process_images(self, file_paths):
        """Process images and extract CURPs (kept for backward compatibility)"""
        self.process_files(file_paths)
    
    def extract_curp_from_image(self, image_path):
        """Extract CURP code from image using OCR"""
        try:
            # Open and process image
            image = Image.open(image_path)
            
            # Convert to RGB if necessary
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Perform OCR
            text = pytesseract.image_to_string(image, lang='spa+eng')
            
            # Search for CURP in the extracted text
            return self.find_curp_in_text(text)
        
        except Exception as e:
            print(f"Error extracting CURP {image_path}: {e}")
            return None
    
    def extract_curp_from_pdf(self, pdf_path):
        """Extract CURP code from PDF using OCR"""
        try:
            # Open PDF document
            pdf_document = fitz.open(pdf_path)
            
            # Process each page of the PDF
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                
                # First, try to extract text directly from PDF
                text = page.get_text()
                curp = self.find_curp_in_text(text)
                if curp:
                    pdf_document.close()
                    return curp
                
                # If direct text extraction doesn't work, convert page to image and use OCR
                # Get page as image (PNG format)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for better OCR
                img_data = pix.tobytes("png")
                
                # Convert to PIL Image
                from io import BytesIO
                image = Image.open(BytesIO(img_data))
                
                # Perform OCR on the image
                ocr_text = pytesseract.image_to_string(image, lang='spa+eng')
                curp = self.find_curp_in_text(ocr_text)
                
                if curp:
                    pdf_document.close()
                    return curp
            
            pdf_document.close()
            return None
            
        except Exception as e:
            print(f"Error procesando PDF {pdf_path}: {e}")
            return None
    
    def find_curp_in_text(self, text):
        """Find CURP pattern in text"""
        # Search for CURP pattern after "Clave:"
        # Look for "Clave:" followed by optional whitespace and then exactly 18 alphanumeric characters
        
        pattern = r'Clave:\s*([A-Z0-9]{18})'
        match = re.search(pattern, text, re.IGNORECASE)
        
        if match:
            return match.group(1).upper()
        
        # Alternative pattern - sometimes OCR might not capture the colon correctly
        pattern2 = r'Clave\s*([A-Z0-9]{18})'
        match2 = re.search(pattern2, text, re.IGNORECASE)
        
        if match2:
            return match2.group(1).upper()
            
        return None
            
        # except Exception as e:
        #     print(f"Error processing image {image_path}: {e}")
        #     return None
    
    def add_result(self, source, curp, status):
        """Add result to the treeview"""
        # Add checkbox symbol (unchecked by default)
        checkbox = "☐"
        item_id = self.tree.insert('', 'end', values=(checkbox, source, curp, status))
        return item_id
    
    def copy_to_clipboard(self):
        """Copy all valid CURPs to clipboard"""
        valid_curps = [curp for curp in self.extracted_curps if curp and len(curp) == 18]
        
        if not valid_curps:
            messagebox.showwarning("No CURPs", "No se encontraron CURPs validas.")
            return
        
        # Join CURPs with newlines
        curps_text = '\n'.join(valid_curps)
        
        # Copy to clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(curps_text)
        self.root.update()  # Ensure clipboard is updated
        
        messagebox.showinfo("Copiadas", f"Se copiaron {len(valid_curps)} CURPs en el portapapeles!")
    
    def download_csv(self):
        """Download results as CSV file"""
        if not self.tree.get_children():
            messagebox.showwarning("No Datos", "No hay datos para descargar.")
            return
        
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Archivo CSV", "*.csv"), ("Todos los archivos", "*.*")],
            title="Guardar archivo CSV"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # Write header
                    writer.writerow(['Origen', 'CURP', 'Estatus'])
                    
                    # Write data
                    for item in self.tree.get_children():
                        values = self.tree.item(item)['values']
                        # Skip the checkbox column when writing to CSV
                        if len(values) >= 4:
                            writer.writerow([values[1], values[2], values[3]])
                
                messagebox.showinfo("Éxito", f"Archivo CSV file guardado.\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar archivo CSV:\n{e}")
    
    def clear_results(self):
        """Clear all results"""
        self.tree.delete(*self.tree.get_children())
        self.extracted_curps.clear()
        self.selected_items.clear()
        self.status_label.config(text="Preparado para procesar archivos")
    
    def on_treeview_click(self, event):
        """Handle treeview click events for checkbox functionality"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#1":  # Select column
                item = self.tree.identify_row(event.y)
                if item:
                    self.toggle_item_selection(item)
    
    def toggle_item_selection(self, item):
        """Toggle selection state of an item"""
        values = list(self.tree.item(item)['values'])
        if len(values) >= 4:
            curp = values[2]
            # Only allow selection of valid CURPs
            if curp and len(str(curp)) == 18 and curp != "Not found" and not str(curp).startswith("Error:"):
                if item in self.selected_items:
                    self.selected_items.remove(item)
                    values[0] = "☐"
                else:
                    self.selected_items.add(item)
                    values[0] = "☑"
                self.tree.item(item, values=values)
    
    def select_all_items(self):
        """Select all valid CURPs"""
        for item in self.tree.get_children():
            values = list(self.tree.item(item)['values'])
            if len(values) >= 4:
                curp = values[2]
                if curp and len(str(curp)) == 18 and curp != "Not found" and not str(curp).startswith("Error:"):
                    if item not in self.selected_items:
                        self.selected_items.add(item)
                        values[0] = "☑"
                        self.tree.item(item, values=values)
    
    def select_none_items(self):
        """Deselect all items"""
        for item in self.tree.get_children():
            values = list(self.tree.item(item)['values'])
            if len(values) >= 4:
                if item in self.selected_items:
                    self.selected_items.remove(item)
                    values[0] = "☐"
                    self.tree.item(item, values=values)
    
    def show_context_menu(self, event):
        """Show context menu on right-click"""
        item = self.tree.identify_row(event.y)
        if item:
            # Select the right-clicked item
            self.tree.selection_set(item)
            self.context_menu_item = item
            # Show context menu
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()
    
    def download_single_selected(self):
        """Download PDF for the right-clicked CURP"""
        if hasattr(self, 'context_menu_item'):
            values = self.tree.item(self.context_menu_item)['values']
            if len(values) >= 4:
                curp = values[2]
                if curp and len(str(curp)) == 18 and curp != "Not found" and not str(curp).startswith("Error:"):
                    # Start download in separate thread
                    threading.Thread(target=self.download_pdfs_worker, args=([curp],), daemon=True).start()
                else:
                    messagebox.showwarning("CURP invalida", "No se pueden descargar PDFs de CURPs invalidas.")
    
    def copy_single_curp(self):
        """Copy the right-clicked CURP to clipboard"""
        if hasattr(self, 'context_menu_item'):
            values = self.tree.item(self.context_menu_item)['values']
            if len(values) >= 4:
                curp = values[2]
                if curp and curp != "Not found":
                    self.root.clipboard_clear()
                    self.root.clipboard_append(str(curp))
                    self.root.update()
                    messagebox.showinfo("Copiada", f"CURP '{curp}' copiada al portapapeles.")
    
    def toggle_single_selection(self):
        """Toggle selection of the right-clicked item"""
        if hasattr(self, 'context_menu_item'):
            self.toggle_item_selection(self.context_menu_item)
    
    def download_selected_pdfs(self):
        """Download PDFs for selected CURPs only"""
        if not self.selected_items:
            messagebox.showwarning("Sin selección", "Por favor, marca con palomita las CURPS que quieras descargar.")
            return
        
        # Get CURPs from selected items
        selected_curps = []
        for item in self.selected_items:
            values = self.tree.item(item)['values']
            if len(values) >= 4:
                curp = values[2]
                if curp and len(str(curp)) == 18 and curp != "Not found":
                    selected_curps.append(curp)
        
        if not selected_curps:
            messagebox.showwarning("No hay CURPs validas", "No se seleccionardon CURPs validas para descargar.")
            return
        
        # Confirm action
        result = messagebox.askyesno(
            "Confirmación de Descarga", 
            f"Se descargaran {len(selected_curps)} CURPs de la página oficial.\n\n"
            "Este proceso puede tardar varios minutos. ¿Quieres continuar?"
        )
        
        if result:
            # Start download process in separate thread
            threading.Thread(target=self.download_pdfs_worker, args=(selected_curps,), daemon=True).start()
    
    def setup_webdriver(self):
        """Setup Chrome WebDriver with download preferences"""
        try:
            # Ask user for download folder
            self.download_folder = filedialog.askdirectory(title="Selecciona carpeta para guardar los CURPs en PDF.")
            if not self.download_folder:
                return None
            
            # Chrome options
            chrome_options = Options()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            
            # Set download preferences
            prefs = {
                "download.default_directory": self.download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # Initialize WebDriver
            #self.driver = webdriver.Chrome(options=chrome_options)
            self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
            return self.driver
            
        except Exception as e:
            messagebox.showerror("Error en el WebDriver", f"No fue posible iniciar el WebDriver:\n{e}\n\nRevisa que el WebDriver está instalado.")
            return None
    
    def download_curp_pdfs(self):
        """Download official CURP PDFs for all valid CURPs"""
        valid_curps = [curp for curp in self.extracted_curps if curp and len(curp) == 18]
        
        if not valid_curps:
            messagebox.showwarning("No CURPs", "No se encontraron CURPs validas para descargar.")
            return
        
        # Confirm action
        result = messagebox.askyesno(
            "Confirmación de Descarga", 
            f"Se descargaran {len(valid_curps)} CURPs de la página oficial.\n\n"
            "Este proceso puede tardar varios minutos. ¿Quieres continuar?"
        )
        
        if not result:
            return
        
        # Start download process in separate thread
        threading.Thread(target=self.download_pdfs_worker, args=(valid_curps,), daemon=True).start()
    
    def download_pdfs_worker(self, curps):
        """Worker function to download PDFs"""
        self.root.after(0, lambda: self.progress.start())
        self.root.after(0, lambda: self.status_label.config(text="Preparando navegador..."))
        
        # Setup WebDriver
        if not self.setup_webdriver():
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.status_label.config(text="Descarga cancelada"))
            return
        
        successful_downloads = 0
        failed_downloads = 0
        
        try:
            for i, curp in enumerate(curps):
                try:
                    self.root.after(0, lambda c=curp, idx=i+1, total=len(curps): 
                                  self.status_label.config(text=f"Descargando {idx}/{total}: {c}"))
                    
                    if self.download_single_curp_pdf(curp):
                        successful_downloads += 1
                    else:
                        failed_downloads += 1
                    
                    # Small delay between requests to be respectful
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"Error descargando CURP {curp}: {e}")
                    failed_downloads += 1
                    continue
        
        finally:
            # Cleanup
            if self.driver:
                self.driver.quit()
                self.driver = None
            
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.status_label.config(
                text=f"Descarga completada: {successful_downloads} exitosas, {failed_downloads} fallida"))
            
            # Show completion message
            self.root.after(0, lambda: messagebox.showinfo(
                "Descarga completa", 
                f"Fueron descargadas {successful_downloads} CURPs en PDF de manera exitosa.\n"
                f"{failed_downloads} downloads failed.\n\n"
                f"Archivos guardados en: {self.download_folder}"
            ))
    
    def download_single_curp_pdf(self, curp):
        """Download PDF for a single CURP"""
        try:
            # Navigate to CURP website
            self.driver.get("https://www.gob.mx/curp/")
            
            # Wait for page to load and find the input field
            wait = WebDriverWait(self.driver, 10)
            curp_input = wait.until(EC.presence_of_element_located((By.ID, "curpinput")))
            
            # Clear and enter CURP
            curp_input.clear()
            curp_input.send_keys(curp)
            
            # Find and click search button
            search_button = self.driver.find_element(By.ID, "searchButton")
            search_button.click()
            
            # Wait for page to refresh and results to load
            time.sleep(3)
            
            # Try to find and click download button
            try:
                download_button = wait.until(EC.element_to_be_clickable((By.ID, "download")))
                download_button.click()
                
                # Wait for download to start
                time.sleep(2)
                
                return True
                
            except TimeoutException:
                print(f"No se encontró el botón de descarga: {curp}")
                return False
            
        except Exception as e:
            print(f"Error procesando la CURP {curp}: {e}")
            return False

def main():
    # Check if Tesseract is installed

    if os.name == 'nt':
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    # else:
    #     pytesseract.pytesseract.tesseract_cmd = r'/home/.steamos/offload/var/lib/flatpak/app/io.github.manisandro.gImageReader/x86_64/stable/c9b91dfdb62d702811ea4e4e83e523e0e2db18b69001a26c3bc6ed25218a0829/files/bin/tesseract'
    
    try:
        pytesseract.get_tesseract_version()
    
    except Exception as e:
        print("Error: Tesseract OCR no está instalado o no se encontró.")
        print("Por favor, instala Tesseract OCR:")
        print("- Windows: Descargalo en https://github.com/UB-Mannheim/tesseract/wiki")
        print("- macOS: brew install tesseract")
        print("- Linux: sudo apt install tesseract-ocr")
        return
    
    # Check if PyMuPDF is installed
    try:
        import fitz
    except ImportError:
        print("Advertencia: PyMuPDF (fitz) not encontrado.")
        print("No se pueden procesar PDFs sin PyMuPDF.")
        print("Instalalo así: pip install PyMuPDF")
        print("\nTodavía puedes usar la aplicación para procesar CURPs en imágenes.")
    
    # Check if ChromeDriver is available
    try:
        # Try to create a webdriver instance briefly to test
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        #test_driver = webdriver.Chrome(options=options)
        test_driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        test_driver.quit()
    except Exception as e:
        print("Advertencia: No se encontró ChromeDriver.")
        print("No es posible descargar PDFs sin ChromeDriver.")
        print("Instala ChromeDriver:")
        print("- Descargalo en https://chromedriver.chromium.org/")
        print("- O usa: pip install webdriver-manager")
        print("\nEL programa puede procesar CURPs, más no descargarlas.")
    
    root = tk.Tk()
    app = CURPExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()