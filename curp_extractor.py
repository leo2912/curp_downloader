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
import time

class CURPExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CURP Extractor - Mexican Identity Document Processor")
        self.root.geometry("800x600")
        
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
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="CURP Extractor", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Upload button
        upload_btn = ttk.Button(main_frame, text="Upload Images", command=self.upload_images)
        upload_btn.grid(row=1, column=0, padx=(0, 10), pady=(0, 10), sticky=tk.W)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to process images")
        self.status_label.grid(row=1, column=2, pady=(0, 10), sticky=tk.E)
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Extracted CURPs", padding="10")
        results_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)
        
        # Selection controls frame
        selection_frame = ttk.Frame(results_frame)
        selection_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Select all/none buttons
        select_all_btn = ttk.Button(selection_frame, text="Select All", command=self.select_all_items)
        select_all_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        select_none_btn = ttk.Button(selection_frame, text="Select None", command=self.select_none_items)
        select_none_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download selected button
        download_selected_btn = ttk.Button(selection_frame, text="Download Selected PDFs", 
                                         command=self.download_selected_pdfs)
        download_selected_btn.pack(side=tk.LEFT)
        
        # Treeview for results
        columns = ('Select', 'File', 'CURP', 'Status')
        self.tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=15)
        
        # Define headings
        self.tree.heading('Select', text='☐')
        self.tree.heading('File', text='File Name')
        self.tree.heading('CURP', text='CURP Code')
        self.tree.heading('Status', text='Status')
        
        # Configure column widths
        self.tree.column('Select', width=50, minwidth=50)
        self.tree.column('File', width=200, minwidth=150)
        self.tree.column('CURP', width=200, minwidth=180)
        self.tree.column('Status', width=100, minwidth=80)
        
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
        self.context_menu.add_command(label="Download This CURP PDF", command=self.download_single_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Copy CURP", command=self.copy_single_curp)
        self.context_menu.add_command(label="Toggle Selection", command=self.toggle_single_selection)
        
        # Action buttons frame
        actions_frame = ttk.Frame(main_frame)
        actions_frame.grid(row=3, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # Copy button
        copy_btn = ttk.Button(actions_frame, text="Copy CURPs to Clipboard", command=self.copy_to_clipboard)
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download CSV button
        download_btn = ttk.Button(actions_frame, text="Download as CSV", command=self.download_csv)
        download_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear button
        clear_btn = ttk.Button(actions_frame, text="Clear Results", command=self.clear_results)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download PDFs button
        download_pdfs_btn = ttk.Button(actions_frame, text="Download CURP PDFs", command=self.download_curp_pdfs)
        download_pdfs_btn.pack(side=tk.LEFT)
        
        # Instructions label
        instructions = ("Instructions:\n"
                       "1. Click 'Upload Images' to select multiple image files\n"
                       "2. The app will extract 18-character CURP codes after 'Clave:'\n"
                       "3. View results in the table below\n"
                       "4. Use checkboxes to select CURPs, or right-click for individual actions\n"
                       "5. Copy codes to clipboard, download as CSV, or download official PDFs")
        
        instructions_label = ttk.Label(main_frame, text=instructions, font=("Arial", 9), 
                                     foreground="gray", justify=tk.LEFT)
        instructions_label.grid(row=4, column=0, columnspan=3, pady=(20, 0), sticky=tk.W)
        
    def upload_images(self):
        """Handle image upload and processing"""
        file_paths = filedialog.askopenfilenames(
            title="Select Images",
            filetypes=[
                ("Image files", "*.jpg *.jpeg *.png *.bmp *.tiff *.gif"),
                ("All files", "*.*")
            ]
        )
        
        if file_paths:
            # Process images in a separate thread to avoid blocking the UI
            threading.Thread(target=self.process_images, args=(file_paths,), daemon=True).start()
    
    def process_images(self, file_paths):
        """Process images and extract CURPs"""
        self.root.after(0, lambda: self.progress.start())
        self.root.after(0, lambda: self.status_label.config(text="Processing images..."))
        
        for i, file_path in enumerate(file_paths):
            try:
                # Update status
                filename = os.path.basename(file_path)
                self.root.after(0, lambda f=filename: self.status_label.config(text=f"Processing: {f}"))
                
                # Extract CURP from image
                curp = self.extract_curp_from_image(file_path)
                
                if curp:
                    # Add to results
                    self.root.after(0, lambda f=filename, c=curp: self.add_result(f, c, "Success"))
                    self.extracted_curps.append(curp)
                else:
                    self.root.after(0, lambda f=filename: self.add_result(f, "Not found", "Failed"))
                    self.extracted_curps.append(None)
                    
            except Exception as e:
                filename = os.path.basename(file_path)
                error_msg = str(e)[:50] + "..." if len(str(e)) > 50 else str(e)
                self.root.after(0, lambda f=filename, err=error_msg: self.add_result(f, f"Error: {err}", "Error"))
        
        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.status_label.config(text=f"Completed. Found {len([c for c in self.extracted_curps if c])} CURPs"))
    
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
            
        except Exception as e:
            print(f"Error processing image {image_path}: {e}")
            return None
    
    def add_result(self, filename, curp, status):
        """Add result to the treeview"""
        # Add checkbox symbol (unchecked by default)
        checkbox = "☐"
        item_id = self.tree.insert('', 'end', values=(checkbox, filename, curp, status))
        return item_id
    
    def copy_to_clipboard(self):
        """Copy all valid CURPs to clipboard"""
        valid_curps = [curp for curp in self.extracted_curps if curp and len(curp) == 18]
        
        if not valid_curps:
            messagebox.showwarning("No CURPs", "No valid CURPs found to copy.")
            return
        
        # Join CURPs with newlines
        curps_text = '\n'.join(valid_curps)
        
        # Copy to clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(curps_text)
        self.root.update()  # Ensure clipboard is updated
        
        messagebox.showinfo("Copied", f"Copied {len(valid_curps)} CURP codes to clipboard!")
    
    def download_csv(self):
        """Download results as CSV file"""
        if not self.tree.get_children():
            messagebox.showwarning("No Data", "No data to download.")
            return
        
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save CSV file"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # Write header
                    writer.writerow(['File Name', 'CURP Code', 'Status'])
                    
                    # Write data
                    for item in self.tree.get_children():
                        values = self.tree.item(item)['values']
                        # Skip the checkbox column when writing to CSV
                        if len(values) >= 4:
                            writer.writerow([values[1], values[2], values[3]])
                
                messagebox.showinfo("Success", f"CSV file saved successfully!\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV file:\n{e}")
    
    def clear_results(self):
        """Clear all results"""
        self.tree.delete(*self.tree.get_children())
        self.extracted_curps.clear()
        self.selected_items.clear()
        self.status_label.config(text="Ready to process images")
    
    def on_treeview_click(self, event):
        """Handle treeview click events for checkbox functionality"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x, event.y)
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
            if curp and len(str(curp)) == 18 and curp != "Not found":
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
                if curp and len(str(curp)) == 18 and curp != "Not found":
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
                if curp and len(str(curp)) == 18 and curp != "Not found":
                    # Start download in separate thread
                    threading.Thread(target=self.download_pdfs_worker, args=([curp],), daemon=True).start()
                else:
                    messagebox.showwarning("Invalid CURP", "Cannot download PDF for invalid CURP.")
    
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
                    messagebox.showinfo("Copied", f"CURP '{curp}' copied to clipboard!")
    
    def toggle_single_selection(self):
        """Toggle selection of the right-clicked item"""
        if hasattr(self, 'context_menu_item'):
            self.toggle_item_selection(self.context_menu_item)
    
    def download_selected_pdfs(self):
        """Download PDFs for selected CURPs only"""
        if not self.selected_items:
            messagebox.showwarning("No Selection", "Please select CURPs to download using the checkboxes.")
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
            messagebox.showwarning("No Valid CURPs", "No valid CURPs selected for download.")
            return
        
        # Confirm action
        result = messagebox.askyesno(
            "Confirm Download", 
            f"This will download {len(selected_curps)} selected CURP validation PDFs from gob.mx.\n\n"
            "This process may take several minutes. Continue?"
        )
        
        if result:
            # Start download process in separate thread
            threading.Thread(target=self.download_pdfs_worker, args=(selected_curps,), daemon=True).start()
    
    def setup_webdriver(self):
        """Setup Chrome WebDriver with download preferences"""
        try:
            # Ask user for download folder
            self.download_folder = filedialog.askdirectory(title="Select folder to save CURP PDFs")
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
            self.driver = webdriver.Chrome(options=chrome_options)
            return self.driver
            
        except Exception as e:
            messagebox.showerror("WebDriver Error", f"Failed to setup WebDriver:\n{e}\n\nMake sure ChromeDriver is installed.")
            return None
    
    def download_curp_pdfs(self):
        """Download official CURP PDFs for all valid CURPs"""
        valid_curps = [curp for curp in self.extracted_curps if curp and len(curp) == 18]
        
        if not valid_curps:
            messagebox.showwarning("No CURPs", "No valid CURPs found to download.")
            return
        
        # Confirm action
        result = messagebox.askyesno(
            "Confirm Download", 
            f"This will download {len(valid_curps)} CURP validation PDFs from gob.mx.\n\n"
            "This process may take several minutes. Continue?"
        )
        
        if not result:
            return
        
        # Start download process in separate thread
        threading.Thread(target=self.download_pdfs_worker, args=(valid_curps,), daemon=True).start()
    
    def download_pdfs_worker(self, curps):
        """Worker function to download PDFs"""
        self.root.after(0, lambda: self.progress.start())
        self.root.after(0, lambda: self.status_label.config(text="Setting up browser..."))
        
        # Setup WebDriver
        if not self.setup_webdriver():
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.status_label.config(text="Download cancelled"))
            return
        
        successful_downloads = 0
        failed_downloads = 0
        
        try:
            for i, curp in enumerate(curps):
                try:
                    self.root.after(0, lambda c=curp, idx=i+1, total=len(curps): 
                                  self.status_label.config(text=f"Downloading {idx}/{total}: {c}"))
                    
                    if self.download_single_curp_pdf(curp):
                        successful_downloads += 1
                    else:
                        failed_downloads += 1
                    
                    # Small delay between requests to be respectful
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"Error downloading CURP {curp}: {e}")
                    failed_downloads += 1
                    continue
        
        finally:
            # Cleanup
            if self.driver:
                self.driver.quit()
                self.driver = None
            
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, lambda: self.status_label.config(
                text=f"Downloads complete: {successful_downloads} successful, {failed_downloads} failed"))
            
            # Show completion message
            self.root.after(0, lambda: messagebox.showinfo(
                "Download Complete", 
                f"Downloaded {successful_downloads} CURP PDFs successfully.\n"
                f"{failed_downloads} downloads failed.\n\n"
                f"Files saved to: {self.download_folder}"
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
                print(f"Download button not found for CURP: {curp}")
                return False
            
        except Exception as e:
            print(f"Error processing CURP {curp}: {e}")
            return False

def main():
    # Check if Tesseract is installed
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:
        print("Error: Tesseract OCR is not installed or not found.")
        print("Please install Tesseract OCR:")
        print("- Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki")
        print("- macOS: brew install tesseract")
        print("- Linux: sudo apt install tesseract-ocr")
        return
    
    # Check if ChromeDriver is available
    try:
        # Try to create a webdriver instance briefly to test
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        test_driver = webdriver.Chrome(options=options)
        test_driver.quit()
    except Exception as e:
        print("Warning: ChromeDriver not found or not working properly.")
        print("PDF download feature will not work without ChromeDriver.")
        print("Install ChromeDriver:")
        print("- Download from https://chromedriver.chromium.org/")
        print("- Or use: pip install webdriver-manager")
        print("\nThe app will still work for CURP extraction without PDF download.")
    
    root = tk.Tk()
    app = CURPExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()