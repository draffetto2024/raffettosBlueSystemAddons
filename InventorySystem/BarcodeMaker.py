import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import sqlite3
import random
import string
import pandas as pd
import os

class ZPLGenerator:
    @staticmethod
    def create_barcode_label(barcode_id, product_info, batch_info):
        """Generate ZPL code for a barcode label"""
        return f"""^XA
^FO50,50^BY3
^BCN,100,Y,N,N
^FD{barcode_id}^FS
^FO50,200^A0N,30,30
^FD{product_info}^FS
^FO50,250^A0N,30,30
^FD{batch_info}^FS
^XZ"""

class AutocompleteEntry(ttk.Entry):
    def __init__(self, master, app_reference, completevalues=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_reference = app_reference
        self.completevalues = completevalues or []
        
        # Create suggestion listbox
        self.suggestion_list = tk.Listbox(master)
        self.suggestion_list.grid_remove()  # Hide initially
        
        # Bind events
        self.bind('<KeyRelease>', self.on_key_release)
        self.suggestion_list.bind('<<ListboxSelect>>', self.on_select)
        self.bind('<FocusOut>', lambda e: self.suggestion_list.grid_remove())
        
    def update_suggestions(self, key=None):
        search_term = self.get().lower()
        
        # If empty, hide suggestions
        if not search_term:
            self.suggestion_list.grid_remove()
            return
            
        # Filter suggestions
        matches = []
        for item in self.completevalues:
            if search_term in item.lower():
                matches.append(item)
                
        # Update listbox
        if matches:
            # Position listbox below entry and match its width
            width = self.winfo_width()
            
            # Update listbox content
            self.suggestion_list.delete(0, tk.END)
            for item in matches[:5]:  # Limit to 5 suggestions
                self.suggestion_list.insert(tk.END, item)
            
            # Configure listbox size and position
            self.suggestion_list.configure(width=0)  # Reset width to allow proper sizing
            self.suggestion_list.grid(row=self.grid_info()['row'] + 1, 
                                    column=self.grid_info()['column'],
                                    sticky='ew')
        else:
            self.suggestion_list.grid_remove()
            
    def on_key_release(self, event):
        # Don't show suggestions for special keys
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Home', 'End', 'Return'):
            return
        self.update_suggestions()
        
    def on_select(self, event):
        if self.suggestion_list.curselection():
            selected_item = self.suggestion_list.get(self.suggestion_list.curselection())
            # Extract just the product code (everything before the " - ")
            product_code = selected_item.split(" - ")[0]
            self.delete(0, tk.END)
            self.insert(0, product_code)
            self.suggestion_list.grid_remove()
            if hasattr(self.app_reference, 'update_selection'):
                self.app_reference.update_selection()

class BarcodeGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Barcode Generator")
        
        # Create output directory if it doesn't exist
        self.output_dir = "generated_barcodes"
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Store product data
        self.products_df = None
        
        # Database initialization
        self.init_database()
        
        # Setup UI
        self.setup_ui()

    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Excel file selection
        ttk.Label(self.main_frame, text="Products Excel File:").grid(row=0, column=0, sticky=tk.W)
        self.excel_path_var = tk.StringVar(value="products.xlsx")
        ttk.Entry(self.main_frame, textvariable=self.excel_path_var).grid(row=0, column=1, sticky=(tk.W, tk.E))
        ttk.Button(self.main_frame, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        ttk.Button(self.main_frame, text="Load Products", command=self.load_products).grid(row=0, column=3)
        
        # Search/Filter frame
        search_frame = ttk.LabelFrame(self.main_frame, text="Product Selection", padding="5")
        search_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, sticky=tk.W)
        self.search_entry = AutocompleteEntry(
            search_frame,
            app_reference=self,
            width=50
        )
        self.search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # Selected product info
        self.selected_info_var = tk.StringVar()
        ttk.Label(search_frame, textvariable=self.selected_info_var).grid(row=1, column=0, columnspan=2, sticky=tk.W)
        
        # Date selection
        ttk.Label(self.main_frame, text="Date:").grid(row=2, column=0, sticky=tk.W)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        self.date_entry = ttk.Entry(self.main_frame, textvariable=self.date_var)
        self.date_entry.grid(row=2, column=1, columnspan=3, sticky=(tk.W, tk.E))
        
        # Batch size
        ttk.Label(self.main_frame, text="Batch Size:").grid(row=3, column=0, sticky=tk.W)
        self.batch_size_var = tk.StringVar()
        self.batch_size_entry = ttk.Entry(self.main_frame, textvariable=self.batch_size_var)
        self.batch_size_entry.grid(row=3, column=1, columnspan=3, sticky=(tk.W, tk.E))
        
        # Generate button
        self.generate_button = ttk.Button(self.main_frame, text="Generate Barcodes",
                                        command=self.generate_barcodes)
        self.generate_button.grid(row=4, column=0, columnspan=4, pady=10)

        # Status text
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self.main_frame, textvariable=self.status_var).grid(row=5, column=0, columnspan=4)

        # Try to load products initially
        self.load_products()

    def update_selection(self):
        """Update the display when a product is selected"""
        product_code = self.search_entry.get()
        if product_code and self.products_df is not None:
            # Find the matching product
            match = self.products_df[self.products_df['product_code'] == product_code]
            if not match.empty:
                row = match.iloc[0]
                self.selected_info_var.set(f"Selected: {row['product_code']} - {row['product_name']}")
                return
        self.selected_info_var.set("")

    def browse_excel(self):
        """Open file dialog to select Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
            self.load_products()

    def load_products_from_excel(self):
        """Load product data from Excel file"""
        try:
            # Read the Excel file
            self.products_df = pd.read_excel(self.excel_path_var.get())
            
            # Ensure required columns exist
            if 'product_code' not in self.products_df.columns or 'product_name' not in self.products_df.columns:
                raise ValueError("Excel file must contain 'product_code' and 'product_name' columns")
            
            # Convert to list of formatted strings for dropdown
            products = [f"{row['product_code']} - {row['product_name']}" 
                       for _, row in self.products_df.iterrows()]
            return sorted(products)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load products from Excel: {str(e)}")
            return []

    def load_products(self):
        """Load products and update suggestions"""
        products = self.load_products_from_excel()
        self.search_entry.completevalues = products
        self.status_var.set(f"Loaded {len(products)} products")

    def init_database(self):
        """Initialize SQLite database with required tables"""
        conn = sqlite3.connect('barcode_system.db')
        c = conn.cursor()
        
        # Create barcodes table
        c.execute('''CREATE TABLE IF NOT EXISTS barcodes
                    (barcode_id TEXT PRIMARY KEY,
                     product_code TEXT,
                     batch_date TEXT,
                     batch_number INTEGER,
                     batch_total INTEGER,
                     zpl_code TEXT)''')
        
        conn.commit()
        conn.close()

    def generate_barcode_id(self):
        """Generate unique alphanumeric barcode ID"""
        chars = string.ascii_uppercase + string.digits
        return ''.join(random.choice(chars) for _ in range(10))

    def generate_barcodes(self):
        """Generate barcodes and store in database"""
        try:
            selection = self.search_entry.get()
            if not selection:
                messagebox.showerror("Error", "Please select a product")
                return
                
            product_code = selection.split(' - ')[0]
            batch_date = self.date_var.get()
            
            try:
                batch_size = int(self.batch_size_var.get())
                if batch_size <= 0:
                    raise ValueError("Batch size must be positive")
            except ValueError as e:
                messagebox.showerror("Error", "Please enter a valid batch size")
                return
            
            # Create batch directory
            batch_dir = os.path.join(self.output_dir, f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            os.makedirs(batch_dir, exist_ok=True)
            
            conn = sqlite3.connect('barcode_system.db')
            c = conn.cursor()
            
            for i in range(batch_size):
                barcode_id = self.generate_barcode_id()
                batch_number = i + 1
                
                # Create ZPL code
                product_info = f"Product: {product_code}"
                batch_info = f"Batch: {batch_number}/{batch_size} - Date: {batch_date}"
                zpl_code = ZPLGenerator.create_barcode_label(barcode_id, product_info, batch_info)
                
                # Save ZPL to file
                filename = f"barcode_{batch_number:03d}_of_{batch_size:03d}.zpl"
                filepath = os.path.join(batch_dir, filename)
                with open(filepath, 'w') as f:
                    f.write(zpl_code)
                
                # Store in database
                c.execute('''INSERT INTO barcodes 
                           (barcode_id, product_code, batch_date, batch_number, batch_total, zpl_code)
                           VALUES (?, ?, ?, ?, ?, ?)''',
                           (barcode_id, product_code, batch_date, batch_number, batch_size, zpl_code))
            
            conn.commit()
            conn.close()
            
            self.status_var.set(f"Generated {batch_size} barcodes in {batch_dir}")
            messagebox.showinfo("Success", 
                f"Generated {batch_size} barcodes\n"
                f"ZPL files saved in: {batch_dir}\n"
                "You can now send these files to your Zebra printer")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Error generating barcodes")

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeGeneratorApp(root)
    root.mainloop()