# -*- coding: utf-8 -*-
"""
Created on Tue Aug 15 16:54:11 2023

@author: Derek
"""

# -*- coding: utf-8 -*-

#the curly version of an apostrophe can cause issues with matching, fixed for now only in phrasequantity

"""
Created on Thu May 11 21:03:11 2023

@author: Derek
"""
import re
import time
import datetime
ACCEPTABLE_PHRASES_FILE = "acceptable_phrases.txt"
import tkinter as tk

import tkinter as tk
from tkinter import messagebox

import sqlite3

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk

import pandas as pd

import os
import sys

packaged_order_dict = {}
order_num = {}
keywords_2d_list = [] #order list of keyword lists. Each list is ordered with how the item should be displayed to match correctly
removals_list = []
quantityphrases = []
quantitypositions = []
incompletephrases = []
secondarykeywords = []
exactphrases = []
exactphraseitems_2d = []

# Get the directory where the script or executable is located
if getattr(sys, 'frozen', False):
    # If the application is frozen with PyInstaller, use this path
    application_path = os.path.dirname(sys.executable)
else:
    # Otherwise use the path to the script file
    application_path = os.path.dirname(os.path.abspath(__file__))

path_to_db = os.path.join(application_path, 'CopyPastePack.db')
path_to_xlsx = os.path.join(application_path, 'UPCCodes.xlsx')
path_to_txt = os.path.join(application_path, 'input.txt')

def initialize_database():
    """Initialize orders table if it doesn't exist"""
    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()
    
    # Create orders table if not exists with startpackingtimestamp
    c.execute('''CREATE TABLE IF NOT EXISTS orders (
        order_num INT,
        item VARCHAR(100),
        item_barcode INT,
        count INT,
        name VARCHAR(100),
        address VARCHAR(100),
        generatedtimestamp VARCHAR(100),
        packedtimestamp VARCHAR(100),
        startpackingtimestamp VARCHAR(100)
    )''')
    
    # Create ordersandtimestampsonly table if not exists
    c.execute('''CREATE TABLE IF NOT EXISTS ordersandtimestampsonly (
        order_num INT,
        generatedtimestamp VARCHAR(100),
        packedtimestamp VARCHAR(100),
        startpackingtimestamp VARCHAR(100)
    )''')
    
    conn.commit()
    conn.close()

def initialize_program():
    """Initialize the program by setting up database and loading today's orders"""
    global order_dict
    
    # First initialize database tables if they don't exist
    initialize_database()
    
    # Then load any existing orders for today
    order_dict = load_todays_orders()
    
    # Start the option selection
    choose_option()


# For handling datetime properly in SQLite3 (add at top of file with other imports)
def adapt_datetime(dt):
    return dt.isoformat()

def convert_datetime(s):
    return datetime.datetime.fromisoformat(s)

# Register the adapter and converter
sqlite3.register_adapter(datetime.datetime, adapt_datetime)
sqlite3.register_converter("timestamp", convert_datetime)




class AddItemDialog:
    def __init__(self, parent, app):
        self.top = tk.Toplevel(parent)
        self.app = app
        
        self.top.title("Add Item")
        
        # Create and pack widgets
        tk.Label(self.top, text="Scan item barcode:").pack(pady=5)
        
        self.barcode_var = tk.StringVar()
        self.barcode_entry = tk.Entry(self.top, textvariable=self.barcode_var)
        self.barcode_entry.pack(pady=5)
        
        tk.Label(self.top, text="Quantity:").pack(pady=5)
        
        self.quantity_var = tk.StringVar(value="1")
        self.quantity_entry = tk.Entry(self.top, textvariable=self.quantity_var)
        self.quantity_entry.pack(pady=5)
        
        tk.Button(self.top, text="Add", command=self.add).pack(pady=10)
        
        # Focus barcode entry
        self.barcode_entry.focus_set()
        
        # Make dialog modal
        self.top.transient(parent)
        self.top.grab_set()
        
    def add(self):
        barcode = self.barcode_var.get().strip()
        try:
            quantity = int(self.quantity_var.get())
            if quantity <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid quantity")
            return
            
        # Find item details from acceptable phrases
        item_found = False
        for phrase in read_acceptable_phrases():
            words = phrase.split()
            if words[-1] == barcode:
                item_name = " ".join(words[:-1])
                item_found = True
                break
                
        if not item_found:
            messagebox.showerror("Error", "Invalid barcode")
            return
            
        try:
            # Check if item already exists in order
            item_exists = False
            for i, (existing_item, existing_barcode, existing_count) in enumerate(self.app.current_order):
                if existing_barcode == barcode:
                    item_exists = True
                    new_count = existing_count + quantity
                    
                    # Update database
                    conn = sqlite3.connect(path_to_db)
                    cursor = conn.cursor()
                    
                    # Update existing item count
                    cursor.execute("""
                        UPDATE orders 
                        SET count = ?
                        WHERE order_num = ? 
                        AND item_barcode = ?
                        AND count > 0
                    """, (new_count, self.app.current_order_num, barcode))
                    
                    conn.commit()
                    conn.close()
                    
                    # Update in-memory order
                    self.app.current_order[i] = (existing_item, existing_barcode, new_count)
                    break
            
            if not item_exists:
                # Original add new item logic
                conn = sqlite3.connect(path_to_db)
                cursor = conn.cursor()
                now = datetime.datetime.now()
                
                cursor.execute("""
                    INSERT INTO orders 
                    (order_num, item, item_barcode, count, generatedtimestamp, startpackingtimestamp)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (self.app.current_order_num, item_name, barcode, quantity, now, now))
                
                conn.commit()
                conn.close()
                
                # Update current order in memory
                self.app.current_order.append((item_name, barcode, quantity))
            
            # Refresh display
            self.app.display_images(self.app.current_order)
            
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            messagebox.showerror("Error", "Failed to add item")
        finally:
            self.top.destroy()

class SwapItemDialog:
    def __init__(self, parent, app):
        self.top = tk.Toplevel(parent)
        self.app = app
        
        self.top.title("Swap Item")
        
        # Create and pack widgets
        tk.Label(self.top, text="Scan new item barcode:").pack(pady=5)
        
        self.barcode_var = tk.StringVar()
        self.barcode_entry = tk.Entry(self.top, textvariable=self.barcode_var)
        self.barcode_entry.pack(pady=5)

        # Add quantity field
        tk.Label(self.top, text="Quantity to swap:").pack(pady=5)
        self.quantity_var = tk.StringVar(value="1")
        self.quantity_entry = tk.Entry(self.top, textvariable=self.quantity_var)
        self.quantity_entry.pack(pady=5)
        
        tk.Button(self.top, text="Swap", command=self.swap).pack(pady=10)
        
        # Focus barcode entry
        self.barcode_entry.focus_set()
        
        # Make dialog modal
        self.top.transient(parent)
        self.top.grab_set()
        
    def swap(self):
        new_barcode = self.barcode_var.get().strip()
        old_item, old_barcode, old_count = self.app.selected_item

        # Prevent swapping with same item
        if new_barcode == old_barcode:
            messagebox.showerror("Error", "Cannot swap item with itself")
            return

        # Get and validate quantity
        try:
            swap_quantity = int(self.quantity_var.get())
            if swap_quantity <= 0 or swap_quantity > old_count:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid quantity (not exceeding current item count)")
            return
        
        # Find new item details from acceptable phrases
        item_found = False
        for phrase in read_acceptable_phrases():
            words = phrase.split()
            if words[-1] == new_barcode:
                new_item_name = " ".join(words[:-1])
                item_found = True
                break
                
        if not item_found:
            messagebox.showerror("Error", "Invalid barcode")
            return
            
        try:
            conn = sqlite3.connect(path_to_db)
            cursor = conn.cursor()
            now = datetime.datetime.now()
            
            # Check if new item already exists in order
            existing_item_index = None
            for i, (existing_item, existing_barcode, existing_count) in enumerate(self.app.current_order):
                if existing_barcode == new_barcode:
                    existing_item_index = i
                    new_count = existing_count + swap_quantity
                    
                    # Update existing item count
                    cursor.execute("""
                        UPDATE orders 
                        SET count = ?
                        WHERE order_num = ? 
                        AND item_barcode = ?
                        AND count > 0
                    """, (new_count, self.app.current_order_num, new_barcode))
                    break

            # Update old item quantity
            remaining_count = old_count - swap_quantity
            if remaining_count > 0:
                cursor.execute("""
                    UPDATE orders 
                    SET count = ?
                    WHERE order_num = ? 
                    AND item_barcode = ?
                """, (remaining_count, self.app.current_order_num, old_barcode))
            else:
                cursor.execute("""
                    UPDATE orders 
                    SET count = 0,
                        packedtimestamp = ?
                    WHERE order_num = ? 
                    AND item_barcode = ?
                """, (now, self.app.current_order_num, old_barcode))

            # If new item doesn't exist, insert it
            if existing_item_index is None:
                cursor.execute("""
                    INSERT INTO orders 
                    (order_num, item, item_barcode, count, generatedtimestamp, startpackingtimestamp)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (self.app.current_order_num, new_item_name, new_barcode, swap_quantity, now, now))
            
            conn.commit()
            
            # Update in-memory order
            # First update old item quantity or remove if zero
            if remaining_count > 0:
                for i, (item, barcode, count) in enumerate(self.app.current_order):
                    if barcode == old_barcode:
                        self.app.current_order[i] = (item, barcode, remaining_count)
                        break
            else:
                self.app.current_order = [i for i in self.app.current_order if i[1] != old_barcode]

            # Then update or add new item
            if existing_item_index is not None:
                self.app.current_order[existing_item_index] = (self.app.current_order[existing_item_index][0], 
                                                             new_barcode, 
                                                             self.app.current_order[existing_item_index][2] + swap_quantity)
            else:
                self.app.current_order.append((new_item_name, new_barcode, swap_quantity))
        
            # Clear selection and refresh display
            self.app.selected_item = None
            self.app.display_images(self.app.current_order)
            
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            messagebox.showerror("Error", "Failed to swap item")
        finally:
            conn.close()
            self.top.destroy()

class App():
    def __init__(self, master, order_dict):
        self.master = master
        self.order_dict = order_dict

        self.selected_item = None  # Add this line
        
        # Create main container frame
        self.main_container = tk.Frame(self.master)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Create fixed header frame
        self.header_frame = tk.Frame(self.main_container)
        self.header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create barcode entry section in header
        self.barcode_label = tk.Label(self.header_frame, text="Scan order barcode:", font=("Arial", 20))
        self.barcode_label.pack(side=tk.LEFT, padx=10)
        
        self.barcode_var = tk.StringVar()
        self.barcode_entry = tk.Entry(self.header_frame, textvariable=self.barcode_var, font=("Arial", 16))
        self.barcode_entry.pack(side=tk.LEFT, padx=10)
        self.barcode_entry.focus()
        self.barcode_entry.bind("<Return>", self.check_order)
        
        # Create buttons frame in header
        self.buttons_frame = tk.Frame(self.header_frame)
        self.buttons_frame.pack(side=tk.RIGHT, padx=10)
        
        # Create abort button
        self.abort_btn = tk.Button(self.buttons_frame, text="Abort Order", command=self.abort_order,
                                 height=2, font=("Arial", 12))
        self.abort_btn.pack(side=tk.LEFT, padx=5)

        # Add new buttons frame for order modification
        self.mod_buttons_frame = tk.Frame(self.buttons_frame)
        self.mod_buttons_frame.pack(side=tk.LEFT, padx=5)
        
        # Add Delete Item button
        self.delete_btn = tk.Button(self.mod_buttons_frame, text="Delete Item", 
                                command=self.delete_item,
                                height=2, font=("Arial", 12))
        self.delete_btn.pack(side=tk.LEFT, padx=5)
        
        # Add Add Item button
        self.add_btn = tk.Button(self.mod_buttons_frame, text="Add Item", 
                                command=self.add_item,
                                height=2, font=("Arial", 12))
        self.add_btn.pack(side=tk.LEFT, padx=5)
        
        # Add Swap Item button
        self.swap_btn = tk.Button(self.mod_buttons_frame, text="Swap Item", 
                                command=self.swap_item,
                                height=2, font=("Arial", 12))
        self.swap_btn.pack(side=tk.LEFT, padx=5)
        
        # Create remaining orders count button
        self.count_var = tk.StringVar()
        self.count_button = tk.Button(self.buttons_frame, textvariable=self.count_var,
                                    command=self.display_remaining, height=2, font=("Arial", 12))
        self.count_button.pack(side=tk.LEFT, padx=5)
        
        # Create scrollable canvas for grid of images
        self.canvas_frame = tk.Frame(self.main_container)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(self.canvas_frame)
        self.scrollbar = tk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        
        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack canvas and scrollbar
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Bind mouse wheel
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Initialize variables for image handling
        self.image_frames = []  # List to store frame references
        self.photo_references = []  # Keep references to prevent garbage collection
        self.item_frames = []  # Add this line to initialize item_frames list
        
        # Create default image
        self.default_photo = self.create_default_image((200, 200))
        
        # Load product data
        try:
            _, self.upc_to_image = read_excel_file(path_to_xlsx, return_mappings=True)
        except Exception as e:
            print(f"Error loading product data: {e}")
            self.upc_to_image = {}
            messagebox.showwarning("Warning", "Error loading product images. The packer will continue without images.")
        
        # Initialize other variables
        self.item_entry = None
        self.current_order = None
        
        self.update_count()
        
    def delete_item(self):
        """Handle item deletion"""
        if not hasattr(self, 'selected_item') or not self.selected_item:
            messagebox.showwarning("Warning", "Please select an item to delete")
            return
            
        item, barcode, count = self.selected_item
        
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete {item}?"):
            try:
                conn = sqlite3.connect(path_to_db)
                cursor = conn.cursor()
                now = datetime.datetime.now()
                
                # Set count to 0 for the existing item
                cursor.execute("""
                    UPDATE orders 
                    SET count = 0, 
                        packedtimestamp = ? 
                    WHERE order_num = ? 
                    AND item = ? 
                    AND item_barcode = ?
                """, (now, self.current_order_num, item, barcode))
                
                conn.commit()
                
                # Update current order in memory
                self.current_order = [i for i in self.current_order if i[1] != barcode]
                
                # Clear selection
                self.selected_item = None
                
                # Refresh display
                self.display_images(self.current_order)
                
                if not self.current_order:
                    self.complete_order()
                    
            except sqlite3.Error as e:
                print(f"Database error: {e}")
                messagebox.showerror("Error", "Failed to delete item")
            finally:
                conn.close()

    def add_item(self):
        """Handle adding new item to order"""
        AddItemDialog(self.master, self)


    def swap_item(self):
        """Handle item swapping"""
        if not hasattr(self, 'selected_item') or not self.selected_item:
            messagebox.showwarning("Warning", "Please select an item to swap")
            return
            
        SwapItemDialog(self.master, self)

    
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def load_product_data(self):
        """Load product data from Excel with error handling"""
        try:
            df = pd.read_excel(path_to_xlsx, engine='openpyxl', dtype=str)
            
            # Handle different possible column structures
            if "Item Name" not in df.columns or "UPC Code" not in df.columns:
                # Try numeric columns if named columns don't exist
                df.columns = ["Item Name", "UPC Code", "Image Path"] if len(df.columns) >= 3 else ["Item Name", "UPC Code"]
            
            # Ensure Image Path column exists
            if "Image Path" not in df.columns:
                df["Image Path"] = ""
            
            # Clean data
            df["Item Name"] = df["Item Name"].str.lower()
            df["UPC Code"] = df["UPC Code"].str.rstrip('.0')
            df["Image Path"] = df["Image Path"].fillna("")
            
            # Create mappings
            upc_to_image = dict(zip(df["UPC Code"], df["Image Path"]))
            items_dict = {row["Item Name"]: (row["UPC Code"], row["Image Path"]) 
                         for _, row in df.iterrows()}
            
            return items_dict, upc_to_image
            
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            raise

    def create_default_image(self, size=(200, 200)):
        """Create a default 'No Image Available' image"""
        try:
            # Create a new image with a light gray background
            img = Image.new('RGB', size, color='lightgray')
            draw = ImageDraw.Draw(img)
            
            # Add text
            text = "No Image\nAvailable"
            
            # Try to use a standard font, fall back to default if not available
            try:
                font = ImageFont.truetype("arial.ttf", 20)
            except:
                font = ImageFont.load_default()
            
            # Center the text
            text_bbox = draw.multiline_textbbox((0, 0), text, font=font, align="center")
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]
            x = (size[0] - text_width) / 2
            y = (size[1] - text_height) / 2
            
            # Draw the text in dark gray
            draw.multiline_text((x, y), text, fill='darkgray', font=font, align="center")
            
            return ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"Error creating default image: {e}")
            # Return None if we can't create the image
            return None

    def load_and_resize_image(self, image_path, target_size=(200, 200)):
        """Load and resize an image from path with error handling"""
        try:
            # Check if image path is empty or None
            if not image_path:
                return self.default_photo

            # Check if file exists
            if not os.path.exists(image_path):
                print(f"Image not found: {image_path}")
                return self.default_photo

            # Try to open and process the image
            with Image.open(image_path) as img:
                # Convert to RGB if necessary
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Resize image maintaining aspect ratio
                img.thumbnail(target_size, Image.Resampling.LANCZOS)
                
                # Create PhotoImage
                return ImageTk.PhotoImage(img)

        except (OSError, IOError) as e:
            print(f"Error loading image {image_path}: {e}")
            return self.default_photo
        except Exception as e:
            print(f"Unexpected error loading image {image_path}: {e}")
            return self.default_photo

    def display_images(self, items):
        """Display images in a grid layout with selectable text information"""
        # Clear existing images - modify this part
        for frame_tuple in self.image_frames:
            frame = frame_tuple[0]  # Get the frame widget from the tuple
            frame.destroy()
        self.image_frames.clear()
        self.photo_references.clear()
        
        # Configure grid columns
        grid_columns = 3  # Number of columns in the grid
        current_row = 0
        current_col = 0
        
        try:
            for item, barcode, count in items:
                # Create frame for this item
                item_frame = tk.Frame(self.scrollable_frame, relief=tk.RAISED, borderwidth=1)
                item_frame.grid(row=current_row, column=current_col, padx=10, pady=10, sticky="nsew")
                
                # Make the entire frame clickable
                item_frame.bind('<Button-1>', lambda e, i=item, b=barcode, c=count: self.select_item(e, i, b, c))
                
                # Get image
                image_path = self.upc_to_image.get(barcode, "")
                photo = self.load_and_resize_image(image_path)
                if photo:
                    self.photo_references.append(photo)
                
                # Create and pack image label
                image_label = tk.Label(item_frame, image=photo if photo else self.default_photo)
                image_label.image = photo if photo else self.default_photo
                image_label.pack(side=tk.LEFT, padx=5, pady=5)
                
                # Make image label clickable too
                image_label.bind('<Button-1>', lambda e, i=item, b=barcode, c=count: self.select_item(e, i, b, c))
                
                # Create frame for text information
                text_frame = tk.Frame(item_frame)
                text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
                
                # Create Text widget for all information
                text_widget = tk.Text(text_frame, wrap=tk.WORD, height=6, width=30)
                text_widget.pack(fill=tk.BOTH, expand=True, pady=5)
                
                # Configure Text widget
                text_widget.configure(font=("Arial", 12))
                text_widget.tag_configure("bold", font=("Arial", 12, "bold"))
                
                # Insert information with tags
                text_widget.insert(tk.END, item + "\n", "bold")
                text_widget.insert(tk.END, f"Quantity: {count}\n")
                text_widget.insert(tk.END, f"UPC: {barcode}")
                
                # Make read-only but still selectable
                text_widget.configure(state="disabled")
                
                # Make text widget clickable too
                text_widget.bind('<Button-1>', lambda e, i=item, b=barcode, c=count: self.select_item(e, i, b, c))
                
                self.image_frames.append((item_frame, item, barcode, count))
                
                # Highlight if this is the selected item
                if self.selected_item and self.selected_item == (item, barcode, count):
                    item_frame.configure(bg='lightblue')
                    text_widget.configure(bg='lightblue')
                    image_label.configure(bg='lightblue')
                    text_frame.configure(bg='lightblue')
                
                # Update grid position
                current_col += 1
                if current_col >= grid_columns:
                    current_col = 0
                    current_row += 1
                        
        except Exception as e:
            print(f"Error displaying images: {e}")
            messagebox.showwarning("Warning", "Error displaying some images.")

    def select_item(self, event, item, barcode, count):
        """Handle item selection"""
        # Update selected item
        self.selected_item = (item, barcode, count)
        
        # Reset all frames to default color
        for frame_tuple in self.image_frames:
            frame = frame_tuple[0]  # Get the frame widget
            frame.configure(bg='SystemButtonFace')
            for widget in frame.winfo_children():
                widget.configure(bg='SystemButtonFace')
        
        # Find and highlight the selected frame
        for frame_tuple in self.image_frames:
            frame, f_item, f_barcode, f_count = frame_tuple
            if (f_item, f_barcode, f_count) == self.selected_item:
                frame.configure(bg='lightblue')
                for widget in frame.winfo_children():
                    widget.configure(bg='lightblue')
                break
        
    def abort_order(self):
        """Handle order abortion"""
        self.current_order = None
        self.reset_window()
        self.barcode_entry.bind("<Return>", self.check_order)
    
    def display_remaining(self):
        """Display remaining orders in a popup window with selectable text"""
        try:
            # Create popup window
            popup = tk.Toplevel(self.master)
            popup.title("Remaining Orders")
            popup.geometry("400x600")
            
            # Create main text widget
            text_widget = tk.Text(popup, wrap=tk.WORD, font=("Arial", 12))
            scrollbar = tk.Scrollbar(popup, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            # Configure tags for formatting
            text_widget.tag_configure("header", font=("Arial", 14, "bold"), spacing3=10)
            text_widget.tag_configure("order_num", font=("Arial", 12, "bold"))
            text_widget.tag_configure("normal", font=("Arial", 12))
            
            # Get remaining orders from database
            conn = sqlite3.connect(path_to_db)
            cursor = conn.cursor()
            
            # Convert today's date to string in ISO format
            today = datetime.date.today().isoformat()
            
            cursor.execute("""
            SELECT DISTINCT order_num, COUNT(item) as item_count, SUM(count) as total_items
            FROM orders
            WHERE packedtimestamp IS NULL 
            AND DATE(generatedtimestamp) = DATE(?)
            GROUP BY order_num
            ORDER BY order_num
            """, (today,))
            
            orders = cursor.fetchall()
            
            # Insert header
            text_widget.insert(tk.END, "Remaining Orders\n\n", "header")
            
            if orders:
                for order_num, item_count, total_items in orders:
                    text_widget.insert(tk.END, f"Order: {order_num}\n", "order_num")
                    text_widget.insert(tk.END, f"Unique Items: {item_count}\n", "normal")
                    text_widget.insert(tk.END, f"Total Items: {total_items}\n", "normal")
                    text_widget.insert(tk.END, "-" * 40 + "\n\n", "normal")
            else:
                text_widget.insert(tk.END, "No remaining orders\n", "normal")
            
            conn.close()
            
            # Make text widget read-only but selectable
            text_widget.configure(state="disabled")
            
            # Add copy functionality
            def copy_selection(event=None):
                try:
                    selected_text = text_widget.get("sel.first", "sel.last")
                    popup.clipboard_clear()
                    popup.clipboard_append(selected_text)
                except tk.TclError:
                    pass  # No selection
                return "break"
            
            text_widget.bind("<Control-c>", copy_selection)
            
            # Pack widgets
            scrollbar.pack(side="right", fill="y")
            text_widget.pack(side="left", fill="both", expand=True)
            
            # Add close button
            close_button = tk.Button(popup, 
                                text="Close", 
                                command=popup.destroy,
                                font=("Arial", 12))
            close_button.pack(pady=10)
            
            # Make the popup modal
            popup.transient(self.master)
            popup.grab_set()
            
            # Center the popup
            popup.update_idletasks()
            width = popup.winfo_width()
            height = popup.winfo_height()
            x = (popup.winfo_screenwidth() // 2) - (width // 2)
            y = (popup.winfo_screenheight() // 2) - (height // 2)
            popup.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            print(f"Error displaying remaining orders: {e}")
            messagebox.showerror("Error", "Failed to display remaining orders")

    def update_count(self):
        """Update the count of remaining orders"""
        try:
            conn = sqlite3.connect(path_to_db)
            cursor = conn.cursor()
            
            # Convert today's date to string in ISO format
            today = datetime.date.today().isoformat()
            
            query = """
            SELECT COUNT(DISTINCT order_num) as order_count,
                COUNT(DISTINCT item) as item_count,
                SUM(count) as total_items
            FROM orders
            WHERE packedtimestamp IS NULL 
            AND DATE(generatedtimestamp) = DATE(?)
            """
            
            cursor.execute(query, (today,))
            order_count, item_count, total_items = cursor.fetchone()
            
            # Handle None values
            total_items = total_items or 0
            item_count = item_count or 0
            order_count = order_count or 0
            
            self.count_var.set(f"Orders: {order_count}\nItems: {total_items}")
            
            conn.close()
            
        except Exception as e:
            print(f"Error updating count: {e}")
            self.count_var.set("Count Error")
    
    def enter_package_number(self, event): # unused
       
        while True:
            package_number = self.package_var.get()
            if package_number:
                global packaged_order_dict
                # Get the current date and time
                now = datetime.datetime.now()
                
                # Format the date and time as mm/dd/yy and a time
                timestamp = now
                global order_num
                
                packaged_order_dict[order_num] = {"package_num": package_number, "packed_time": timestamp}

                print("Packed orders:")
                print(packaged_order_dict)
                
                self.package_entry.pack_forget()
                self.reset_window()
                self.barcode_entry.focus()
                self.barcode_entry.bind("<Return>", self.check_order)
                break
            self.master.update()

    def reset_window(self):
        """Reset the window state with proper error handling"""
        # Clear barcode entry
        self.barcode_var.set("")
        
        # Clear images and frames with robust error handling
        if hasattr(self, 'image_frames'):
            for frame_data in self.image_frames:
                try:
                    if isinstance(frame_data, tuple):
                        # Check if the first element is a frame
                        if isinstance(frame_data[0], tk.Frame):
                            frame_data[0].destroy()
                        else:
                            print(f"Warning: Expected Frame, got {type(frame_data[0])}")
                    elif isinstance(frame_data, tk.Frame):
                        frame_data.destroy()
                    else:
                        print(f"Warning: Unexpected frame data type: {type(frame_data)}")
                except Exception as e:
                    print(f"Error destroying frame: {str(e)}")
                    continue
                    
            self.image_frames.clear()

        if hasattr(self, 'photo_references'):
            self.photo_references.clear()
        
        # Reset entry widgets
        try:
            self.barcode_entry.pack(side=tk.LEFT, padx=10)
            self.barcode_label.pack(side=tk.LEFT, padx=10)
            self.barcode_entry.focus()
        except Exception as e:
            print(f"Error resetting entry widgets: {str(e)}")
        
        # Clear item entry if it exists
        if hasattr(self, 'item_entry') and self.item_entry is not None:
            try:
                self.item_entry.destroy()
                self.item_entry = None
            except Exception as e:
                print(f"Error destroying item entry: {str(e)}")
            
    def check_order(self, event):
        """Handle order barcode scanning"""
        order_num = self.barcode_var.get()
        if order_num in self.order_dict:
            self.current_order = self.order_dict[order_num]
            self.current_order_num = order_num  # Store for later use
            
            # Log start packing timestamp
            try:
                conn = sqlite3.connect(path_to_db)
                cursor = conn.cursor()
                now = datetime.datetime.now()
                
                # Update both tables with start packing timestamp
                cursor.execute("""
                    UPDATE orders 
                    SET startpackingtimestamp = ? 
                    WHERE order_num = ? AND startpackingtimestamp IS NULL
                """, (now, order_num))
                
                cursor.execute("""
                    UPDATE ordersandtimestampsonly 
                    SET startpackingtimestamp = ? 
                    WHERE order_num = ? AND startpackingtimestamp IS NULL
                """, (now, order_num))
                
                conn.commit()
                
            except sqlite3.Error as e:
                print(f"Database error when logging start time: {e}")
            finally:
                conn.close()
                
            # Display the order items
            self.display_images(self.current_order)
            
            # Hide barcode entry and show item entry
            self.barcode_entry.pack_forget()
            self.barcode_label.pack_forget()
            self.barcode_entry.unbind("<Return>")
            self.barcode_entry.delete(0, tk.END)
            
            # Create and setup item entry
            self.item_entry = tk.Entry(self.header_frame, font=("Arial", 16))
            self.item_entry.pack(side=tk.LEFT, padx=10)
            self.item_entry.focus()
            self.item_entry.bind("<Return>", self.scan_item)
        else:
            messagebox.showerror("Error", "Order not found")
            self.barcode_entry.delete(0, tk.END)


    def scan_item(self, event):
        """Handle item barcode scanning with real-time updates"""
        item_barcode = self.item_entry.get()
        remaining_counts = {item[1]: item[2] for item in self.current_order}
        count_remaining = remaining_counts.get(item_barcode)
    
        if count_remaining is None:
            tk.messagebox.showerror("Error", "Invalid item barcode")
            self.item_entry.delete(0, tk.END)
            return
    
        count_remaining -= 1
    
        if count_remaining == 0:
            self.current_order = [item for item in self.current_order if item[1] != item_barcode]
        else:
            self.current_order = [(item[0], item[1], count_remaining) if item[1] == item_barcode else item 
                                for item in self.current_order]
    
        self.item_entry.delete(0, tk.END)
        
        # Update display immediately
        if self.current_order:
            self.display_images(self.current_order)
        
        # Order has been filled
        if len(self.current_order) == 0:
            self.complete_order()
        
    def complete_order(self):
        """Handle order completion"""
        self.item_entry.unbind("<Return>")
        self.item_entry.destroy()
        self.item_entry = None
        
        # Update database
        conn = sqlite3.connect(path_to_db)
        cursor = conn.cursor()
        
        now = datetime.datetime.now()
        
        # Update orders
        cursor.execute("UPDATE orders SET packedtimestamp = ? WHERE order_num = ?",
                      (now, self.current_order_num))
        
        # Update timestamps
        cursor.execute("UPDATE ordersandtimestampsonly SET packedtimestamp = ? WHERE order_num = ?",
                      (now, self.current_order_num))
        
        conn.commit()
        conn.close()
        
        self.update_count()
        self.reset_window()
        
        # Clear the display
        for frame in self.image_frames:
            frame.destroy()
        self.image_frames.clear()
        self.photo_references.clear()
        
        self.barcode_entry.focus()
        self.barcode_entry.bind("<Return>", self.check_order)


    def run(self):
        #self.exit_btn.pack(side=tk.BOTTOM)
        self.master.mainloop()

# Add at start of program:
def add_startpacking_column():
    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()
    try:
        c.execute('ALTER TABLE orders ADD COLUMN startpackingtimestamp VARCHAR(100)')
    except sqlite3.OperationalError:
        pass  # Column already exists
    conn.close()

def read_excel_file(file_path, return_mappings=False):
    """
    Reads an Excel file with product information including optional image paths
    Args:
        file_path: Path to the Excel file
        return_mappings: If True, returns dictionaries for image mapping; if False, returns phrases list
    """
    try:
        # Read the Excel file into a dataframe
        df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
        
        # Get number of columns
        num_columns = len(df.columns)
        
        if num_columns == 2:
            # If only 2 columns, assume they are Item Name and UPC Code
            if "Item Name" not in df.columns or "UPC Code" not in df.columns:
                df.columns = ["Item Name", "UPC Code"]
            # Add empty Image Path column
            df["Image Path"] = ""
        elif num_columns >= 3:
            # If 3 or more columns, use first three
            df = df.iloc[:, :3]  # Take only first three columns
            df.columns = ["Item Name", "UPC Code", "Image Path"]
        else:
            raise ValueError(f"Excel file must have at least 2 columns. Found {num_columns} columns.")

        # Clean the data
        df["Item Name"] = df["Item Name"].str.lower()
        df["UPC Code"] = df["UPC Code"].str.rstrip('.0')
        df["Image Path"] = df["Image Path"].fillna("")

        if return_mappings:
            # Create mappings for image handling
            upc_to_image = dict(zip(df["UPC Code"], df["Image Path"]))
            items_dict = {row["Item Name"]: (row["UPC Code"], row["Image Path"]) 
                         for _, row in df.iterrows()}
            return items_dict, upc_to_image
        else:
            # Return phrases for order processing
            phrases = []
            for _, row in df.iterrows():
                item_name = row["Item Name"]
                upc_code = row["UPC Code"]
                phrases.append(f"{item_name} {upc_code}")
            return phrases

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        if return_mappings:
            return {}, {}
        return []

def start_packing_sequence(order_dict):
    root = tk.Tk()
    root.title("Order Packer")
    root.geometry("1600x800")
    app = App(root, order_dict)

    # Focus the window
    root.lift()  # Lift the window to the top
    root.attributes('-topmost', True)  # Make it topmost
    root.after_idle(root.attributes, '-topmost', False)  # Disable topmost after focusing
    root.focus_force()  # Force focus

    app.run()

def split_data(text, keywords):
    # Clean text and keywords with line-by-line stripping
    def clean_lines(text):
        """Helper to clean whitespace line by line"""
        return '\n'.join(line.strip() for line in text.splitlines())
    
    text = clean_lines(text)
    keywords = [clean_lines(kw) for kw in keywords if clean_lines(kw)]
    
    # Initialize the list to store the result
    results = []
    
    # While loop to handle multiple blocks of text
    while keywords:
        # Only consider keywords that are in the text
        present_keywords = [keyword for keyword in keywords if keyword in text]
        
        if not present_keywords:
            break
        
        # Sort keywords based on their order of appearance in the text
        present_keywords.sort(key=text.index)
        
        # Initialize a dictionary for this block of text
        block_results = {}
        
        # Loop through the keywords
        for keyword in present_keywords:
            parts = text.split(keyword, 1)
            if len(parts) > 1:
                before, after = parts
                # Clean the extracted text
                block_results[keyword] = clean_lines(after.split('\n', 1)[0])
                text = after
            else:
                break

        if block_results:
            results.append(block_results)
        keywords = [kw for kw in keywords if kw not in block_results]
    
    return results

def read_acceptable_phrases():
    ACCEPTABLE_PHRASES_FILE = path_to_xlsx

    return read_excel_file(ACCEPTABLE_PHRASES_FILE)

def print_blocks(blocks, block_separator="=" * 50):
    print("\nBlocks:")
    for i, block in enumerate(blocks, 1):
        print(f"\n{block_separator}")
        print(f"Block {i}:")
        print(f"{block_separator}")
        
        # Split the block into lines and print each line
        lines = block.split('\n')
        for line in lines:
            print(line.strip())

def split_blocks(text, keyword):
    def clean_lines(text):
        """Helper to clean whitespace line by line"""
        return '\n'.join(line.strip() for line in text.splitlines())

    # Clean both text and keyword
    text = clean_lines(text)
    keyword = clean_lines(keyword)
    
    # Split the text on the keyword
    parts = text.split(keyword)
    
    # The first part doesn't need the keyword prepended
    blocks = [clean_lines(parts[0])]
    
    # For all other parts, prepend the keyword and clean
    for part in parts[1:]:
        blocks.append(clean_lines(keyword + part))
    
    return blocks

def replace_special_characters(text: str) -> str:
    # Define character mappings
    character_mappings = {
        'â€™': "'",  # Smart apostrophe replacement
        # Add more mappings here as needed, for example:
        # 'â€"': '-',  # Em dash
        # 'â€œ': '"',  # Smart quotes (opening)
        # 'â€': '"',   # Smart quotes (closing)
    }

    # Replace each special character with its standard equivalent
    processed_text = text
    for special_char, standard_char in character_mappings.items():
        processed_text = processed_text.replace(special_char, standard_char)
    
    return processed_text

def enter_text(acceptable_phrases):
    with open(path_to_txt, "r") as file:
        text_input = file.read()
    order_dict = {}
    packaged_order_dict = {}
    
    def clean_lines(text):
        """Helper to clean whitespace line by line"""
        return '\n'.join(line.strip() for line in text.splitlines())

    
    text_input = clean_lines(text_input.lower())
    text_input = replace_special_characters(text_input)
    
    # Connect to the database
    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()   
    
    c.execute("SELECT DISTINCT * FROM blockseperator WHERE keywordstype = 'blockseperator'")
    rows = c.fetchall()
    for row in rows:
        blockseperatorkeyword = clean_lines(row[1].lower())
        packagenumberphrase = clean_lines(row[2].lower())
    
    conn.close()
    
    blocks = split_blocks(text_input, blockseperatorkeyword)

    for block in blocks:
        order_num = ""
        # Clean each phrase individually
        phrases = [clean_lines(phrase) for phrase in block.split("\n") if phrase.strip()]
        modified_phrases = []
        phrases_length = len(phrases)
        phrasequantity = 1

        # Extract order number with clean lines
        for phrase in phrases:
            clean_phrase = clean_lines(phrase)
            match = re.search(f"{re.escape(packagenumberphrase)} ([^\s]+)", clean_phrase)
            if match:
                order_num = ''.join([char for char in match.group(1) if char.isdigit()])
                print(f"Found order number: {order_num}")
                break
        
        if not order_num:
            print("Warning: No order number found for this block. Skipping...")
            continue

        # Process pairings with clean lines
        splitting_keywords_list = [None] * len(keywords_2d_list)
        keyword_positions = dict()

        block_lines = [clean_lines(line) for line in block.split("\n") if line.strip()]
        clean_block = '\n'.join(block_lines)

        # Find keywords in cleaned lines
        for i, line in enumerate(block_lines):
            for j, keywords_list in enumerate(keywords_2d_list):
                for keyword in keywords_list:
                    clean_keyword = clean_lines(keyword)
                    keyword_position = line.find(clean_keyword)
                    if keyword_position != -1:
                        current_position = (i, keyword_position)
                        if clean_keyword not in keyword_positions or current_position < keyword_positions[clean_keyword]:
                            keyword_positions[clean_keyword] = current_position

        # Process each keywords list
        for keyword, position in keyword_positions.items():
            for i, keywords_list in enumerate(keywords_2d_list):
                if keyword in [clean_lines(k) for k in keywords_list]:
                    if splitting_keywords_list[i] is None or position < keyword_positions[splitting_keywords_list[i]]:
                        splitting_keywords_list[i] = keyword

        # Pairing Logic with clean lines
        stringtobeadded = ""
        
        for i, keywords_list in enumerate(keywords_2d_list):
            splittingkeyword = splitting_keywords_list[i]
            if splittingkeyword and splittingkeyword in clean_block:
                sub_blocks = clean_block.split(splittingkeyword)
                sub_blocks = [clean_lines(splittingkeyword + sub_block) for sub_block in sub_blocks if sub_block.strip()]
                prev_sub_block = ""

                for sub_block in sub_blocks:
                    processed_block = split_data(sub_block, keywords_list)
                    for dict_ in processed_block:
                        for keyword in keywords_list:
                            clean_keyword = clean_lines(keyword)
                            if dict_ is not None and clean_keyword in dict_:
                                stringtobeadded += clean_lines(dict_[clean_keyword]) + " "
                    
                    # Find quantity with clean lines
                    phrasequantity = 1
                    last_line = ""
                    last_phrase = ""

                    prev_lines = [clean_lines(line) for line in prev_sub_block.splitlines()]
                    for line in prev_lines:
                        for quantityphrase in quantityphrases:
                            clean_quantphrase = clean_lines(quantityphrase)
                            if clean_quantphrase in line:
                                last_line = line
                                last_phrase = clean_quantphrase
                    
                    for quantindex, phrase in enumerate(quantityphrases):
                        clean_phrase = clean_lines(phrase)
                        if clean_phrase in last_phrase:
                            seperatedphrase = last_line.split()
                            try:
                                phrasequantity = int(seperatedphrase[int(quantitypositions[quantindex])])
                            except (IndexError, ValueError):
                                phrasequantity = 1
                                
                    prev_sub_block = sub_block
                    
                    # Clean removals and string
                    for removal in removals_list:
                        stringtobeadded = stringtobeadded.replace(clean_lines(removal), "")
                    stringtobeadded = ' '.join(stringtobeadded.split())
                    
                    # Add cleaned string to modified phrases
                    clean_string = clean_lines(stringtobeadded)
                    if clean_string:
                        for _ in range(int(phrasequantity)):
                            modified_phrases.append(clean_string)
                    stringtobeadded = ""

        print("PHRASES:", phrases)

        # Process incomplete phrases and exact matches
        i = 0
        incompletekeyword = ""
        while i < len(phrases):
            phrase = clean_lines(phrases[i])
            
            # Handle quantity phrases
            for quantindex, quantphrase in enumerate(quantityphrases):
                clean_quantphrase = clean_lines(quantphrase)
                if clean_quantphrase in phrase:
                    seperatedphrase = phrase.split()
                    try:
                        phrasequantity = int(seperatedphrase[int(quantitypositions[quantindex])])
                    except (IndexError, ValueError):
                        phrasequantity = 1
            
            # Handle incomplete phrases
            for incompindex, incomp in enumerate(incompletephrases):
                clean_incomp = clean_lines(incomp)
                print("PHRASE: ", phrase)
                print("INCOMP", incomp)
                if clean_incomp in phrase:
                    incompletekeyword = clean_lines(secondarykeywords[incompindex])
            
            if incompletekeyword:
                for _ in range(phrasequantity):
                    modified_phrases.append(clean_lines(phrase + " " + incompletekeyword))
            
            # Handle exact phrases
            for exactindex, exactphrase in enumerate(exactphrases):
                linebreaks = exactphrase.count("\n")
                remaining_lines = phrases[i:i + linebreaks + 1]
                blockforexactcheck = clean_lines("\n".join(remaining_lines))
                
                clean_exactphrase = clean_lines(exactphrase)
                if clean_exactphrase in blockforexactcheck:
                    for item in exactphraseitems_2d[exactindex]:
                        clean_item = clean_lines(item)
                        for _ in range(phrasequantity):
                            modified_phrases.append(clean_item)
            else:
                for _ in range(phrasequantity):
                    modified_phrases.append(clean_lines(phrase))
            
            i += 1

        # Match with acceptable phrases
        for modified_phrase in modified_phrases:
            clean_modified = clean_lines(modified_phrase)
            for acceptable_phrase in acceptable_phrases:
                clean_acceptable = clean_lines(acceptable_phrase)
                acceptable_words = clean_acceptable.split()[:-1]
                modified_words = clean_modified.split()
                
                if set(modified_words) == set(acceptable_words):
                    barcode = acceptable_phrase.split()[-1]
                    item = " ".join(acceptable_words)
                    
                    if order_num not in order_dict:
                        order_dict[order_num] = [(item, barcode, 1)]
                    else:
                        item_found = False
                        for i, (item_, barcode_, count) in enumerate(order_dict[order_num]):
                            if item_ == item and barcode_ == barcode:
                                item_found = True
                                order_dict[order_num][i] = (item_, barcode_, count + 1)
                                break
                        if not item_found:
                            order_dict[order_num].append((item, barcode, 1))

    return order_dict

def print_orders(order_dict):
    print("Order Dictionary:")
    print("{:<15} {:<30} {:<15} {:<10}".format('Order Number', 'Item', 'Barcode', 'Count'))
    print("=" * 75)

    for order_num, items in order_dict.items():
        if items:
            for item, barcode, count in items:
                print("{:<15} {:<30} {:<15} {:<10}".format(order_num, item, barcode, count))
        else:
            print("{:<15} {:<60}".format(order_num, "No items found for this order"))
        print("-" * 75)  # Add a separator line after each order
    print(f"\nTotal number of orders: {len(order_dict)}")

def display_phrases_as_table(phrases):
    # Print header
    print("| Item Name".ljust(50), "| UPC Code".ljust(20), "|")
    print("-" * 75)
    
    # Print each row
    for phrase in phrases:
        parts = phrase.split()
        upc = parts[-1]
        item_name = " ".join(parts[:-1])
        print("|", item_name.ljust(48), "|", upc.ljust(18), "|")

def display_todays_products():
    conn = None
    try:
        # Read all products from the UPC Excel file with header row
        df = pd.read_excel(path_to_xlsx, engine='openpyxl', dtype=str)
        
        # Create a view with just the first two columns
        valid_products = df.iloc[:, :2].copy()
        
        # Get column names from first two columns
        first_col = df.columns[0]  # First column
        second_col = df.columns[1]  # Second column
        
        # Filter out any rows with missing data in first two columns
        valid_products = valid_products[valid_products[first_col].notna() & 
                                      valid_products[second_col].notna()]
        
        # Filter out any rows that contain headers
        valid_products = valid_products[~valid_products[first_col].str.lower().isin(['item name', 'product name', 'name', 'upc', 'upc code'])]
        
        # Now we can safely rename the columns since we only have two
        valid_products.columns = ["Item Name", "UPC Code"]
        
        # Convert item names to lowercase to match database
        valid_products["Item Name"] = valid_products["Item Name"].str.lower()
        # Remove .0 from UPC codes if present and ensure they're strings
        valid_products["UPC Code"] = valid_products["UPC Code"].astype(str).str.rstrip('.0')
        
        # Connect to database and get today's orders
        conn = sqlite3.connect(path_to_db)
        cursor = conn.cursor()
        
        # Get today's date
        today = datetime.date.today()
        
        # Query to get product counts for today's orders
        query = """
        SELECT 
            o.item,
            SUM(o.count) as total_count,
            o.item_barcode
        FROM orders o
        WHERE DATE(o.generatedtimestamp) = ?
        AND o.packedtimestamp IS NULL
        GROUP BY o.item, o.item_barcode
        """
        
        cursor.execute(query, (today,))
        today_products = {item.lower(): (count, barcode) for item, count, barcode in cursor.fetchall()}
        
        # Calculate the maximum lengths for formatting
        max_item_length = max(len(str(item)) for item in valid_products["Item Name"])
        max_item_length = min(max_item_length, 50)  # Cap at 50 characters
        
        # Print header
        print("\nComplete Products Summary (In UPC Excel Order)")
        print("=" * (max_item_length + 35))
        print(f"{'Product Name'.ljust(max_item_length)} | {'Quantity'.center(10)} | {'UPC'.center(15)}")
        print("-" * (max_item_length + 35))
        
        # Print each product in the order from the Excel file
        total_items = 0
        total_active_products = 0
        
        for _, row in valid_products.iterrows():
            try:
                item_name = str(row["Item Name"]).strip()
                upc = str(row["UPC Code"]).strip()
                
                if not item_name or not upc:  # Skip if either is empty after stripping
                    continue
                    
                # Get count from today's orders, or 0 if not present
                count = 0
                if item_name in today_products:
                    count = today_products[item_name][0]
                    total_items += count
                    total_active_products += 1
                
                truncated_item = item_name[:max_item_length]
                print(f"{truncated_item.ljust(max_item_length)} | {str(count).center(10)} | {str(upc).center(15)}")
                
            except Exception as e:
                print(f"Error processing row: {e}")
                continue  # Skip any problematic rows
        
        # Print footer with totals
        print("-" * (max_item_length + 35))
        print(f"Total Products in Database: {len(valid_products)}")
        print(f"Products with Orders Today: {total_active_products}")
        print(f"Total Items to Pack: {total_items}")
        
    except sqlite3.Error as e:
        print(f"Database error: {e}")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()  # This will print the full error trace
    finally:
        if conn:
            conn.close()

def load_todays_orders():
    """Load any unpacked orders from today from the database in the format expected by start_packing_sequence"""
    try:
        print("Starting load_todays_orders()")
        conn = sqlite3.connect(path_to_db)
        cursor = conn.cursor()
        
        # Get today's date
        today = datetime.datetime.now().date()
        print(f"Today's date: {today}")
        
        # Query to get all unpacked orders from today
        query = """
        SELECT 
            o.order_num,
            o.item,
            o.item_barcode,
            o.count
        FROM orders o
        WHERE DATE(o.generatedtimestamp) = ?
        AND o.packedtimestamp IS NULL
        ORDER BY o.order_num
        """
        print(f"Executing query with date: {today}")
        cursor.execute(query, (today,))
        results = cursor.fetchall()
        print(f"Query returned {len(results)} rows")
        
        # Debug print some sample results
        if results:
            print("Sample of first result row:", results[0])
        
        # Create order dictionary in the original format
        order_dict = {}
        for order_num, item, barcode, count in results:
            print(f"Processing order: {order_num}, item: {item}, barcode: {barcode}, count: {count}")
            
            # Ensure order_num is a string
            order_num = str(order_num).strip()
            if order_num not in order_dict:
                order_dict[order_num] = []
                print(f"Created new order entry for order number: {order_num}")
                
            # Add the item tuple in the format (item_name, barcode, count)
            item_tuple = (str(item).strip(), str(barcode).strip(), int(count))
            order_found = False
            
            # Update count if item already exists
            for i, (existing_item, existing_barcode, existing_count) in enumerate(order_dict[order_num]):
                if existing_item == item_tuple[0] and existing_barcode == item_tuple[1]:
                    order_found = True
                    print(f"Found existing item in order {order_num}")
                    break
                    
            if not order_found:
                order_dict[order_num].append(item_tuple)
                print(f"Added new item to order {order_num}")
            
        print(f"Final order dictionary contains {len(order_dict)} orders")
        print("Order numbers in dictionary:", list(order_dict.keys()))
        return order_dict
        
    except sqlite3.Error as e:
        print(f"SQLite error in load_todays_orders: {e}")
        print(f"Error type: {type(e)}")
        return {}
    except Exception as e:
        print(f"Unexpected error in load_todays_orders: {e}")
        print(f"Error type: {type(e)}")
        import traceback
        traceback.print_exc()
        return {}
    finally:
        if conn:
            conn.close()
            print("Database connection closed")

# Step 3: Define a function to prompt user to choose an option and perform that option
def choose_option():
    global acceptable_phrases
    global order_dict
    global packaged_order_dict
    acceptable_phrases = read_acceptable_phrases()
    while True:
        # Removed redundant order loading since it's done at program start
        option = input("""Choose an option:
1. Upload today's orders from input.txt file
2. Start Packing Today's Orders
3. Check customer order numbers and related phrases
4. Delete today's orders
5. View UPC Codes
6. Exit
7. Today's Products at a Glance
8. Generate and print Pasta Cut Pick List\n""")

        
        if option == "4":
            confirmation = input("Are you sure you want to delete today's orders? Yes/No: ").lower()
            if confirmation in ["yes", "y"]:
                try:
                    conn = sqlite3.connect(path_to_db)
                    cursor = conn.cursor()
                    
                    # Get today's date at midnight (start of day)
                    today_start = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                    # Get tomorrow's date at midnight
                    tomorrow_start = today_start + datetime.timedelta(days=1)
                    
                    # Delete from orders table
                    cursor.execute("""
                        DELETE FROM orders 
                        WHERE generatedtimestamp >= ? 
                        AND generatedtimestamp < ?
                    """, (today_start, tomorrow_start))
                    orders_deleted = cursor.rowcount
                    
                    # Delete from ordersandtimestampsonly table
                    cursor.execute("""
                        DELETE FROM ordersandtimestampsonly 
                        WHERE generatedtimestamp >= ? 
                        AND generatedtimestamp < ?
                    """, (today_start, tomorrow_start))
                    timestamps_deleted = cursor.rowcount
                    
                    conn.commit()
                    
                    # Clear the in-memory order dictionary
                    order_dict = {}
                    
                    print(f"Successfully deleted {orders_deleted} orders and {timestamps_deleted} timestamps from today")
                    
                except sqlite3.Error as e:
                    print(f"Database error occurred: {e}")
                    conn.rollback()
                except Exception as e:
                    print(f"An error occurred: {e}")
                    conn.rollback()
                finally:
                    conn.close()
            else:
                print("Delete operation cancelled")
           
        elif option == "1":
            
           # Clear all global lists before loading new data
            global keywords_2d_list
            global removals_list
            global quantityphrases
            global quantitypositions
            global incompletephrases
            global secondarykeywords
            global exactphrases
            global exactphraseitems_2d
            
            # Reset all global lists
            keywords_2d_list = []
            removals_list = []
            quantityphrases = []
            quantitypositions = []
            incompletephrases = []
            secondarykeywords = []
            exactphrases = []
            exactphraseitems_2d = []
            
            # Connect to the SQLite database
            conn = sqlite3.connect(path_to_db)
            c = conn.cursor()
            
            try:
                # Load pairings
                c.execute("SELECT DISTINCT * FROM pairings WHERE keywordstype = 'pairing'")
                for row in c.fetchall():
                    keywordstype, keywords_string, removals_string = row
                    keywords_list = keywords_string.lower().split('<')
                    removals = removals_string.lower().split('<')
                    keywords_2d_list.append(keywords_list)
                    removals_list.extend(removals)
                
                # Load exact phrases
                c.execute("SELECT DISTINCT * FROM exactphrases WHERE keywordstype = 'exactphrase'")
                for row in c.fetchall():
                    keywordstype, exactphrases_string, exactphraseitems_string = row
                    exactphrases_list = exactphrases_string.lower().split('<')
                    exactitems = exactphraseitems_string.lower().split('<')
                    exactphrases.extend(exactphrases_list)
                    exactphraseitems_2d.append(exactitems)
                
                # Load quantities
                c.execute("SELECT DISTINCT * FROM quantitys WHERE keywordstype = 'Quantity'")
                for row in c.fetchall():
                    quantityphrase = row[1].lower()
                    quantityphrases.append(quantityphrase)
                    quantitypositions.append(int(row[2]))
                
                # Load incomplete phrases
                c.execute("SELECT DISTINCT * FROM incompletephrases WHERE keywordstype = 'Incomplete Phrase'")
                for row in c.fetchall():
                    incompletephrase = row[1].lower().strip()
                    secondarykeyword = row[2].lower().strip()
                    incompletephrases.append(incompletephrase)
                    secondarykeywords.append(secondarykeyword)
                
                # Process orders
                order_dict = enter_text(acceptable_phrases)
                
                # Delete existing orders for today before inserting new ones
                today = datetime.datetime.now().date()
                cursor = conn.cursor()
                cursor.execute("DELETE FROM orders WHERE DATE(generatedtimestamp) = ?", (today,))
                cursor.execute("DELETE FROM ordersandtimestampsonly WHERE DATE(generatedtimestamp) = ?", (today,))
                
                # Insert new orders
                generatedtimestamp = datetime.datetime.now()
                for order_num, items in order_dict.items():
                    for item, barcode, count in items:
                        cursor.execute("""
                            INSERT INTO orders 
                            (order_num, item, item_barcode, count, generatedtimestamp) 
                            VALUES (?, ?, ?, ?, ?)""",
                            (order_num, item, barcode, count, generatedtimestamp))
                    
                    cursor.execute("""
                        INSERT INTO ordersandtimestampsonly 
                        (order_num, generatedtimestamp) 
                        VALUES (?, ?)""",
                        (order_num, generatedtimestamp))
                
                conn.commit()
                print("Orders uploaded successfully!")
                
            except Exception as e:
                print(f"Error during upload: {str(e)}")
                conn.rollback()
            finally:
                conn.close()
        elif option == "2" or option == "3":
            # Check if there are any orders for today
            current_orders = load_todays_orders()
            if not current_orders:
                print("\nNo orders found for today. Please upload orders using option 1 first.")
                continue
            
            # Update order_dict with current orders
            order_dict = current_orders
            
            if option == "2":
                try:
                    start_packing_sequence(order_dict)
                except Exception as e:
                    print(f"Error starting packing sequence: {e}")
            else:  # option == "3"
                try:
                    print_orders(order_dict)
                except Exception as e:
                    print(f"\nError displaying orders: {str(e)}")
           
        elif option == "6":
            break
        elif option == "7":
            display_todays_products()
        elif option == "5":
            display_phrases_as_table(acceptable_phrases)
        elif option == "8":
            try:
                # [All your existing code for reading data and processing pasta cuts - keep everything the same until the file creation part]
                
                # Read all products from the UPC Excel file
                df = pd.read_excel(path_to_xlsx, engine='openpyxl', dtype=str)
                
                # Create a view with just the first two columns
                valid_products = df.iloc[:, :2].copy()
                
                # Get column names from first two columns
                first_col = df.columns[0]
                second_col = df.columns[1]
                
                # Filter out any rows with missing data in first two columns
                valid_products = valid_products[valid_products[first_col].notna() & 
                                            valid_products[second_col].notna()]
                
                # Filter out any rows that contain headers
                valid_products = valid_products[~valid_products[first_col].str.lower().isin(['item name', 'product name', 'name', 'upc', 'upc code'])]
                
                # Rename columns
                valid_products.columns = ["Item Name", "UPC Code"]
                
                # Convert item names to lowercase to match database
                valid_products["Item Name"] = valid_products["Item Name"].str.lower()
                valid_products["UPC Code"] = valid_products["UPC Code"].astype(str).str.rstrip('.0')
                
                # Filter for only pasta cuts
                pasta_cuts = ["linguine", "fettuccine", "pappardelle"]
                pasta_products = valid_products[valid_products["Item Name"].str.contains('|'.join(pasta_cuts), case=False)]
                
                # Connect to database and get today's orders
                conn = sqlite3.connect(path_to_db)
                cursor = conn.cursor()
                
                # Get today's date
                today = datetime.date.today()
                
                # Query to get product counts for today's orders
                query = """
                SELECT 
                    o.item,
                    SUM(o.count) as total_count
                FROM orders o
                WHERE DATE(o.generatedtimestamp) = ?
                AND o.packedtimestamp IS NULL
                GROUP BY o.item
                """
                
                cursor.execute(query, (today,))
                today_orders = {item.lower(): count for item, count in cursor.fetchall()}
                conn.close()
                
                # Group pasta by type and collect data
                pasta_by_type = {}
                
                for _, row in pasta_products.iterrows():
                    item_name = str(row["Item Name"]).strip()
                    upc = str(row["UPC Code"]).strip()
                    
                    # Get quantity from today's orders
                    quantity = today_orders.get(item_name, 0)
                    
                    if quantity > 0:  # Only include items with quantity > 0
                        # Determine pasta type and cut
                        for cut in pasta_cuts:
                            if cut in item_name:
                                # Extract type by removing the cut from the item name
                                pasta_type = item_name.replace(cut, "").strip()
                                
                                if pasta_type not in pasta_by_type:
                                    pasta_by_type[pasta_type] = []
                                
                                pasta_by_type[pasta_type].append({
                                    'name': item_name,
                                    'quantity': quantity,
                                    'upc': upc
                                })
                                break
                
                if pasta_by_type:
                    # Create the pick list content
                    content = []
                    content.append("PASTA CUTS PICK SHEET")
                    content.append("=" * 69)
                    content.append(f"{'Product Name'.ljust(35)} | {'Quantity'.center(10)} | {'UPC'.center(15)}")
                    content.append("-" * 69)
                    
                    # Sort types and add items
                    for pasta_type in sorted(pasta_by_type.keys()):
                        items = pasta_by_type[pasta_type]
                        items.sort(key=lambda x: x['name'])
                        
                        for item in items:
                            truncated_name = item['name'][:35]
                            content.append(f"{truncated_name.ljust(35)} | {str(item['quantity']).center(10)} | {str(item['upc']).center(15)}")
                        
                        # Add empty line between types (except for the last type)
                        if pasta_type != sorted(pasta_by_type.keys())[-1]:
                            content.append("")
                    
                    content.append("-" * 69)
                    
                    # Display the content
                    print("\n")
                    for line in content:
                        print(line)
                    
                    # Create a simple text file
                    filename = os.path.join(application_path, 'pasta_cuts_pick_sheet.txt')
                    
                    with open(filename, 'w') as file:
                        for line in content:
                            file.write(line + '\n')
                    
                    print(f"\nPasta cuts pick sheet saved as: {filename}")
                    
                    # Open Notepad with the file and then send Ctrl+P to open print dialog
                    import subprocess
                    import time
                    
                    try:
                        # Open the file in Notepad
                        process = subprocess.Popen(['notepad.exe', filename])
                        
                        # Wait a moment for Notepad to fully open
                        time.sleep(2)
                        
                        # Send Ctrl+P to open the print dialog
                        import pyautogui
                        pyautogui.hotkey('ctrl', 'p')
                        
                        print("Notepad opened with print dialog. Select your printer and print options!")
                        print("The print dialog should now be open for you to choose printer, copies, etc.")
                        
                    except ImportError:
                        print("pyautogui not installed. Install it with: pip install pyautogui")
                        print("Or manually press Ctrl+P when Notepad opens.")
                        # Still open notepad without the automatic Ctrl+P
                        subprocess.Popen(['notepad.exe', filename])
                        
                    except Exception as e:
                        print(f"Error opening Notepad or sending Ctrl+P: {e}")
                        print("Trying to open file normally...")
                        os.startfile(filename)
                    
                    # Exit the program
                    break
                    
                else:
                    print("No pasta cuts needed for today's orders.")
                    
            except Exception as e:
                print(f"Error generating pasta cuts pick sheet: {e}")
                import traceback
                traceback.print_exc()



        
        else:
            print("Invalid option. Try again.")

# Step 4: Call choose_option() function to start the program
initialize_program()