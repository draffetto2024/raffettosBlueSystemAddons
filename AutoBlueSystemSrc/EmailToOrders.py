import imaplib
import email
import email.parser
from email.header import decode_header
from bs4 import BeautifulSoup
from email.utils import parseaddr
import re
import pandas as pd
from nltk.tokenize import word_tokenize
import sqlite3
from datetime import datetime, timedelta
import time
from tkcalendar import DateEntry
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import scrolledtext
import pyautogui
import keyboard
import traceback

def read_grid_excel(file_path):
    df = pd.read_excel(file_path, header=0, index_col=0)
    
    # Initialize a dictionary to store customer-specific product codes
    customer_product_codes = {}
    
    for customer_email, row in df.iterrows():
        customer_codes = {}
        for product_type, value in row.items():
            if pd.notna(value):  # Check if the value is not NaN
                amount, product_code, unit = value.split(maxsplit=2)
                customer_codes[product_type.lower()] = (int(amount), product_code, unit.lower())
        customer_product_codes[customer_email.lower()] = customer_codes
    
    return customer_product_codes

def extract_orders(email_text, customer_codes):
    orders = []
    email_lines = email_text.strip().split('\n')
    print("\n=== Starting Order Extraction ===")
    print(f"Email lines:\n{email_lines}")

    for raw_text, (product_info, product_code, enters) in customer_codes.items():
        print(f"\nChecking match pattern:")
        print(f"Raw text: '{raw_text}'")
        pattern_lines = raw_text.split('\n')
        pattern_length = len(pattern_lines)
        print(f"Pattern length: {pattern_length} lines")

        # For each possible starting position in email
        for i in range(len(email_lines) - pattern_length + 1):
            print(f"\nChecking email chunk starting at line {i}:")
            chunk = email_lines[i:i+pattern_length]
            print(f"Email chunk: {chunk}")
            print(f"Pattern lines: {pattern_lines}")
            
            # Check if pattern matches this chunk
            matches = True
            for pattern_line, chunk_line in zip(pattern_lines, chunk):
                cleaned_pattern = clean_line(pattern_line)
                cleaned_chunk = clean_line(chunk_line)
                print(f"Comparing:")
                print(f"Pattern: '{cleaned_pattern}'")
                print(f"Chunk: '{cleaned_chunk}'")
                if cleaned_pattern not in cleaned_chunk:
                    matches = False
                    break
                    
            if matches:
                print("Match found!")
                chunk_text = '\n'.join(chunk)
                quantity = extract_quantity(chunk_text)
                orders.append((quantity, enters, chunk_text, product_code, quantity, 0))
                break

    print("\n=== Orders Found ===")
    for order in orders:
        print(f"Quantity: {order[0]}, Code: {order[3]}, Text: {order[2]}")

    return orders

def extract_quantity(text):
    quantity = re.search(r'\d+', text)
    return int(quantity.group()) if quantity else 1

def clean_line(line):
    return re.sub(r'[^a-zA-Z0-9 ]', '', line).lower()


def write_orders_to_db(db_path, customer_email, customer_id, orders, raw_email, email_sent_date):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    date_generated = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_processed = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_to_be_entered = datetime.now().strftime('%Y-%m-%d')

    for quantity, enters, raw_product, product_code, _, _ in orders:
        cursor.execute('''
        INSERT INTO orders (customer, customer_id, quantity, item, item_id, enters, date_generated, date_processed, entered_status, raw_email, email_sent_date, date_to_be_entered)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (customer_email, customer_id, quantity, raw_product, product_code, enters, date_generated, date_processed, 0, raw_email, email_sent_date, date_to_be_entered))

    conn.commit()
    conn.close()

# Function to extract email content
def get_email_content(msg):
    """Only processes forwarded emails as attachments"""
    print("\n=== Processing Email Content ===")
    attachment_data = get_attachment_content(msg)
    
    if not attachment_data['content']:
        print("No attachment content found")
        return {'content': None, 'from_header': None}
    
    print("\n=== Using Attachment Content ===")
    print(attachment_data['content'])
    if attachment_data['from_header']:
        print(f"\n=== Original From Header ===")
        print(attachment_data['from_header'])
    
    return attachment_data['content']

# Function to extract email address from the "From" field
# def extract_email_address(from_field):
#     name, email_address = parseaddr(from_field)
#     return email_address

def extract_email_address(email_content):
    # Look for the "From:" line in the email content
    from_line_match = re.search(r'From:.*?<(.+?)>', email_content, re.IGNORECASE | re.DOTALL)
    
    if from_line_match:
        return from_line_match.group(1)
    else:
        # If no match found, try to find any email address in the content
        email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email_content)
        if email_match:
            return email_match.group(0)
        else:
            logging.warning("No email address found in the content")
            return None

def cleanup_nomatch_entries(db_path, customer_id, email_sent_date):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
    DELETE FROM orders 
    WHERE customer_id = ? 
    AND email_sent_date = ? 
    AND item_id = 'NOMATCH'
    AND EXISTS (
        SELECT 1 
        FROM orders 
        WHERE customer_id = ? 
        AND email_sent_date = ? 
        AND item_id != 'NOMATCH'
    )
    ''', (customer_id, email_sent_date, customer_id, email_sent_date))
    
    conn.commit()
    conn.close()

def process_nomatch_entries(db_path, customer_product_codes):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get all NOMATCH entries
    cursor.execute('''
    SELECT raw_email, customer_id, email_sent_date
    FROM orders 
    WHERE item_id = 'NOMATCH'
    ''')
    nomatch_entries = cursor.fetchall()
    conn.close()
    
    for raw_email, customer_id, email_sent_date in nomatch_entries:
        customer_codes = customer_product_codes.get(customer_id, {})
        
        # Try to extract orders with current codes
        orders = extract_orders(raw_email, customer_codes)
        
        if orders:
            # Write new matches to database
            write_orders_to_db(db_path, customer_id, customer_id, orders, raw_email, email_sent_date)
            # Delete the NOMATCH entry
            delete_nomatch_entry(db_path, customer_id, email_sent_date)

def delete_nomatch_entry(db_path, customer_id, email_sent_date):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
    DELETE FROM orders 
    WHERE customer_id = ? 
    AND email_sent_date = ? 
    AND item_id = 'NOMATCH'
    ''', (customer_id, email_sent_date))
    conn.commit()
    conn.close()

# Modify the process_email function to call cleanup_nomatch_entries
def process_email(from_, body, customer_codes, customer_id, db_path, mail, msg_id, email_sent_date):
    BODY = body if not isinstance(body, dict) else body.get('content', '')
    
    # First, try to get existing orders
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
    SELECT COUNT(*) 
    FROM orders 
    WHERE customer_id = ? 
    AND email_sent_date = ? 
    AND item_id != 'NOMATCH'
    ''', (customer_id, email_sent_date))
    existing_orders = cursor.fetchone()[0]
    conn.close()

    # Extract new orders
    orders = extract_orders(BODY, customer_codes)

    # If we found matches or already have existing matches
    if orders or existing_orders > 0:
        if orders:  # If new matches found, write them
            write_orders_to_db(db_path, from_, customer_id, orders, BODY, email_sent_date)
        # Clean up any NOMATCH entries
        cleanup_nomatch_entries(db_path, customer_id, email_sent_date)
    else:  # No matches found and no existing matches
        write_orders_to_db(db_path, from_, customer_id, 
                          [(1, "4E", "NO MATCHES FOUND", "NOMATCH", 1, 0)], 
                          BODY, email_sent_date)
    
    if mail and msg_id:
        move_email(mail, msg_id, "EnteredIntoABS")
    
    return True

# Function to move an email to another folder
def move_email(mail, msg_id, destination_folder):
    result = mail.copy(msg_id, destination_folder)
    if result[0] == 'OK':
        mail.store(msg_id, '+FLAGS', '\\Deleted')
        mail.expunge()
    # else:
    #     #print(f"Failed to move email ID {msg_id} to {destination_folder}")

def initialize_database(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer TEXT,
        customer_id TEXT,
        quantity INTEGER,
        item TEXT,
        item_id TEXT,
        enters TEXT,
        date_generated TEXT,
        date_processed TEXT,
        entered_status INTEGER DEFAULT 0,
        raw_email TEXT,
        email_sent_date TEXT,
        date_to_be_entered TEXT
    )
    ''')
    
    conn.commit()
    conn.close()


def create_gui(db_path, email_to_customer_ids, product_enters_mapping):
    root = tk.Tk()
    root.title("Email Order Processor")
    root.geometry("1800x750")  # Slightly increased height to accommodate the new button layout

    # Create main frame
    main_frame = ttk.Frame(root)
    main_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

    # Create a frame for date selection
    date_frame = ttk.Frame(main_frame)
    date_frame.pack(pady=10, fill=tk.X)

    # Add date selection label and entry
    ttk.Label(date_frame, text="Select Date:").pack(side=tk.LEFT, padx=(0, 10))
    date_entry = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
    date_entry.pack(side=tk.LEFT)
    date_entry.set_date(datetime.now().date())  # Set default date to today

    # Add a button to refresh the grid based on the selected date
    refresh_button = ttk.Button(date_frame, text="Refresh", command=lambda: populate_grid(date_entry.get_date()))
    refresh_button.pack(side=tk.LEFT, padx=(10, 0))

    # Create a canvas with a scrollbar
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Pack the canvas and scrollbar
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Dictionary to store the state of checkboxes
    checkbox_vars = {}

    # Update headers to include Date to be Entered
    headers = ["Raw Email Content", "Matched Products", "Quantity", "Enters", "Product Codes", "Customer", "Email Sent Date", "Date to be Entered", "Entered Status", "Select", "Delete"]

    # Modify populate_grid to clean up NOMATCH entries when loading orders
    def populate_grid(filter_date=None):
        for widget in scrollable_frame.winfo_children():
            widget.destroy()

        # Update headers
        for col, header in enumerate(headers):
            label = ttk.Label(scrollable_frame, text=header, font=("Arial", 10, "bold"))
            label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Get all orders for the date range
        if filter_date:
            next_date = filter_date + timedelta(days=1)
            cursor.execute(''' 
            SELECT DISTINCT customer_id, email_sent_date 
            FROM orders 
            WHERE date_to_be_entered >= ? AND date_to_be_entered < ?
            ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
        else:
            cursor.execute('''
            SELECT DISTINCT customer_id, email_sent_date 
            FROM orders 
            WHERE date_to_be_entered IS NOT NULL
            ''')

        # Process NOMATCH entries after getting the initial data
        cursor.execute('''
        SELECT raw_email, customer_id, email_sent_date
        FROM orders 
        WHERE item_id = 'NOMATCH'
        ''')
        nomatch_entries = cursor.fetchall()
        
        for raw_email, customer_id, email_sent_date in nomatch_entries:
            customer_codes = customer_product_codes.get(customer_id, {})
            orders = extract_orders(raw_email, customer_codes)
            
            if orders:
                write_orders_to_db(db_path, customer_id, customer_id, orders, raw_email, email_sent_date)
                cursor.execute('''
                DELETE FROM orders 
                WHERE customer_id = ? 
                AND email_sent_date = ? 
                AND item_id = 'NOMATCH'
                ''', (customer_id, email_sent_date))
                conn.commit()

        # Now get the cleaned up orders for display
        if filter_date:
            cursor.execute(''' 
            SELECT raw_email, customer, item_id, quantity, enters, item, email_sent_date, 
                date_to_be_entered, entered_status, customer_id
            FROM orders
            WHERE date_to_be_entered >= ? AND date_to_be_entered < ?
            ORDER BY entered_status ASC, date_to_be_entered ASC
            ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
        else:
            cursor.execute('''
            SELECT raw_email, customer, item_id, quantity, enters, item, email_sent_date, 
                date_to_be_entered, entered_status, customer_id
            FROM orders
            WHERE date_to_be_entered IS NOT NULL
            ORDER BY entered_status ASC, date_to_be_entered ASC
            ''')

        orders = cursor.fetchall()
        conn.close()

        # Group orders by a unique identifier (raw_email + customer + email_sent_date)
        unique_orders = {}
        for order in orders:
            unique_key = (order[0], order[1], order[6])  # raw_email, customer, email_sent_date
            if unique_key not in unique_orders:
                unique_orders[unique_key] = []
            unique_orders[unique_key].append(order)

        for row, ((raw_email, customer, email_sent_date), order_list) in enumerate(unique_orders.items(), start=1):
            matched_products = []
            quantities = []
            enters_list = []
            product_codes = []
            customer_ids = set()
            date_to_be_entered = "N/A"
            entered_status = order_list[0][8]  # Last item is always entered_status

            for order in order_list:
                _, _, item_id, quantity, enters, item, _, date_to_be_entered, _, customer_id = order
                matched_products.append(item)
                quantities.append(str(quantity))
                enters_list.append(enters)
                product_codes.append(item_id)
                customer_ids.add(customer_id)

            # Create other UI elements as before
            products_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=30, height=5)
            products_text.insert(tk.END, "\n".join(matched_products))
            products_text.config(state=tk.DISABLED)
            products_text.grid(row=row, column=1, padx=5, pady=5, sticky="nsew")

            # Email content frame
            email_frame = ttk.Frame(scrollable_frame)
            email_frame.grid(row=row, column=0, padx=5, pady=5, sticky="nsew")

            email_text = tk.Text(email_frame, wrap=tk.WORD, width=30, height=5)
            email_text.insert(tk.END, raw_email)
            email_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            email_scrollbar = ttk.Scrollbar(email_frame, orient="vertical", command=email_text.yview)
            email_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            email_text.configure(yscrollcommand=email_scrollbar.set)
            email_text.config(state=tk.DISABLED)
            email_text.bind("<Button-1>", lambda e, text=raw_email: show_email_content(text))

            # Quantity
            quantity_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=10, height=5)
            quantity_text.insert(tk.END, "\n".join(quantities))
            quantity_text.config(state=tk.DISABLED)
            quantity_text.grid(row=row, column=2, padx=5, pady=5, sticky="nsew")

            # Enters
            enters_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=10, height=5)
            enters_text.insert(tk.END, "\n".join(enters_list))
            enters_text.config(state=tk.DISABLED)
            enters_text.grid(row=row, column=3, padx=5, pady=5, sticky="nsew")

            # Product Codes
            codes_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=20, height=5)
            codes_text.insert(tk.END, "\n".join(product_codes))
            codes_text.config(state=tk.DISABLED)
            codes_text.grid(row=row, column=4, padx=5, pady=5, sticky="nsew")

            # Customer - Just display the customer ID directly
            customer_label = ttk.Label(scrollable_frame, text=list(customer_ids)[0] if customer_ids else "Unknown")
            customer_label.grid(row=row, column=5, padx=5, pady=5, sticky="nsew")

            # Email Sent Date
            email_sent_date_label = ttk.Label(scrollable_frame, text=email_sent_date)
            email_sent_date_label.grid(row=row, column=6, padx=5, pady=5, sticky="nsew")

            # Date to be Entered
            date_to_be_entered_label = ttk.Label(scrollable_frame, text=date_to_be_entered)
            date_to_be_entered_label.grid(row=row, column=7, padx=5, pady=5, sticky="nsew")

            # Entered Status
            entered_status_label = ttk.Label(scrollable_frame, text="Entered" if entered_status else "Not Entered")
            entered_status_label.grid(row=row, column=8, padx=5, pady=5, sticky="nsew")

            # Add Delete button
            delete_button = ttk.Button(scrollable_frame, text="Delete", 
                                    command=lambda r=raw_email, c=list(customer_ids)[0], d=email_sent_date: delete_order(r, c, d))
            delete_button.grid(row=row, column=10, padx=5, pady=5, sticky="nsew")

            # Checkbox
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(scrollable_frame, variable=var)
            checkbox.grid(row=row, column=9, padx=5, pady=5, sticky="nsew")
            checkbox_vars[row] = var

    def show_email_content(email_text):
        # Create a new top-level window
        email_window = tk.Toplevel(root)
        email_window.title("Full Email Content")
        email_window.geometry("800x600")

        # Add a scrolled text widget to display the email content
        email_content = scrolledtext.ScrolledText(email_window, wrap=tk.WORD, width=80, height=30)
        email_content.pack(expand=True, fill='both', padx=10, pady=10)

        # Insert the email text and make it read-only
        email_content.insert(tk.INSERT, email_text)
        email_content.config(state='disabled')

        # Add a close button
        close_button = ttk.Button(email_window, text="Close", command=email_window.destroy)
        close_button.pack(pady=10)

    def add_matching():
        selected_rows = [row for row, var in checkbox_vars.items() if var.get()]
        if not selected_rows:
            messagebox.showwarning("Warning", "Please select a row to add matching.")
            return

        row = selected_rows[0]  # Use the first selected row
        customer_id = scrollable_frame.grid_slaves(row=row, column=5)[0]['text']
        
        # Updated way to get the email content
        email_frame = scrollable_frame.grid_slaves(row=row, column=0)[0]
        email_text_widget = email_frame.winfo_children()[0]
        raw_email = email_text_widget.get("1.0", tk.END).strip()
        email_sent_date = scrollable_frame.grid_slaves(row=row, column=6)[0]['text']

        # Create Quick Add window
        popup = tk.Toplevel(root)
        popup.title("Quick Add Matching")
        popup.geometry("400x200")

        tk.Label(popup, text="Matching Phrase:").pack()
        matching_phrase_entry = tk.Entry(popup, width=50)
        matching_phrase_entry.pack()

        tk.Label(popup, text="Matched Product:").pack()
        matched_product_entry = tk.Entry(popup, width=50)
        matched_product_entry.pack()

        def submit_matching():
            matching_phrase = matching_phrase_entry.get().strip().lower()
            matched_product = matched_product_entry.get().strip()

            if not matching_phrase or not matched_product:
                messagebox.showwarning("Warning", "Both fields must be filled.")
                return

            # Update Excel file and database, then confirm success without closing the popup
            if update_customer_excel_file(customer_id, matching_phrase, matched_product):
                reload_order(customer_id, raw_email, email_sent_date)
                messagebox.showinfo("Success", "Match added successfully. You can add another or close the window.")
            else:
                messagebox.showerror("Error", "Failed to add match. Please try again.")

        # Add the submit button
        submit_button = tk.Button(popup, text="Submit", command=submit_matching)
        submit_button.pack(pady=10)

    def update_excel_with_new_matching(customer_id, matching_phrase, matched_product):
        directory_path = 'customer_product_codes'
        for filename in os.listdir(directory_path):
            if filename.endswith('.xlsx') and customer_id in filename:
                file_path = os.path.join(directory_path, filename)
                df = pd.read_excel(file_path, header=None, names=['raw_text', 'product_info'])
                new_row = pd.DataFrame({'raw_text': [matching_phrase], 'product_info': [matched_product]})
                df = pd.concat([df, new_row], ignore_index=True)
                df.to_excel(file_path, index=False, header=False)
                break

    def reload_order(customer_id, raw_email, email_sent_date):
        global product_enters_mapping
        # Delete the order from the database
        delete_order_from_db(db_path, raw_email, customer_id, email_sent_date)

        # Reprocess the email content
        customer_product_codes, _ = read_customer_excel_files('customer_product_codes', product_enters_mapping)
        customer_codes = customer_product_codes.get(customer_id, {})
        if customer_codes:
            process_email(customer_id, raw_email, customer_codes, customer_id, db_path, None, None, email_sent_date)
        else:
            print(f"Warning: No customer codes found for customer {customer_id}")

        # Refresh the grid
        populate_grid(date_entry.get_date())

    def perform_auto_entry(orders):
        #print(f"Performing auto entry for {len(orders)} orders")
        # Sequence to enter the order entry screen
        into_order_sequence = [
            'KEY:1',
            'KEY:enter',
        ]
        # Execute the sequence to enter the order entry screen
        auto_order_entry(into_order_sequence)

        for i, order in enumerate(orders, 1):
            #print(f"Processing order {i} of {len(orders)}: {order}")

            # Extract customer ID and other details from the order
            raw_email, customer_id, email_sent_date, item_ids, quantities, enters, items = order

            #print(f"Customer ID: {customer_id}")
            #print(f"Email Sent Date: {email_sent_date}")
            #print(f"Item IDs: {item_ids}")
            #print(f"Quantities: {quantities}")
            #print(f"Enters: {enters}")

            pre_order_entry = [
                'KEY:enter',
                f'INPUT:{customer_id}',
                'KEY:enter',
                'KEY:enter', #no PO number
                'KEY:enter', #our van. It will auto fill
                'INPUT:N', #No to standard order
                'KEY:enter',
            ]
            # Execute the pre-order entry sequence
            auto_order_entry(pre_order_entry)

            # Generate the sequence for this specific order
            order_sequence = generate_order_sequence((customer_id, item_ids, quantities, enters, items))
            
            # Execute the sequence for this order
            auto_order_entry(order_sequence)

        #print("Auto entry completed")


    def auto_enter_selected():
        selected_orders = []
        already_entered_orders = []
        no_match_orders = []
        time.sleep(5)

        for row, var in checkbox_vars.items():
            if var.get():
                email_frame = scrollable_frame.grid_slaves(row=row, column=0)[0]
                email_text_widget = email_frame.winfo_children()[0]
                raw_email_content = email_text_widget.get("1.0", tk.END).strip()
                customer_id = scrollable_frame.grid_slaves(row=row, column=5)[0]['text']
                email_sent_date = scrollable_frame.grid_slaves(row=row, column=6)[0]['text']
                item_ids = scrollable_frame.grid_slaves(row=row, column=4)[0].get("1.0", tk.END).strip()
                
                if "NOMATCH" in item_ids:
                    no_match_orders.append((raw_email_content, customer_id, email_sent_date))
                    continue

                quantities = scrollable_frame.grid_slaves(row=row, column=2)[0].get("1.0", tk.END).strip()
                enters = scrollable_frame.grid_slaves(row=row, column=3)[0].get("1.0", tk.END).strip()
                items = scrollable_frame.grid_slaves(row=row, column=1)[0].get("1.0", tk.END).strip()
                
                entered_status = check_entered_status(db_path, raw_email_content, customer_id, email_sent_date)
                if entered_status == 1:
                    already_entered_orders.append((raw_email_content, customer_id, email_sent_date, item_ids, quantities, enters, items))
                else:
                    selected_orders.append((raw_email_content, customer_id, email_sent_date, item_ids, quantities, enters, items))

        if no_match_orders:
            messagebox.showwarning("Cannot Process", "Selected orders contain unmatched items that cannot be auto-entered.")
            return
            
        if selected_orders:
            perform_auto_entry(selected_orders)
            print("Stuff", [(order[0], order[1], order[2]) for order in selected_orders])
            update_entered_status(db_path, [(order[0], order[1], order[2]) for order in selected_orders])
            
            for order in selected_orders:
                move_email_by_content(order[0], order[1], order[2], "EnteredIntoABS", "ProcessedEmails")
            
            message = f"{len(selected_orders)} orders have been entered and marked as processed."
            if already_entered_orders:
                message += f"\n{len(already_entered_orders)} orders were already entered and skipped."
            
            messagebox.showinfo("Info", message)
            populate_grid(date_entry.get_date())
        elif already_entered_orders:
            messagebox.showwarning("Warning", f"All selected orders ({len(already_entered_orders)}) have already been entered.")
        else:
            messagebox.showwarning("Warning", "No rows selected.")


    def auto_enter_all():
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        filter_date = date_entry.get_date()
        next_date = filter_date + timedelta(days=1)
        
        cursor.execute('''
        SELECT raw_email, customer, customer_id, item_id, quantity, enters, item, entered_status, email_sent_date, date_to_be_entered
        FROM orders
        WHERE date_to_be_entered >= ? AND date_to_be_entered < ?
        ORDER BY entered_status ASC, email_sent_date ASC
        ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
        
        orders = cursor.fetchall()
        conn.close()

        unentered_orders = {}
        already_entered_orders = {}
        has_no_match = False

        for order in orders:
            raw_email, customer, customer_id, item_id, quantity, enters, item, entered_status, email_sent_date, date_to_be_entered = order
            order_key = (raw_email, customer_id, email_sent_date)
            
            # Check for NOMATCH entries
            if "NOMATCH" in item_id:
                has_no_match = True
                continue
                
            if entered_status == 0:
                if order_key not in unentered_orders:
                    unentered_orders[order_key] = [[], [], [], []]
                unentered_orders[order_key][0].append(item_id)
                unentered_orders[order_key][1].append(str(quantity) if quantity is not None else 'N/A')
                unentered_orders[order_key][2].append(enters if enters is not None else '4E')
                unentered_orders[order_key][3].append(item)
            else:
                if order_key not in already_entered_orders:
                    already_entered_orders[order_key] = [[], [], [], []]
                already_entered_orders[order_key][0].append(item_id)
                already_entered_orders[order_key][1].append(str(quantity) if quantity is not None else 'N/A')
                already_entered_orders[order_key][2].append(enters if enters is not None else '4E')
                already_entered_orders[order_key][3].append(item)

        if has_no_match:
            messagebox.showwarning("Cannot Process", "Some orders contain unmatched items that cannot be auto-entered. Please review and match these items first.")
            return

        unentered_orders_list = [
            (raw_email, customer_id, email_sent_date, '\n'.join(order_data[0]), '\n'.join(order_data[1]),
            '\n'.join(order_data[2]), '\n'.join(order_data[3]))
            for (raw_email, customer_id, email_sent_date), order_data in unentered_orders.items()
        ]

        if unentered_orders_list:
            perform_auto_entry(unentered_orders_list)
            update_entered_status(db_path, [(order[0], order[1], order[2]) for order in unentered_orders_list])
            
            for order in unentered_orders_list:
                move_email_by_content(order[0], order[1], order[2], "EnteredIntoABS", "ProcessedEmails")
            
            message = f"{len(unentered_orders_list)} orders have been entered and marked as processed."
            if already_entered_orders:
                message += f"\n{len(already_entered_orders)} orders were already entered and skipped."
            
            messagebox.showinfo("Info", message)
            populate_grid(date_entry.get_date())
        elif already_entered_orders:
            messagebox.showinfo("Info", f"All orders for the selected date ({len(already_entered_orders)}) have already been entered.")
        else:
            messagebox.showwarning("Warning", "No orders found for the selected date.")


    def delete_order(raw_email, customer_id, email_sent_date):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if the order has been entered
        cursor.execute('''
        SELECT entered_status
        FROM orders
        WHERE raw_email = ? AND customer_id = ? AND email_sent_date = ?
        LIMIT 1
        ''', (raw_email, customer_id, email_sent_date))
        
        result = cursor.fetchone()
        conn.close()
        
        if result is None:
            messagebox.showerror("Error", "Order not found in the database.")
            return
        
        entered_status = result[0]
        
        if entered_status == 1:  # Order has been entered
            if messagebox.askyesno("Confirm Delete", "This order has already been entered. Deleting it is unusual. Are you sure you want to proceed?"):
                if messagebox.askyesno("Final Confirmation", "This action cannot be undone. Are you absolutely sure you want to delete this entered order?"):
                    perform_delete(raw_email, customer_id, email_sent_date, entered=True)
        else:  # Order has not been entered
            if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this unentered order?"):
                perform_delete(raw_email, customer_id, email_sent_date, entered=False)

    def perform_delete(raw_email, customer_id, email_sent_date, entered):
        delete_order_from_db(db_path, raw_email, customer_id, email_sent_date)
        if entered:
            print("RAW_EMAIL", raw_email)
            move_email_by_content(raw_email, customer_id, email_sent_date, "ProcessedEmails", "Orders")
            messagebox.showinfo("Success", "Entered order deleted and email moved back to Orders inbox from ProcessedEmails.")
        else:
            move_email_by_content(raw_email, customer_id, email_sent_date, "EnteredIntoABS", "Orders")
            messagebox.showinfo("Success", "Unentered order deleted and email moved back to Orders inbox from EnteredIntoABS.")
        populate_grid(date_entry.get_date())  # Refresh the grid

    # Create a frame for buttons on the right side
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(side=tk.RIGHT, pady=10, padx=10, fill=tk.Y)

    # Add buttons to auto-enter selected rows and all unentered emails
    auto_enter_button = ttk.Button(button_frame, text="Auto Enter Selected Rows", command=auto_enter_selected)
    auto_enter_button.pack(side=tk.TOP, pady=(0, 5), fill=tk.X)

    auto_enter_all_button = ttk.Button(button_frame, text="Auto Enter All Unentered", command=auto_enter_all)
    auto_enter_all_button.pack(side=tk.TOP, pady=(0, 5), fill=tk.X)

    # Add new button for adding matching below the other buttons
    add_matching_button = ttk.Button(button_frame, text="Add Matching for Selected Order", command=add_matching)
    add_matching_button.pack(side=tk.TOP, pady=(0, 5), fill=tk.X)

    # Populate the grid with today's date
    populate_grid(datetime.now().date())

    root.mainloop()

def move_email_by_content(raw_email_content, customer_id, email_sent_date, source_folder, destination_folder):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        mail.login(username, password)
        mail.select(source_folder)
        
        status, messages = mail.search(None, "ALL")
        
        print("MESSAGES", messages)
        if status == 'OK' and messages[0]:
            for msg_id in messages[0].split():
                status, msg_data = mail.fetch(msg_id, "(RFC822)")
                if status == 'OK':
                    email_msg = email.message_from_bytes(msg_data[0][1])
                    email_body = get_email_content(email_msg)
                    email_date = email.utils.parsedate_to_datetime(email_msg['Date']).strftime('%Y-%m-%d %H:%M:%S')

                    print("EMAIL MSG", email_msg)
                    print("EMAIL_BODY", email_body)
                    
                    # Handle email_body being either dict or string
                    if isinstance(email_body, dict):
                        email_content = email_body.get('content', '')
                    else:
                        email_content = email_body
                        
                    if (email_content.strip() == raw_email_content.strip() and
                        email_date == email_sent_date):
                        move_email(mail, msg_id, destination_folder)
                        break
    finally:
        mail.close()
        mail.logout()


def check_entered_status(db_path, raw_email_content, customer_id, email_sent_date):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
    SELECT entered_status
    FROM orders
    WHERE raw_email = ? AND customer_id = ? AND email_sent_date = ?
    LIMIT 1
    ''', (raw_email_content, customer_id, email_sent_date))
    
    result = cursor.fetchone()
    conn.close()
    
    return result[0] if result else None


def read_customer_ids(file_path):
    df = pd.read_excel(file_path, header=None, names=['email', 'customer_id'])
    return dict(zip(df['email'].str.lower(), df['customer_id']))

def generate_order_sequence(order):
    #print(f"Generating order sequence for: {order}")
    customer_id, item_ids, quantities, enters, items = order
    
    sequence = []
    
    item_ids_list = item_ids.split('\n')
    quantities_list = quantities.split('\n')
    enters_list = enters.split('\n')
    
    for product_code, quantity, enters in zip(item_ids_list, quantities_list, enters_list):
        #print(f"Processing: Product Code: {product_code}, Quantity: {quantity}, Enters: {enters}")
        if enters.strip().endswith('E'):
            num_enters = int(enters.strip()[0])
            sequence.extend([
                f'INPUT:{product_code.strip()}',
                'KEY:enter',
                f'INPUT:{quantity.strip()}',
                'KEY:enter',
            ])
            sequence.extend(['KEY:enter'] * (num_enters - 1))  # Subtract 1 because we already added one 'enter'
        # else:
        #     #print(f"Warning: Unexpected enters format '{enters}' for item {product_code}")
    
    #END ORDER ROUTINE BELOW
    sequence.extend([
                'KEY:up',
                'KEY:enter',
                'WAIT:1.0', #PAUSE FOR INVOICE GENERATION
                'INPUT:T',
                'KEY:enter',
                'WAIT:2.0' #WAIT FOR FINAL CALCULATION
            ])

    #print(f"Generated sequence: {sequence}")
    return sequence


def auto_order_entry(keystroke_sequence):
    #print(f"Executing keystroke sequence: {keystroke_sequence}")
    for action in keystroke_sequence:
        #print(f"Executing action: {action}")
        if action.startswith('INPUT:'):
            # Input data
            data = action.split(':', 1)[1].upper()  # Convert to uppercase
            #print(f"Typing: {data}")
            keyboard.write(data)
        elif action.startswith('KEY:'):
            # Press a specific key
            key = action.split(':')[1]
            #print(f"Pressing key: {key}")
            keyboard.press_and_release(key)
        elif action.startswith('WAIT:'):
            # Wait for a specified number of seconds
            wait_time = float(action.split(':')[1])
            #print(f"Waiting for {wait_time} seconds")
            time.sleep(wait_time)
        else:
            # Assume it's text to type
            text = action.upper()  # Convert to uppercase
            #print(f"Typing: {text}")
            keyboard.write(text)
        
        time.sleep(0.25)  # Short pause between actions
    #print("Keystroke sequence completed")

def delete_order_from_db(db_path, raw_email, customer_id, email_sent_date):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
    DELETE FROM orders
    WHERE raw_email = ? AND customer_id = ? AND email_sent_date = ?
    ''', (raw_email, customer_id, email_sent_date))
    
    conn.commit()
    conn.close()

def update_entered_status(db_path, orders):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    for raw_email, customer_id, email_sent_date in orders:
        cursor.execute('''
        UPDATE orders
        SET entered_status = 1
        WHERE customer_id = ? AND email_sent_date = ?
        ''', (customer_id, email_sent_date))

    conn.commit()
    conn.close()

def move_email_back_to_orders(raw_email_content):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(username, password)
    mail.select("EnteredIntoABS")
    status, messages = mail.search(None, "ALL")
    messages = messages[0].split()
    for msg_id in messages:
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                body = get_email_content(msg)
                if body.strip() == raw_email_content:
                    move_email(mail, msg_id, "Orders")
    mail.close()
    mail.logout()

import os


def read_product_enters_mapping(file_path):
    print(f"Reading product enters mapping from {file_path}")
    
    try:
        # Try reading as Excel first
        df = pd.read_excel(file_path, header=None)
        if df.shape[1] < 7:
            print("Excel file doesn't have enough columns. Trying CSV...")
            # If Excel reading fails, try CSV
            df = pd.read_csv(file_path, header=None)
        
        print(f"The file has {df.shape[1]} columns.")
        print("Columns found:")
        for i, col in enumerate(df.columns):
            print(f"Column {i}: {col}")
        
        # Check if file has at least 7 columns
        if df.shape[1] < 7:
            print(f"Warning: The file has only {df.shape[1]} columns. Expected at least 7.")
            print("Unable to proceed with mapping. Please check the file.")
            return {}
        
        # Select columns A and G (index 0 and 6) since we now expect 7 columns
        df = df.iloc[:, [0, 6]]
        df.columns = ['product_code', 'enters']
        
        print("Using column A for product codes and column G for enters")
        
        # Create the mapping
        mapping = dict(zip(df['product_code'].astype(str).str.upper(), df['enters']))
        
        # Fill missing enters with '4E' and ensure all values are strings
        mapping = {k: (str(v) if pd.notna(v) else '4E') for k, v in mapping.items()}
        
        print(f"Loaded {len(mapping)} product code to enters mappings")
        print("Sample of product_enters_mapping:")
        for i, (code, enters) in enumerate(list(mapping.items())[:5]):
            print(f"  {code}: {enters}")
        if len(mapping) > 5:
            print("  ...")
        
        # Debug: Print the first few rows of the DataFrame
        print("\nFirst few rows of the data:")
        print(df.head())
        
        return mapping
    
    except Exception as e:
        print(f"Error reading file: {str(e)}")
        print(f"File path: {os.path.abspath(file_path)}")
        return {}

def read_customer_excel_files(directory_path, product_enters_mapping):
    print(f"Reading customer Excel files from {directory_path}")
    customer_product_codes = {}
    email_name_to_customer_ids = {}  # Changed from email_to_customer_ids
    
    for filename in os.listdir(directory_path):
        if filename.endswith('.xlsx'):
            try:
                # Parse filename that contains name and email: john_doe_(johndoe@gmail.com)-johndo.xlsx
                match = re.match(r'(.+?)\((.+?)\)-(.+?)\.xlsx$', filename)
                if not match:
                    print(f"Warning: Invalid filename format for {filename}")
                    continue
                
                name = match.group(1).strip('_').replace('_', ' ')  # Convert john_doe to john doe
                email = match.group(2).lower()
                customer_id = match.group(3)
                
                file_path = os.path.join(directory_path, filename)
                customer_codes = read_single_customer_file(file_path, customer_id, product_enters_mapping)
                
                customer_product_codes[customer_id] = customer_codes
                
                # Create composite key from name and email
                key = (name.lower(), email)
                email_name_to_customer_ids[key] = customer_id
                
                print(f"Processed {len(customer_codes)} product codes for customer {customer_id}")
            except ValueError as e:
                print(f"Warning: Invalid filename format for {filename}: {e}")
                continue
    
    return customer_product_codes, email_name_to_customer_ids


def read_single_customer_file(file_path, customer_id, product_enters_mapping):
    print(f"Reading file: {file_path}")
    df = pd.read_excel(file_path, header=None, names=['raw_text', 'product_info'])
    
    customer_codes = {}
    for _, row in df.iterrows():
        raw_text = row['raw_text'].strip()
        product_info = str(row['product_info'])
        
        parts = product_info.split()
        if len(parts) >= 2:
            product_code = parts[1].upper()
            enters = product_enters_mapping.get(product_code)
            if enters is None:
                print(f"  Warning: No matching enters found for product code {product_code}. Defaulting to 4E.")
                enters = '4E'
            else:
                print(f"  Matched: {product_code} -> {enters}")
        else:
            print(f"  Warning: Unexpected format in '{product_info}' for customer {customer_id}")
            product_code = "000000"
            enters = "4E"
        
        customer_codes[raw_text] = (product_info, product_code, enters)
    
    print(f"Processed {len(customer_codes)} entries for customer {customer_id}")
    return customer_codes

def update_customer_excel_file(customer_id, matching_phrase, matched_product):
    directory_path = 'customer_product_codes'
    
    for filename in os.listdir(directory_path):
        if filename.endswith('.xlsx'):
            # Check if any part of the customer_id is in the filename
            if any(part.lower() in filename.lower() for part in customer_id.split()):
                file_path = os.path.join(directory_path, filename)
                df = pd.read_excel(file_path, header=None, names=['raw_text', 'product_info'])
                new_row = pd.DataFrame({'raw_text': [matching_phrase], 'product_info': [matched_product]})
                df = pd.concat([df, new_row], ignore_index=True)
                df.to_excel(file_path, index=False, header=False)
                print(f"Updated Excel file {filename} for customer {customer_id} with new matching: {matching_phrase} -> {matched_product}")
                return True
    
    print(f"Warning: Excel file for customer {customer_id} not found. New matching not added.")
    return False


def ensure_imap_folder_exists(mail, folder_name):
    try:
        folders = [mailbox.decode().split('"/"')[-1] for mailbox in mail.list()[1]]
        if folder_name not in folders:
            result, data = mail.create(folder_name)
            if result == 'OK':
                logging.info(f"Folder {folder_name} created successfully")
            elif b'[ALREADYEXISTS]' in data[0]:
                logging.info(f"Folder {folder_name} already exists")
            else:
                logging.error(f"Failed to create folder {folder_name}: {data}")
                return False
        else:
            logging.info(f"Folder {folder_name} already exists")
        return True
    except Exception as e:
        logging.error(f"Error checking/creating folder {folder_name}: {str(e)}")
        return False

def move_email_to_folder(mail, email_uid, destination_folder):
    try:
        result, data = mail.uid('copy', email_uid, destination_folder)
        if result == 'OK':
            mail.uid('store', email_uid, '+FLAGS', '\\Deleted')
            mail.expunge()
            logging.info(f"Email moved to {destination_folder} successfully")
            return True
        else:
            logging.error(f"Failed to move email to {destination_folder}: {data}")
            return False
    except Exception as e:
        logging.error(f"Error moving email to {destination_folder}: {str(e)}")
        return False

def check_imap_connection(mail):
    try:
        status, response = mail.noop()
        return status == 'OK'
    except:
        return False

def extract_email_and_name(email_data):
    """Extract email and name from both forwarded and original email formats"""
    print("\n=== Extracting Name and Email ===")
    
    # First check the From header if available
    if email_data.get('from_header'):
        print(f"Using From header: {email_data['from_header']}")
        print("Here")
        name, email = parseaddr(email_data['from_header'])
        print("NAME", name)
        print("EMAIL", email)
        if email:
            print(f"Found in header - Name: {name}, Email: {email}")
            return name, email
    
    # If no header or header parsing failed, check content
    email_content = email_data.get('content', '')
    print("Checking content first few lines:")
    print('\n'.join(email_content.split('\n')[:6]))
    
    # Try original format
    original_pattern = r'^([^<]+)\s*<([^>]+)>'
    original_match = re.search(original_pattern, email_content, re.MULTILINE)
    
    # Try forwarded format
    forward_pattern = r'From:\s*([^<]+)\s*<([^>]+)>'
    forward_match = re.search(forward_pattern, email_content, re.MULTILINE)
    
    if original_match:
        name = original_match.group(1).strip()
        email = original_match.group(2).strip()
        print(f"Found in original format - Name: {name}, Email: {email}")
        return name, email
    elif forward_match:
        name = forward_match.group(1).strip()
        email = forward_match.group(2).strip()
        print(f"Found in forwarded format - Name: {name}, Email: {email}")
        return name, email
    
    # Fallback: any email
    email_pattern = r'<([^>]+@[^>]+)>'
    email_match = re.search(email_pattern, email_content)
    if email_match:
        email = email_match.group(1)
        print(f"Found email only: {email}")
        return None, email
    
    print("No name/email found")
    return None, None

def get_attachment_content(msg):
    """Extract content and metadata from email attachments (.eml files)"""
    print("\n=== Processing Email Attachment ===")
    
    attachment_data = {'content': None, 'from_header': None}
    
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'message/rfc822':
                print("Found RFC822 attachment")
                attached_email = part.get_payload()[0]
                
                # Get the From header from the attached email
                attachment_data['from_header'] = attached_email.get('From')
                print(f"From header in attachment: {attachment_data['from_header']}")
                
                # Try HTML first
                for subpart in attached_email.walk():
                    if subpart.get_content_type() == 'text/html':
                        print("Found HTML content in attachment")
                        html_content = subpart.get_payload(decode=True).decode()
                        soup = BeautifulSoup(html_content, 'html.parser')
                        
                        # Extract data from table rows
                        rows = soup.find_all('tr')
                        if rows:  # If we found table rows
                            order_text = []
                            print("\n=== Processing Table Rows ===")
                            for row in rows:
                                cells = row.find_all(['td', 'th'])
                                if cells:
                                    row_text = [cell.get_text(strip=True) for cell in cells]
                                    if any(text for text in row_text):
                                        line = ' '.join(row_text)
                                        print(f"Row content: {line}")
                                        order_text.append(line)
                            
                            if order_text:  # If we successfully extracted table data
                                attachment_data['content'] = '\n'.join(order_text)
                                return attachment_data
                
                # If no HTML tables found, try plain text but structure it
                for subpart in attached_email.walk():
                    if subpart.get_content_type() == 'text/plain':
                        print("Processing plain text content")
                        text_content = subpart.get_payload(decode=True).decode()
                        lines = text_content.split('\n')
                        structured_lines = []
                        
                        # Group lines that seem to be part of the same item
                        current_item = []
                        for line in lines:
                            line = line.strip()
                            if not line:  # Skip empty lines
                                continue
                            if line.startswith('R') or 'RAVIOLI' in line:  # Potential product line
                                if current_item:  # Save previous item if exists
                                    structured_lines.append(' '.join(current_item))
                                current_item = [line]
                            elif current_item:  # Continuation of current item
                                current_item.append(line)
                            else:  # Header or other information
                                structured_lines.append(line)
                        
                        # Add last item if exists
                        if current_item:
                            structured_lines.append(' '.join(current_item))
                        
                        attachment_data['content'] = '\n'.join(structured_lines)
                        return attachment_data
                        
    return attachment_data


import logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

if __name__ == "__main__":
    global product_enters_mapping
    product_enters_mapping = read_product_enters_mapping('ProductSheetWithEnters.xlsx')

    # Read customer-specific product codes from individual Excel files
    customer_product_codes, email_name_to_customer_ids = read_customer_excel_files('customer_product_codes', product_enters_mapping)

    # Email details
    username = "gingoso2@gmail.com"
    password = "soiz avjw bdtu hmtn"  # utilizing an app password

    # Database path
    db_path = 'orders.db'

    # Initialize the database
    initialize_database(db_path)

    # Connect to the server
    mail = imaplib.IMAP4_SSL("imap.gmail.com")

    try:
        mail.login(username, password)
        mail.select("Orders")
        status, messages = mail.uid('search', None, "ALL")

        if status == 'OK':
            email_uids = messages[0].split()
            logging.info(f"Number of messages found: {len(email_uids)}")

            for email_uid in email_uids:
                try:
                    if not check_imap_connection(mail):
                        logging.warning("IMAP connection lost. Attempting to reconnect...")
                        mail.logout()
                        mail = imaplib.IMAP4_SSL("imap.gmail.com")
                        mail.login(username, password)
                        mail.select("Orders")

                    logging.debug(f"Processing email UID: {email_uid.decode()}")
                    status, msg_data = mail.uid('fetch', email_uid, "(RFC822)")

                    if status == 'OK' and msg_data and msg_data[0] is not None:
                        if isinstance(msg_data[0], tuple):
                            msg = email.message_from_bytes(msg_data[0][1])
                            
                            # Try to get content from attachment first, then fall back to email body
                            body = get_email_content(msg)

                            if body:
                                email_dict = get_attachment_content(msg)
                                name, email_address = extract_email_and_name(email_dict)
                                logging.info(f"Extracted email: {email_address}, name: {name}")

                                # Extract the email sent date
                                date_tuple = email.utils.parsedate_tz(msg.get('Date'))
                                if date_tuple:
                                    email_sent_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime('%Y-%m-%d %H:%M:%S')
                                else:
                                    email_sent_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                                # Look up customer ID using both name and email
                                lookup_key = (name.lower() if name else None, email_address.lower())
                                customer_id = email_name_to_customer_ids.get(lookup_key)
                                
                                if customer_id:
                                    customer_codes = customer_product_codes.get(customer_id, {})
                                    process_result = process_email(email_address, body, customer_codes, customer_id, 
                                                                db_path, mail, email_uid, email_sent_date)
                                    logging.info(f"Email processed for customer {customer_id}: {process_result}")
                                    
                                    if move_email_to_folder(mail, email_uid, "EnteredIntoABS"):
                                        logging.info(f"Email from {email_address} moved to EnteredIntoABS folder")
                                    else:
                                        logging.error(f"Failed to move email from {email_address} to EnteredIntoABS folder")
                                else:
                                    logging.warning(f"No matching customer found for {name} <{email_address}>")
                                    if move_email_to_folder(mail, email_uid, "CustomerNotFound"):
                                        logging.info(f"Email moved to CustomerNotFound folder")
                                    else:
                                        logging.error(f"Failed to move email to CustomerNotFound folder")
                            else:
                                logging.warning(f"No body content for email UID: {email_uid.decode()}")
                                if move_email_to_folder(mail, email_uid, "CustomerNotFound"):
                                    logging.info(f"Email with no body content moved to CustomerNotFound folder")
                                else:
                                    logging.error(f"Failed to move email with no body content to CustomerNotFound folder")
                        else:
                            logging.warning(f"Unexpected data structure for email UID: {email_uid.decode()}")
                            if move_email_to_folder(mail, email_uid, "CustomerNotFound"):
                                logging.info(f"Email with unexpected data structure moved to CustomerNotFound folder")
                            else:
                                logging.error(f"Failed to move email with unexpected data structure to CustomerNotFound folder")
                    else:
                        logging.error(f"Failed to fetch email UID: {email_uid.decode()} or email data is None")
                        if move_email_to_folder(mail, email_uid, "CustomerNotFound"):
                            logging.info(f"Email that failed to fetch moved to CustomerNotFound folder")
                        else:
                            logging.error(f"Failed to move email that failed to fetch to CustomerNotFound folder")

                except Exception as e:
                    logging.error(f"Error processing email UID {email_uid.decode()}: {str(e)} {e}")
                    traceback.print_exc()  # This will print the stack trace to the console for debugging

                    if move_email_to_folder(mail, email_uid, "CustomerNotFound"):
                        logging.info(f"Email that caused an error moved to CustomerNotFound folder")
                    else:
                        logging.error(f"Failed to move email that caused an error to CustomerNotFound folder")
                    continue

        else:
            logging.error("Failed to search for emails.")
    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")
    finally:
        try:
            mail.close()
            mail.logout()
        except:
            pass

    # Create and run the GUI
    create_gui(db_path, email_name_to_customer_ids, product_enters_mapping)