import imaplib
import email
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

    #print("Extracting orders from email:")
    #print("Customer codes:", customer_codes)
    #print("Email content:")
    #print(email_text)

    def find_multi_line_match(lines, start_index, match_lines):
        if start_index + len(match_lines) > len(lines):
            return False
        return all(clean_line(lines[start_index + i]).startswith(clean_line(match_line))
                   for i, match_line in enumerate(match_lines))

    # Separate single-line and multi-line matches
    single_line_matches = {k: v for k, v in customer_codes.items() if '\n' not in k}
    multi_line_matches = {k: v for k, v in customer_codes.items() if '\n' in k}

    # Sort multi-line matches by number of lines (descending)
    # This is where we're extracting the line count from the Excel data
    sorted_multi_line_matches = sorted(multi_line_matches.items(), key=lambda x: len(x[0].split('\n')), reverse=True)

    i = 0
    while i < len(email_lines):
        matched = False

        # Check for multi-line matches first
        for match, (product_info, product_code, enters) in sorted_multi_line_matches:
            match_lines = match.split('\n')
            if find_multi_line_match(email_lines, i, match_lines):
                chunk = '\n'.join(email_lines[i:i+len(match_lines)])
                #print(f"Multi-line match found: {match} in {chunk}")
                quantity = extract_quantity(chunk)
                orders.append((quantity, enters, clean_line(chunk), product_code, quantity, 0))
                i += len(match_lines)
                matched = True
                break

        # If no multi-line match, check for single-line match
        if not matched:
            line = email_lines[i]
            cleaned_line = clean_line(line)
            for raw_text, (product_info, product_code, enters) in single_line_matches.items():
                if all(word.lower() in cleaned_line for word in raw_text.split()):
                    #print(f"Single-line match found: {raw_text} in {cleaned_line}")
                    quantity = extract_quantity(cleaned_line)
                    orders.append((quantity, enters, cleaned_line, product_code, quantity, 0))
                    matched = True
                    break

        # if not matched:
            #print(f"No match found for line: {cleaned_line}")

        i += 1  # Always increment i, whether matched or not

    #print(f"Extracted orders: {orders}")
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
    body = None
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if content_type == "text/plain" and "attachment" not in content_disposition:
                try:
                    body = part.get_payload(decode=True).decode()
                    break
                except:
                    continue
            elif "text/html" in content_type:
                try:
                    soup = BeautifulSoup(part.get_payload(decode=True).decode(), "html.parser")
                    body = soup.get_text()
                    break
                except:
                    continue
    else:
        content_type = msg.get_content_type()
        try:
            body = msg.get_payload(decode=True).decode()
            if "text/html" in content_type:
                soup = BeautifulSoup(body, "html.parser")
                body = soup.get_text()
        except:
            pass
    
    if body:
        body = body.replace('\r\n', '\n').strip()
    return body

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

def process_email(from_, body, customer_codes, customer_id, db_path, mail, msg_id, email_sent_date):
    email_address = extract_email_address(body)
    #print(f"Processing email from: {email_address}")
    #print(f"Customer ID: {customer_id}")
    #print(f"Customer codes: {customer_codes}")

    # Extract orders from the email body using the customer-specific product codes
    orders = extract_orders(body, customer_codes)
    #print(f"Orders extracted: {orders}")

    if orders:
        # Write orders to the SQL database
        write_orders_to_db(db_path, email_address, customer_id, orders, body, email_sent_date)

        # Move the email to the EnteredIntoABS inbox only if mail object is provided
        if mail and msg_id:
            move_email(mail, msg_id, "EnteredIntoABS")
            #print("Email processed successfully and moved to EnteredIntoABS")
        # else:
        #     #print("Email processed successfully (no email movement performed)")
        return True
    else:
        #print("No orders extracted from the email")
        return False

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

    def populate_grid(filter_date=None):
        for widget in scrollable_frame.winfo_children():
                widget.destroy()

        # Update headers
        for col, header in enumerate(headers):
            label = ttk.Label(scrollable_frame, text=header, font=("Arial", 10, "bold"))
            label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        if filter_date:
            next_date = filter_date + timedelta(days=1)
            cursor.execute('''
            SELECT raw_email, customer, item_id, quantity, enters, item, email_sent_date, date_to_be_entered, entered_status, customer_id
            FROM orders
            WHERE date_to_be_entered >= ? AND date_to_be_entered < ?
            ORDER BY entered_status ASC, date_to_be_entered ASC
            ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
        else:
            cursor.execute('''
            SELECT raw_email, customer, item_id, quantity, enters, item, email_sent_date, date_to_be_entered, entered_status, customer_id
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
        email_text_widget = email_frame.winfo_children()[0]  # The Text widget is the first child of the Frame
        raw_email = email_text_widget.get("1.0", tk.END).strip()
        
        email_sent_date = scrollable_frame.grid_slaves(row=row, column=6)[0]['text']

        # Create pop-up window
        popup = tk.Toplevel(root)
        popup.title("Add New Matching")
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

            # Update Excel file
            if update_customer_excel_file(customer_id, matching_phrase, matched_product):
                # Reload the order
                reload_order(customer_id, raw_email, email_sent_date)
                popup.destroy()
                messagebox.showinfo("Success", "New matching added and order reloaded.")
            else:
                messagebox.showerror("Error", "Failed to update Excel file. Please try again.")

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

        #print("Beginning Auto Order Entry in 5 seconds.")
        time.sleep(5)

        for row, var in checkbox_vars.items():
            if var.get():
                #print(f"Processing row {row}")
                raw_email_content = scrollable_frame.grid_slaves(row=row, column=0)[0].get("1.0", tk.END).strip()
                customer_id = scrollable_frame.grid_slaves(row=row, column=5)[0]['text']
                email_sent_date = scrollable_frame.grid_slaves(row=row, column=6)[0]['text']
                item_ids = scrollable_frame.grid_slaves(row=row, column=4)[0].get("1.0", tk.END).strip()
                quantities = scrollable_frame.grid_slaves(row=row, column=2)[0].get("1.0", tk.END).strip()
                enters = scrollable_frame.grid_slaves(row=row, column=3)[0].get("1.0", tk.END).strip()
                items = scrollable_frame.grid_slaves(row=row, column=1)[0].get("1.0", tk.END).strip()
                
                entered_status = check_entered_status(db_path, raw_email_content, customer_id, email_sent_date)
                if entered_status == 1:
                    already_entered_orders.append((raw_email_content, customer_id, email_sent_date, item_ids, quantities, enters, items))
                    #print(f"Order already entered: {customer_id}, {item_ids}")
                else:
                    selected_orders.append((raw_email_content, customer_id, email_sent_date, item_ids, quantities, enters, items))
                    #print(f"Order selected for processing: {customer_id}, {item_ids}")
        
        #print(f"Total orders selected for processing: {len(selected_orders)}")
        #print(f"Total orders already entered: {len(already_entered_orders)}")

        if selected_orders:
            #print("Selected orders:")
            # for order in selected_orders:
            #     #print(order)
            
            # Perform auto order entry
            perform_auto_entry(selected_orders)
            
            # Update the entered status in the database for selected rows
            update_entered_status(db_path, [(order[0], order[1], order[2]) for order in selected_orders])
            
            # Move processed emails to the ProcessedEmails inbox
            for order in selected_orders:
                move_email_by_content(order[0], order[1], order[2], "EnteredIntoABS", "ProcessedEmails")
            
            message = f"{len(selected_orders)} orders have been entered and marked as processed."
            if already_entered_orders:
                message += f"\n{len(already_entered_orders)} orders were already entered and skipped."
            
            messagebox.showinfo("Info", message)
            
            # Refresh the grid to reflect changes
            populate_grid(date_entry.get_date())
        elif already_entered_orders:
            messagebox.showwarning("Warning", f"All selected orders ({len(already_entered_orders)}) have already been entered.")
        else:
            #print("No rows selected.")
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

        #print("Beginning Auto Order Entry in 5 seconds.")
        time.sleep(5)

        for order in orders:
            raw_email, customer, customer_id, item_id, quantity, enters, item, entered_status, email_sent_date, date_to_be_entered = order
            order_key = (raw_email, customer_id, email_sent_date)
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

        unentered_orders_list = [
            (raw_email, customer_id, email_sent_date, '\n'.join(order_data[0]), '\n'.join(order_data[1]),
            '\n'.join(order_data[2]), '\n'.join(order_data[3]))
            for (raw_email, customer_id, email_sent_date), order_data in unentered_orders.items()
        ]

        #print(f"Total orders to be processed: {len(unentered_orders_list)}")
        #print(f"Total orders already entered: {len(already_entered_orders)}")

        if unentered_orders_list:
            #print("Orders to be processed:")
            # for order in unentered_orders_list:
            #     #print(order)
            
            # Perform auto order entry
            perform_auto_entry(unentered_orders_list)
            
            # Update the entered status in the database for processed orders
            update_entered_status(db_path, [(order[0], order[1], order[2]) for order in unentered_orders_list])
            
            # Move processed emails to the ProcessedEmails inbox
            for order in unentered_orders_list:
                move_email_by_content(order[0], order[1], order[2], "EnteredIntoABS", "ProcessedEmails")
            
            message = f"{len(unentered_orders_list)} orders have been entered and marked as processed."
            if already_entered_orders:
                message += f"\n{len(already_entered_orders)} orders were already entered and skipped."
            
            messagebox.showinfo("Info", message)
            
            # Refresh the grid to reflect changes
            populate_grid(date_entry.get_date())
        elif already_entered_orders:
            messagebox.showinfo("Info", f"All orders for the selected date ({len(already_entered_orders)}) have already been entered.")
        else:
            #print("No orders found for the selected date.")
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
        
        # Search for all emails in the folder
        status, messages = mail.search(None, "ALL")
        
        if status == 'OK' and messages[0]:
            for msg_id in messages[0].split():
                # Fetch the email content
                status, msg_data = mail.fetch(msg_id, "(RFC822)")
                if status == 'OK':
                    email_msg = email.message_from_bytes(msg_data[0][1])
                    email_body = get_email_content(email_msg)
                    email_date = email.utils.parsedate_to_datetime(email_msg['Date']).strftime('%Y-%m-%d %H:%M:%S')
                    
                    if (email_body.strip() == raw_email_content.strip() and
                        email_date == email_sent_date):
                        # Move the email
                        move_email(mail, msg_id, destination_folder)
                        #print(f"Moved email with ID {msg_id} to {destination_folder}")
                        break
        #     else:
        #         #print(f"Email not found in {source_folder} folder")
        # else:
        #     #print(f"No emails found in {source_folder} folder")
    
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
        WHERE raw_email = ? AND customer_id = ? AND email_sent_date = ?
        ''', (raw_email, customer_id, email_sent_date))

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

def extract_email_and_name(email_content):
    # Look for the specific forward format in metadata: "From: Name <email@domain.com>"
    forward_pattern = r'From: ([^<]+)<([^>]+)>'
    matches = re.search(forward_pattern, email_content)
    
    if matches:
        name = matches.group(1).strip()
        email = matches.group(2).strip()
        return name, email
    else:
        # If no match found, try to find just the email
        email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email_content)
        if email_match:
            return None, email_match.group(0)
        else:
            logging.warning("No email address found in the content")
            return None, None


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
        # Login to your account
        mail.login(username, password)

        # Ensure the CustomerNotFound folder exists
        if not ensure_imap_folder_exists(mail, 'CustomerNotFound'):
            logging.error("Unable to proceed due to folder creation issue.")
            raise Exception("Folder creation failed")
        else:
            logging.info("CustomerNotFound folder is ready")

        # Select the "Orders" mailbox
        mail.select("Orders")

        # Search for all emails in the "Orders" inbox
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
                            body = get_email_content(msg)

                            if body:
                                name, email_address = extract_email_and_name(body)
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
                    logging.error(f"Error processing email UID {email_uid.decode()}: {str(e)}")
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