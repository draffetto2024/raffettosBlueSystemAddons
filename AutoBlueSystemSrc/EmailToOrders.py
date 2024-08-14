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
import pyautogui

# Function to read customer product grid from an Excel file
def read_grid_excel(file_path):
    df = pd.read_excel(file_path, header=0, index_col=0)
    
    # Initialize a dictionary to store customer-specific product codes
    customer_product_codes = {}
    
    for customer_email, row in df.iterrows():
        customer_codes = {}
        for product_type, value in row.items():
            if pd.notna(value):  # Check if the value is not NaN
                amount, product_code = value.split(maxsplit=1)
                customer_codes[product_type.lower()] = (int(amount), product_code)
        customer_product_codes[customer_email.lower()] = customer_codes
    
    return customer_product_codes

# Function to clean a line by removing non-letter characters except spaces and numbers
def clean_line(line):
    return re.sub(r'[^a-zA-Z0-9 ]', '', line).lower()

# Function to extract orders based on customer-specific product codes
def extract_orders(email_text, customer_codes):
    orders = []
    lines = email_text.strip().split('\n')
    
    for line in lines:
        if line.strip():  # Skip empty lines
            cleaned_line = clean_line(line)

            print(cleaned_line)
            print(customer_codes)
            
            if cleaned_line in customer_codes:
                matched_amount, product_code = customer_codes[cleaned_line]
                orders.append((matched_amount, 'cases', cleaned_line, product_code, 0))  # Assuming no lbs information is needed
    
    return orders

def write_orders_to_db(db_path, customer_email, customer_id, orders, raw_email, email_sent_date):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    date_generated = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_processed = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    for count, count_type, product, product_id, lbs in orders:
        cases = count if count_type == 'cases' else 0
        lbs = count if count_type == 'lbs' else lbs
        cursor.execute('''
        INSERT INTO orders (customer, customer_id, cases, lbs, item, item_id, date_generated, date_processed, entered_status, raw_email, email_sent_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (customer_email, customer_id, cases, lbs, product, product_id, date_generated, date_processed, 0, raw_email, email_sent_date))
    
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
def extract_email_address(from_field):
    name, email_address = parseaddr(from_field)
    return email_address

def process_email(from_, body, customer_product_codes, customer_ids, db_path, mail, msg_id, email_sent_date):
    email_address = extract_email_address(from_)
    
    if email_address in customer_product_codes:
        customer_codes = customer_product_codes[email_address]
        customer_id = customer_ids.get(email_address, email_address)  # Use email as fallback if no ID found
        
        # Extract orders from the email body using the customer-specific product codes
        orders = extract_orders(body, customer_codes)
        print(f"Orders extracted: {orders}")
        
        # Write orders to the SQL database
        write_orders_to_db(db_path, email_address, customer_id, orders, body, email_sent_date)
        
        # Move the email to the EnteredIntoABS inbox
        move_email(mail, msg_id, "EnteredIntoABS")
        
        return True  # Customer found and processed
    else:
        print(f"No matching customer found for email from: {email_address}")
        
        return False  # Customer not found

# Function to move an email to another folder
def move_email(mail, msg_id, destination_folder):
    result = mail.copy(msg_id, destination_folder)
    if result[0] == 'OK':
        mail.store(msg_id, '+FLAGS', '\\Deleted')
        mail.expunge()
    else:
        print(f"Failed to move email ID {msg_id} to {destination_folder}")

def initialize_database(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Create table if it doesn't exist
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer TEXT,
        customer_id TEXT,
        cases INTEGER,
        lbs FLOAT,
        item TEXT,
        item_id TEXT,
        date_generated TEXT,
        date_processed TEXT,
        entered_status INTEGER DEFAULT 0,
        raw_email TEXT,
        email_sent_date TEXT
    )
    ''')
    
    conn.commit()
    conn.close()

def create_gui(db_path):
    root = tk.Tk()
    root.title("Email Order Processor")
    root.geometry("1800x700")  # Increased height to accommodate new controls

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

    # Update headers to include Email Sent Date
    headers = ["Raw Email Content", "Matched Products", "Cases", "Lbs", "Product Codes", "Customer", "Email Sent Date", "Entered Status", "Select", "Delete"]

    def populate_grid(filter_date=None):
        for widget in scrollable_frame.winfo_children():
            widget.destroy()

        # Recreate headers
        for col, header in enumerate(headers):
            label = ttk.Label(scrollable_frame, text=header, font=("Arial", 10, "bold"))
            label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Check if email_sent_date column exists
        cursor.execute("PRAGMA table_info(orders)")
        columns = [column[1] for column in cursor.fetchall()]

        if 'email_sent_date' in columns:
            if filter_date:
                next_date = filter_date + timedelta(days=1)
                cursor.execute('''
                SELECT raw_email, customer, item_id, cases, lbs, item_id, email_sent_date, entered_status
                FROM orders
                WHERE email_sent_date >= ? AND email_sent_date < ?
                ORDER BY entered_status ASC, email_sent_date ASC
                ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
            else:
                cursor.execute('''
                SELECT raw_email, customer, item_id, cases, lbs, item_id, email_sent_date, entered_status
                FROM orders
                WHERE email_sent_date IS NOT NULL
                ORDER BY entered_status ASC, email_sent_date ASC
                ''')
        else:
            cursor.execute('''
            SELECT raw_email, customer, item_id, cases, lbs, item_id, entered_status
            FROM orders
            ORDER BY entered_status ASC
            ''')

        orders = cursor.fetchall()
        conn.close()

        email_to_orders = {}
        for order in orders:
            raw_email = order[0]
            if raw_email not in email_to_orders:
                email_to_orders[raw_email] = []
            email_to_orders[raw_email].append(order)

        for row, (raw_email, order_list) in enumerate(email_to_orders.items(), start=1):
            # Raw Email Content
            email_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=30, height=5)
            email_text.insert(tk.END, raw_email)
            email_text.config(state=tk.DISABLED)
            email_text.grid(row=row, column=0, padx=5, pady=5, sticky="nsew")

            matched_products = []
            cases_list = []
            lbs_list = []
            product_codes = []
            customers = set()
            email_sent_date = "N/A"
            entered_status = order_list[0][-1]  # Last item is always entered_status

            for order in order_list:
                _, customer, item_id, cases, lbs, _, email_sent_date = order[:7]  # First 7 items are always the same
                matched_products.append(item_id)
                cases_list.append(str(cases) if cases else "N/A")
                lbs_list.append(str(lbs) if lbs else "N/A")
                product_codes.append(item_id)
                customers.add(customer)

            # Matched Products
            products_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=30, height=5)
            products_text.insert(tk.END, "\n".join(matched_products))
            products_text.config(state=tk.DISABLED)
            products_text.grid(row=row, column=1, padx=5, pady=5, sticky="nsew")

            # Cases
            cases_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=10, height=5)
            cases_text.insert(tk.END, "\n".join(cases_list))
            cases_text.config(state=tk.DISABLED)
            cases_text.grid(row=row, column=2, padx=5, pady=5, sticky="nsew")

            # Lbs
            lbs_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=10, height=5)
            lbs_text.insert(tk.END, "\n".join(lbs_list))
            lbs_text.config(state=tk.DISABLED)
            lbs_text.grid(row=row, column=3, padx=5, pady=5, sticky="nsew")

            # Product Codes
            codes_text = tk.Text(scrollable_frame, wrap=tk.WORD, width=20, height=5)
            codes_text.insert(tk.END, "\n".join(product_codes))
            codes_text.config(state=tk.DISABLED)
            codes_text.grid(row=row, column=4, padx=5, pady=5, sticky="nsew")

            # In the populate_grid function, replace the customer label creation with:
            customer_id = customer_ids.get(customer, customer)  # Use email as fallback if no ID found
            customer_label = ttk.Label(scrollable_frame, text=customer_id)
            customer_label.grid(row=row, column=5, padx=5, pady=5, sticky="nsew")

            # Email Sent Date
            email_sent_date_label = ttk.Label(scrollable_frame, text=email_sent_date)
            email_sent_date_label.grid(row=row, column=6, padx=5, pady=5, sticky="nsew")

            # Entered Status
            entered_status_label = ttk.Label(scrollable_frame, text="Entered" if entered_status else "Not Entered")
            entered_status_label.grid(row=row, column=7, padx=5, pady=5, sticky="nsew")

            # Add Delete button
            delete_button = ttk.Button(scrollable_frame, text="Delete", command=lambda r=raw_email: delete_order(r))
            delete_button.grid(row=row, column=9, padx=5, pady=5, sticky="nsew")

            # Checkbox
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(scrollable_frame, variable=var)
            checkbox.grid(row=row, column=8, padx=5, pady=5, sticky="nsew")
            checkbox_vars[row] = var

    def perform_auto_entry(orders):
        print(f"Performing auto entry for {len(orders)} orders")
        # Sequence to enter the order entry screen
        into_order_sequence = [
            'KEY:enter',
            'KEY:1',
            'KEY:enter',
        ]
        # Execute the sequence to enter the order entry screen
        auto_order_entry(into_order_sequence)

        for i, order in enumerate(orders, 1):
            print(f"Processing order {i} of {len(orders)}: {order}")

            # Extract customer ID and other details from the order
            raw_email, customer, customer_id, item_ids, cases, lbs, items = order
            
            pre_order_entry = [
                'KEY:enter',
                f'INPUT:{customer_id}',
                'KEY:enter',
                'KEY:enter',
            ]
            # Execute the pre-order entry sequence
            auto_order_entry(pre_order_entry)

            # Generate the sequence for this specific order
            order_sequence = generate_order_sequence((customer_id, item_ids, cases, lbs, items))
            
            # Execute the sequence for this order
            auto_order_entry(order_sequence)

        print("Auto entry completed")

    def auto_enter_selected():
        selected_orders = []
        already_entered_orders = []

        print("Beginning Auto Order Entry in 5 seconds.")
        time.sleep(5)

        for row, var in checkbox_vars.items():
            if var.get():
                print(f"Processing row {row}")
                raw_email_content = scrollable_frame.grid_slaves(row=row, column=0)[0].get("1.0", tk.END).strip()
                customer = scrollable_frame.grid_slaves(row=row, column=5)[0]['text']
                item_ids = scrollable_frame.grid_slaves(row=row, column=4)[0].get("1.0", tk.END).strip()
                cases = scrollable_frame.grid_slaves(row=row, column=2)[0].get("1.0", tk.END).strip()
                lbs = scrollable_frame.grid_slaves(row=row, column=3)[0].get("1.0", tk.END).strip()
                items = scrollable_frame.grid_slaves(row=row, column=1)[0].get("1.0", tk.END).strip()
                
                entered_status = check_entered_status(db_path, raw_email_content)
                if entered_status == 1:
                    already_entered_orders.append((raw_email_content, customer, item_ids, cases, lbs, items))
                    print(f"Order already entered: {customer}, {item_ids}")
                else:
                    selected_orders.append((raw_email_content, customer, item_ids, cases, lbs, items))
                    print(f"Order selected for processing: {customer}, {item_ids}")
        
        print(f"Total orders selected for processing: {len(selected_orders)}")
        print(f"Total orders already entered: {len(already_entered_orders)}")

        if selected_orders:
            print("Selected orders:")
            for order in selected_orders:
                print(order)
            
            # Perform auto order entry
            perform_auto_entry(selected_orders)
            
            # Update the entered status in the database for selected rows
            update_entered_status(db_path, [order[0] for order in selected_orders])
            
            # Move processed emails to the ProcessedEmails inbox
            for order in selected_orders:
                move_email_by_content(order[0], "EnteredIntoABS", "ProcessedEmails")
            
            message = f"{len(selected_orders)} orders have been entered and marked as processed."
            if already_entered_orders:
                message += f"\n{len(already_entered_orders)} orders were already entered and skipped."
            
            messagebox.showinfo("Info", message)
            
            # Refresh the grid to reflect changes
            populate_grid(date_entry.get_date())
        elif already_entered_orders:
            messagebox.showwarning("Warning", f"All selected orders ({len(already_entered_orders)}) have already been entered.")
        else:
            print("No rows selected.")
            messagebox.showwarning("Warning", "No rows selected.")

    def auto_enter_all():
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        filter_date = date_entry.get_date()
        next_date = filter_date + timedelta(days=1)
        
        cursor.execute('''
        SELECT raw_email, customer, customer_id, item_id, cases, lbs, item, entered_status
        FROM orders 
        WHERE email_sent_date >= ? AND email_sent_date < ?
        ORDER BY entered_status ASC, email_sent_date ASC
        ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
        
        orders = cursor.fetchall()
        conn.close()

        unentered_orders = {}
        already_entered_orders = {}

        print("Beginning Auto Order Entry in 5 seconds.")
        time.sleep(5)

        for order in orders:
            raw_email, customer, customer_id, item_id, cases, lbs, item, entered_status = order
            order_key = (raw_email, customer)  # Use both raw_email and customer as the key
            if entered_status == 0:
                if order_key not in unentered_orders:
                    unentered_orders[order_key] = [customer_id, [], [], [], []]
                unentered_orders[order_key][1].append(item_id)
                unentered_orders[order_key][2].append(str(cases))
                unentered_orders[order_key][3].append(str(lbs))
                unentered_orders[order_key][4].append(item)
            else:
                if order_key not in already_entered_orders:
                    already_entered_orders[order_key] = [customer_id, [], [], [], []]
                already_entered_orders[order_key][1].append(item_id)
                already_entered_orders[order_key][2].append(str(cases))
                already_entered_orders[order_key][3].append(str(lbs))
                already_entered_orders[order_key][4].append(item)

        unentered_orders_list = [(raw_email, customer, *order_data) for (raw_email, customer), order_data in unentered_orders.items()]

        print(f"Total orders to be processed: {len(unentered_orders_list)}")
        print(f"Total orders already entered: {len(already_entered_orders)}")

        if unentered_orders_list:
            print("Orders to be processed:")
            for order in unentered_orders_list:
                print(order)
            
            # Perform auto order entry
            perform_auto_entry(unentered_orders_list)
            
            # Update the entered status in the database for processed orders
            update_entered_status(db_path, [order[0] for order in unentered_orders_list])
            
            # Move processed emails to the ProcessedEmails inbox
            for order in unentered_orders_list:
                move_email_by_content(order[0], "EnteredIntoABS", "ProcessedEmails")
            
            message = f"{len(unentered_orders_list)} orders have been entered and marked as processed."
            if already_entered_orders:
                message += f"\n{len(already_entered_orders)} orders were already entered and skipped."
            
            messagebox.showinfo("Info", message)
            
            # Refresh the grid to reflect changes
            populate_grid(date_entry.get_date())
        elif already_entered_orders:
            messagebox.showinfo("Info", f"All orders for the selected date ({len(already_entered_orders)}) have already been entered.")
        else:
            print("No orders found for the selected date.")
            messagebox.showwarning("Warning", "No orders found for the selected date.")


    def delete_order(raw_email):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this order?"):
            delete_order_from_db(db_path, raw_email)
            move_email_back_to_orders(raw_email)
            messagebox.showinfo("Success", "Order deleted and email moved back to Orders inbox.")
            populate_grid(date_entry.get_date())  # Refresh the grid

    def process_orders(orders_to_process):
        if orders_to_process:
            print("Processing orders:")
            for order in orders_to_process:
                print(order)
            
            # Perform auto order entry
            perform_auto_entry(orders_to_process)
            
            # Update the entered status in the database for processed orders
            update_entered_status(db_path, [order[0] for order in orders_to_process])
            
            # Move processed emails to the ProcessedEmails inbox
            for order in orders_to_process:
                move_email_by_content(order[0], "EnteredIntoABS", "ProcessedEmails")
            
            message = f"{len(orders_to_process)} orders have been entered and marked as processed."
            
            messagebox.showinfo("Info", message)
            
            # Refresh the grid to reflect changes
            populate_grid(date_entry.get_date())
        else:
            print("No orders to process.")
            messagebox.showwarning("Warning", "No orders to process.")

    # Add buttons to auto-enter selected rows and all unentered emails
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(pady=10)

    auto_enter_button = ttk.Button(button_frame, text="Auto Enter Selected Rows", command=auto_enter_selected)
    auto_enter_button.pack(side=tk.LEFT, padx=5)

    auto_enter_all_button = ttk.Button(button_frame, text="Auto Enter All Unentered", command=auto_enter_all)
    auto_enter_all_button.pack(side=tk.LEFT, padx=5)

    # Populate the grid with today's date
    populate_grid(datetime.now().date())

    root.mainloop()



# def generate_keystroke_sequence(order):
#     # Assuming order is a tuple with (raw_email_content, customer, item_id, cases, lbs, item)
#     _, _, product_code, cases, lbs, item = order
    
#     sequence = [
#         f'INPUT:{product_code}',
#         'KEY:enter',
#         f'INPUT:{cases}',
#         'KEY:enter',
#         'KEY:enter',
#         'KEY:enter',
#         'KEY:enter',
#         'WAIT:0.1'
#     ]
#     return sequence

def move_email_by_content(raw_email_content, source_folder, destination_folder):
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
                    email_body = get_email_content(email.message_from_bytes(msg_data[0][1]))
                    if email_body.strip() == raw_email_content.strip():
                        # Move the email
                        move_email(mail, msg_id, destination_folder)
                        print(f"Moved email with ID {msg_id} to {destination_folder}")
                        break
            else:
                print(f"Email not found in {source_folder} folder")
        else:
            print(f"No emails found in {source_folder} folder")
    
    finally:
        mail.close()
        mail.logout()

def check_entered_status(db_path, raw_email_content):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
    SELECT entered_status
    FROM orders
    WHERE raw_email = ?
    LIMIT 1
    ''', (raw_email_content,))
    
    result = cursor.fetchone()
    conn.close()
    
    return result[0] if result else None

def read_customer_ids(file_path):
    df = pd.read_excel(file_path, header=None, names=['email', 'customer_id'])
    return dict(zip(df['email'].str.lower(), df['customer_id']))

def generate_order_sequence(order):
    # Assuming order is now a tuple with (customer_id, item_ids, cases, lbs, items)
    customer_id, item_ids, cases, lbs, items = order
    
    sequence = []
    
    for product_code, case, lb, item in zip(item_ids, cases, lbs, items):
        sequence.extend([
            f'INPUT:{product_code.strip()}',
            'KEY:enter',
            f'INPUT:{case.strip()}',
            'KEY:enter',
            'KEY:enter',
            'KEY:enter',
            'KEY:enter',
            'WAIT:0.1'
        ])
    return sequence

def auto_order_entry(keystroke_sequence):

    for action in keystroke_sequence:
        if action.startswith('INPUT:'):
            # Input data
            data = action.split(':', 1)[1]  # Use maxsplit=1 to handle colons in the data
            pyautogui.typewrite(str(data))
        elif action.startswith('KEY:'):
            # Press a specific key
            key = action.split(':')[1]
            pyautogui.press(key)
        elif action.startswith('WAIT:'):
            # Wait for a specified number of seconds
            wait_time = float(action.split(':')[1])
            time.sleep(wait_time)
        else:
            # Assume it's a keystroke or text to type
            pyautogui.typewrite(action)
        
        time.sleep(0.5)  # Short pause between action

def delete_order_from_db(db_path, raw_email):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute('''
    DELETE FROM orders
    WHERE raw_email = ?
    ''', (raw_email,))
    
    conn.commit()
    conn.close()

# Function to update the entered status of orders in the database
def update_entered_status(db_path, raw_emails):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    for raw_email in raw_emails:
        cursor.execute('''
        UPDATE orders
        SET entered_status = 1
        WHERE raw_email = ?
        ''', (raw_email,))

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

if __name__ == "__main__":
    # Read customer-specific product codes from the grid Excel sheet
    customer_product_codes = read_grid_excel('customer_product_grid.xlsx')

    # Read customer IDs from Excel file
    customer_ids = read_customer_ids('customers.xlsx')  # Replace with your actual file name

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

        # Select the "Orders" mailbox
        mail.select("Orders")

        # Search for all emails in the "Orders" inbox
        status, messages = mail.uid('search', None, "ALL")

        if status == 'OK':
            email_uids = messages[0].split()
            print(f"Number of messages found: {len(email_uids)}")

            for email_uid in email_uids:
                try:
                    print(f"\nProcessing email UID: {email_uid.decode()}")
                    status, msg_data = mail.uid('fetch', email_uid, "(RFC822)")

                    if status == 'OK' and msg_data and msg_data[0] is not None:
                        if isinstance(msg_data[0], tuple):
                            msg = email.message_from_bytes(msg_data[0][1])
                            
                            body = get_email_content(msg)

                            if body:
                                from_ = msg.get("From")
                                email_address = extract_email_address(from_)

                                # Extract the email sent date
                                date_tuple = email.utils.parsedate_tz(msg.get('Date'))
                                if date_tuple:
                                    email_sent_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime('%Y-%m-%d %H:%M:%S')
                                else:
                                    email_sent_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Use current time if date parsing fails

                                process_result = process_email(from_, body, customer_product_codes, customer_ids, db_path, mail, email_uid, email_sent_date)
                                print(f"Email processed: {process_result}")

                                # Move the processed email to "EnteredIntoABS" folder
                                mail.uid('copy', email_uid, "EnteredIntoABS")
                                mail.uid('store', email_uid, '+FLAGS', '\\Deleted')
                                mail.expunge()

                            else:
                                print(f"No body content for email UID: {email_uid.decode()}")
                        else:
                            print(f"Unexpected data structure for email UID: {email_uid.decode()}")
                    else:
                        print(f"Failed to fetch email UID: {email_uid.decode()} or email data is None")

                    # Short pause to allow server to process the move
                    time.sleep(1)

                except Exception as e:
                    print(f"Error processing email UID {email_uid.decode()}: {str(e)}")
                    continue

        else:
            print("Failed to search for emails.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        try:
            mail.close()
            mail.logout()
        except:
            pass

    # Create and run the GUI
    create_gui(db_path)
