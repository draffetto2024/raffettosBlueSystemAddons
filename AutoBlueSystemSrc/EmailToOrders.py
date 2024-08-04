import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
from email.utils import parseaddr
import re
import math

import Levenshtein
import pandas as pd
from nltk.tokenize import word_tokenize
from itertools import permutations, combinations
import sqlite3
from datetime import datetime, timedelta
import time
from tkcalendar import DateEntry

# Function to read product names, IDs, and lbs per case from an Excel file
def read_product_list_from_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    product_list = df.iloc[:, 0].str.lower().tolist()  # Assuming product names are in the first column
    product_ids = df.iloc[:, 1].tolist()  # Assuming product IDs are in the second column
    lbs_per_case = df.iloc[:, 2].fillna(0).tolist()  # Assuming lbs per case are in the third column, 0 if not specified
    return product_list, product_ids, lbs_per_case

# Function to read customer names and IDs from an Excel file
def read_customer_list_from_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    customer_list = df.iloc[:, 0].str.lower().tolist()  # Assuming customer names are in the first column
    customer_ids = df.iloc[:, 1].tolist()  # Assuming customer IDs are in the second column
    return customer_list, customer_ids

# Function to find the closest match using Levenshtein distance
def closest_match(phrase, match_list, match_ids, lbs_per_case, threshold=2):
    closest_match = None
    closest_id = None
    closest_lbs_per_case = None
    min_distance = float('inf')
    
    for item, item_id, item_lbs_per_case in zip(match_list, match_ids, lbs_per_case):
        if isinstance(phrase, str) and isinstance(item, str):
            distance = Levenshtein.distance(phrase, item)
            if distance <= threshold and distance < min_distance:
                min_distance = distance
                closest_match = item
                closest_id = item_id
                closest_lbs_per_case = item_lbs_per_case
                
    return closest_match, closest_id, closest_lbs_per_case

# Function to find the closest match using Levenshtein distance
def closest_customer_match(phrase, match_list, match_ids, threshold=2):
    closest_match = None
    closest_id = None
    min_distance = float('inf')
    
    for item, item_id in zip(match_list, match_ids):
        if isinstance(phrase, str) and isinstance(item, str):
            distance = Levenshtein.distance(phrase, item)
            if distance <= threshold and distance < min_distance:
                min_distance = distance
                closest_match = item
                closest_id = item_id
                
    return closest_match, closest_id

# Function to generate all permutations of words in a list
def generate_permutations(words):
    all_permutations = []
    for i in range(len(words), 0, -1):  # Start with the longest set and go to smaller sets
        for comb in combinations(words, i):
            all_permutations.extend(permutations(comb))
    return all_permutations

# Function to clean a line by removing non-letter characters except spaces and numbers
def clean_line(line):
    return re.sub(r'[^a-zA-Z0-9 ]', '', line).lower()

# Function to process each line and extract order details
def process_line(line, product_list, product_ids, lbs_per_case):
    cleaned_line = clean_line(line)
    words = word_tokenize(cleaned_line)
    count = None
    count_type = 'cases'  # Default count type
    product = None
    product_id = None
    product_lbs_per_case = 0
    
    for i, word in enumerate(words):
        if word.isdigit():
            count = int(word)
            if i + 1 < len(words) and words[i + 1].lower() in ['lbs', 'lb']:
                count_type = 'lbs'
            potential_product_words = words[i + 1:]
            all_permutations = generate_permutations(potential_product_words)
            for permutation in all_permutations:
                permutation_phrase = " ".join(permutation)
                product, product_id, product_lbs_per_case = closest_match(permutation_phrase, product_list, product_ids, lbs_per_case)
                if product:
                    if count_type == 'lbs' and product_lbs_per_case > 0:   
                        count = math.ceil(count / product_lbs_per_case)
                        count_type = 'cases'
                    return count, count_type, product, product_id  # Exit as soon as a match is found
            break
    
    if count is None:
        # Assume 1 case if no count is provided
        count = 1
        all_permutations = generate_permutations(words)
        for permutation in all_permutations:
            permutation_phrase = " ".join(permutation)
            product, product_id, product_lbs_per_case = closest_match(permutation_phrase, product_list, product_ids, lbs_per_case)
            if product:
                return count, count_type, product, product_id  # Exit as soon as a match is found
    
    return count, count_type, product, product_id

# Main function to extract orders from email text
def extract_orders(email_text, product_list, product_ids, lbs_per_case):
    orders = []
    lines = email_text.strip().split('\n')
    
    for line in lines:
        if line.strip():  # Skip empty lines
            count, count_type, product, product_id = process_line(line, product_list, product_ids, lbs_per_case)
            if product:
                orders.append((count, count_type, product, product_id))
    
    return orders

# Modify the write_orders_to_db function to include email_sent_date
def write_orders_to_db(db_path, customer, customer_id, orders, raw_email, email_sent_date):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Create table if it doesn't exist (add email_sent_date column)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer TEXT,
        customer_id INTEGER,
        cases INTEGER,
        lbs FLOAT,
        item TEXT,
        item_id INTEGER,
        date_generated TEXT,
        date_processed TEXT,
        entered_status INTEGER DEFAULT 0,
        raw_email TEXT,
        email_sent_date TEXT
    )
    ''')
    
    date_generated = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_processed = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    for count, count_type, product, product_id in orders:
        cases = count if count_type == 'cases' else 0
        lbs = count if count_type == 'lbs' else 0.0
        cursor.execute('''
        INSERT INTO orders (customer, customer_id, cases, lbs, item, item_id, date_generated, date_processed, entered_status, raw_email, email_sent_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (customer, customer_id, cases, lbs, product, product_id, date_generated, date_processed, 0, raw_email, email_sent_date))
    
    conn.commit()
    conn.close()

# Function to update the entered status of orders in the database
def update_entered_status(db_path, raw_emails):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    for raw_email in raw_emails:
        # Update the entered_status to 1 for matching orders
        cursor.execute('''
        UPDATE orders
        SET entered_status = 1
        WHERE raw_email = ?
        ''', (raw_email,))

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
    
    # Clean up the body text
    if body:
        body = body.replace('\r\n', '\n').strip()
    return body

# Function to extract email address from the "From" field
def extract_email_address(from_field):
    name, email_address = parseaddr(from_field)
    return email_address

# Modify the process_email function to extract and pass the email sent date
def process_email(from_, body, product_list, product_ids, lbs_per_case, customer_list, customer_ids, db_path, mail, msg_id, email_sent_date):
    email_address = extract_email_address(from_)
    
    # Find the closest matching customer
    customer, customer_id = closest_customer_match(email_address, customer_list, customer_ids)
    
    if customer:
        # Extract orders from the email body
        orders = extract_orders(body, product_list, product_ids, lbs_per_case)
        print(f"Orders extracted: {orders}")
        
        # Write orders to the SQL database (now including email_sent_date)
        write_orders_to_db(db_path, customer, customer_id, orders, body, email_sent_date)
        
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

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

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
    headers = ["Raw Email Content", "Matched Products", "Cases", "Lbs", "Product Codes", "Customer", "Email Sent Date", "Entered Status", "Select"]

    # Function to populate the grid with data from the database
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
                SELECT raw_email, customer, item, cases, lbs, item_id, email_sent_date, entered_status
                FROM orders
                WHERE email_sent_date >= ? AND email_sent_date < ?
                ORDER BY entered_status ASC, email_sent_date ASC
                ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
            else:
                cursor.execute('''
                SELECT raw_email, customer, item, cases, lbs, item_id, email_sent_date, entered_status
                FROM orders
                WHERE email_sent_date IS NOT NULL
                ORDER BY entered_status ASC, email_sent_date ASC
                ''')
        else:
            cursor.execute('''
            SELECT raw_email, customer, item, cases, lbs, item_id, entered_status
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
                _, customer, item, cases, lbs, item_id = order[:6]  # First 6 items are always the same
                matched_products.append(item)
                cases_list.append(str(cases) if cases else "N/A")
                lbs_list.append(str(lbs) if lbs else "N/A")
                product_codes.append(str(item_id))
                customers.add(customer)
                if len(order) > 7:  # If email_sent_date exists
                    email_sent_date = order[6]

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

            # Customer
            customer_label = ttk.Label(scrollable_frame, text=", ".join(customers))
            customer_label.grid(row=row, column=5, padx=5, pady=5, sticky="nsew")

            # Email Sent Date
            email_sent_date_label = ttk.Label(scrollable_frame, text=email_sent_date)
            email_sent_date_label.grid(row=row, column=6, padx=5, pady=5, sticky="nsew")

            # Entered Status
            entered_status_label = ttk.Label(scrollable_frame, text="Entered" if entered_status else "Not Entered")
            entered_status_label.grid(row=row, column=7, padx=5, pady=5, sticky="nsew")

            # Checkbox
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(scrollable_frame, variable=var)
            checkbox.grid(row=row, column=8, padx=5, pady=5, sticky="nsew")
            checkbox_vars[row] = var

    # Function to print selected rows and update the status
    def auto_enter_selected():
        selected_raw_emails = []
        for row, var in checkbox_vars.items():
            if var.get():
                # Get the raw email content for the selected row
                raw_email_content = scrollable_frame.grid_slaves(row=row, column=0)[0].get("1.0", tk.END).strip()
                selected_raw_emails.append(raw_email_content)
        
        if selected_raw_emails:
            print("Selected rows:")
            for raw_email in selected_raw_emails:
                print(raw_email)
            # Update the entered status in the database for selected rows
            update_entered_status(db_path, selected_raw_emails)
            messagebox.showinfo("Info", "Selected orders have been marked as entered.")
            # Move processed emails to the ProcessedEmails inbox
            for raw_email in selected_raw_emails:
                move_email_by_content(raw_email, "ProcessedEmails")
            # Refresh the grid to reflect changes
            populate_grid(date_entry.get_date())
        else:
            print("No rows selected.")
            messagebox.showwarning("Warning", "No rows selected.")

    # Function to auto-enter all unentered emails for the selected date
    def auto_enter_all():
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        filter_date = date_entry.get_date()
        next_date = filter_date + timedelta(days=1)
        
        cursor.execute('''
        SELECT raw_email FROM orders 
        WHERE entered_status = 0 AND email_sent_date >= ? AND email_sent_date < ?
        ''', (filter_date.strftime('%Y-%m-%d'), next_date.strftime('%Y-%m-%d')))
        
        unentered_emails = [row[0] for row in cursor.fetchall()]
        
        if unentered_emails:
            update_entered_status(db_path, unentered_emails)
            for raw_email in unentered_emails:
                move_email_by_content(raw_email, "ProcessedEmails")
            messagebox.showinfo("Info", f"{len(unentered_emails)} orders have been marked as entered.")
            populate_grid(filter_date)
        else:
            messagebox.showinfo("Info", "No unentered orders found for the selected date.")
        
        conn.close()

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
# Update these functions to handle the case when there are no more emails to process
def update_entered_status(db_path, raw_emails):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    for raw_email in raw_emails:
        # Update the entered_status to 1 for matching orders
        cursor.execute('''
        UPDATE orders
        SET entered_status = 1
        WHERE raw_email = ?
        ''', (raw_email,))

    conn.commit()
    conn.close()

def move_email_by_content(raw_email_content, destination_folder):
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
                    move_email(mail, msg_id, destination_folder)
    mail.close()
    mail.logout()

if __name__ == "__main__":
    # Read product list and IDs from Excel
    product_list, product_ids, lbs_per_case = read_product_list_from_excel('products.xlsx')

    # Read customer list and IDs from Excel
    customer_list, customer_ids = read_customer_list_from_excel('customers.xlsx')

    # Email details
    username = "gingoso2@gmail.com"
    password = "soiz avjw bdtu hmtn"  # utilizing an app password

    # Database path
    db_path = 'orders.db'

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

                                process_result = process_email(from_, body, product_list, product_ids, lbs_per_case, customer_list, customer_ids, db_path, mail, email_uid, email_sent_date)
                                print(f"Email processed: {process_result}")

                                # Move the processed email to "EnteredintoABS" folder
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