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
from datetime import datetime

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

# Function to write orders to a SQL database
def write_orders_to_db(db_path, customer, customer_id, orders):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Create table if it doesn't exist
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
        date_processed TEXT
    )
    ''')
    
    date_generated = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_processed = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    for count, count_type, product, product_id in orders:
        cases = count if count_type == 'cases' else 0
        lbs = count if count_type == 'lbs' else 0.0
        cursor.execute('''
        INSERT INTO orders (customer, customer_id, cases, lbs, item, item_id, date_generated, date_processed)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (customer, customer_id, cases, lbs, product, product_id, date_generated, date_processed))
    
    conn.commit()
    conn.close()

# Read product list and IDs from Excel
product_list, product_ids, lbs_per_case = read_product_list_from_excel('products.xlsx')

# Read customer list and IDs from Excel
customer_list, customer_ids = read_customer_list_from_excel('customers.xlsx')

# Email details
username = "gingoso2@gmail.com"
password = "soiz avjw bdtu hmtn"  # utilizing an app password

# Connect to the server
mail = imaplib.IMAP4_SSL("imap.gmail.com")

# Login to your account
mail.login(username, password)

# Select the "Orders" mailbox
mail.select("Orders")

# Search for all emails in the "Orders" inbox
status, messages = mail.search(None, "ALL")

# Convert messages to a list of email IDs
messages = messages[0].split()

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

# Function to process the email and perform a task
def process_email(from_, body, product_list, product_ids, lbs_per_case, customer_list, customer_ids, db_path):
    email_address = extract_email_address(from_)
    
    # Find the closest matching customer
    customer, customer_id = closest_customer_match(email_address, customer_list, customer_ids)
    
    if customer:
        # Extract orders from the email body
        orders = extract_orders(body, product_list, product_ids, lbs_per_case)
        print(f"Orders extracted: {orders}")
        
        # Write orders to the SQL database
        write_orders_to_db(db_path, customer, customer_id, orders)
        
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

# Process each email in the "Orders" inbox
for msg_id in messages:
    # Fetch the email by ID
    status, msg_data = mail.fetch(msg_id, "(RFC822)")
    
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()

            print("Subject:", subject)

            # Email sender
            from_ = msg.get("From")
            print("From:", from_)

            # Extract email address from the "From" field
            email_address = extract_email_address(from_)
            print("Email Address:", email_address)

            # Email content
            body = get_email_content(msg)
            print("Body:", body)
            print("-" * 50)

            if body:
                # Process the email
                customer_found = process_email(email_address, body, product_list, product_ids, lbs_per_case, customer_list, customer_ids, 'orders.db')

                # Move the email to the appropriate folder based on whether the customer was found
                # if customer_found:
                #     move_email(mail, msg_id, "ProcessedEmails")
                # else:
                #     move_email(mail, msg_id, "CustomerNotFound")

# Close the connection and logout
mail.close()
mail.logout()
