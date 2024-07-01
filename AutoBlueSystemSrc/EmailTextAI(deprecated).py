import re
import Levenshtein
import pandas as pd
from nltk.tokenize import word_tokenize
from itertools import permutations, combinations
import sqlite3
from datetime import datetime

# Function to read product names and IDs from an Excel file
def read_product_list_from_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    product_list = df.iloc[:, 0].str.lower().tolist()  # Assuming product names are in the first column
    product_ids = df.iloc[:, 1].tolist()  # Assuming product IDs are in the second column
    return product_list, product_ids

# Function to read customer names and IDs from an Excel file
def read_customer_list_from_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    customer_list = df.iloc[:, 0].str.lower().tolist()  # Assuming customer names are in the first column
    customer_ids = df.iloc[:, 1].tolist()  # Assuming customer IDs are in the second column
    return customer_list, customer_ids

# Function to find the closest match using Levenshtein distance
def closest_match(phrase, match_list, match_ids, threshold=2):
    closest_match = None
    closest_id = None
    min_distance = float('inf')
    
    for item, item_id in zip(match_list, match_ids):
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
def process_line(line, product_list, product_ids):
    cleaned_line = clean_line(line)
    words = word_tokenize(cleaned_line)
    count = None
    product = None
    product_id = None
    
    for i, word in enumerate(words):
        if word.isdigit():
            count = int(word)
            potential_product_words = words[i+1:]
            all_permutations = generate_permutations(potential_product_words)
            for permutation in all_permutations:
                permutation_phrase = " ".join(permutation)
                product, product_id = closest_match(permutation_phrase, product_list, product_ids)
                if product:
                    return count, product, product_id  # Exit as soon as a match is found
            break
    
    if count is None:
        # Assume 1 case if no count is provided
        count = 1
        all_permutations = generate_permutations(words)
        for permutation in all_permutations:
            permutation_phrase = " ".join(permutation)
            product, product_id = closest_match(permutation_phrase, product_list, product_ids)
            if product:
                return count, product, product_id  # Exit as soon as a match is found
    
    return count, product, product_id

# Main function to extract orders from email text
def extract_orders(email_text, product_list, product_ids):
    orders = []
    lines = email_text.strip().split('\n')
    
    for line in lines:
        if line.strip():  # Skip empty lines
            count, product, product_id = process_line(line, product_list, product_ids)
            if product:
                orders.append((count, product, product_id))
    
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
    
    for count, product, product_id in orders:
        cursor.execute('''
        INSERT INTO orders (customer, customer_id, cases, lbs, item, item_id, date_generated, date_processed)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (customer, customer_id, count, 0, product, product_id, date_generated, date_processed))
    
    conn.commit()
    conn.close()

# Sample email text
email_text = """
2 cases  Plain Egg Linguine
1 Plain Egg Fettuccine
3 sheets Lasagne (extra thin!!!!)

Thanks. Please deliver to 166 culverview lane, Branchville, NJ, 07826
"""

# Read product list and IDs from Excel
product_list, product_ids = read_product_list_from_excel('products.xlsx')

# Read customer list and IDs from Excel
customer_list, customer_ids = read_customer_list_from_excel('customers.xlsx')

# Print the customer list
print("Customer List:", customer_list)
print("Customer IDs:", customer_ids)

# Assume we have a function to determine the customer from the email text
customer = "john doe"  # Example customer name
customer_id = 1  # Example customer ID

# Extract orders from the email text
orders = extract_orders(email_text, product_list, product_ids)
print("Orders:", orders)

# Write orders to the SQL database
write_orders_to_db('orders.db', customer, customer_id, orders)
