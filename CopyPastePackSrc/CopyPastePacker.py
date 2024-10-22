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
from tkinter import messagebox

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

class App():
    def __init__(self, master, order_dict):
        self.master = master
        
        
        self.order_dict = order_dict
        self.item_list = tk.Listbox(self.master, font=("Arial", 20))
        self.item_list.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        
        self.barcode_var = tk.StringVar()
        self.package_var = tk.StringVar()
        self.order_label = tk.Label(self.master, text="")
        
        self.package_label = tk.Label(self.master, text="")
        self.package_entry = tk.Entry(self.master, textvariable=self.package_var)

        self.barcode_label = tk.Label(self.master, text="Scan order barcode:", font=("Arial", 20))
        self.barcode_entry = tk.Entry(self.master, textvariable=self.barcode_var)
        self.barcode_entry.focus()
        self.barcode_entry.bind("<Return>", self.check_order)
        
        self.barcode_label.pack(side=tk.TOP)
        self.barcode_entry.pack(side=tk.TOP)
        
        self.abort_btn = tk.Button(self.master, text="Abort Order", command=self.abort_order, height=4)
        self.abort_btn.pack(side=tk.BOTTOM)

        self.item_entry = None
        self.current_order = None
        
        self.count_var = tk.StringVar()
        
        self.count_button = tk.Button(self.master, textvariable=self.count_var, command = self.display_remaining, height =4)
        self.count_button.pack(side = tk.BOTTOM)
        
        self.update_count()
        
    def abort_order(self):
        self.current_order = None
        self.reset_window()
        self.barcode_entry.bind("<Return>", self.check_order)
    
    def display_remaining(self):
        conn = sqlite3.connect(path_to_db)
        cursor = conn.cursor()
        
        # Get today's date
        today = datetime.date.today()

        # Execute the SQL query
        cursor.execute("""
        SELECT DISTINCT order_num
        FROM orders
        WHERE packedtimestamp IS NULL AND DATE(generatedtimestamp) = ?
        """, (today,))
        
        order_ids = cursor.fetchall()
        self.item_list.delete(0, tk.END)
        self.item_list.insert(tk.END, "Order IDs:")
        for order_id in order_ids:
            self.item_list.insert(tk.END, order_id[0])
        conn.close()

    def update_count(self):
        
        conn = sqlite3.connect(path_to_db)
        cursor = conn.cursor()
        
        # Get today's date
        today = datetime.date.today()
        
        query = """
        SELECT COUNT(DISTINCT order_num)
        FROM orders
        WHERE packedtimestamp IS NULL AND DATE(generatedtimestamp) = ?
        """
        
        cursor.execute(query, (today,))
        count = cursor.fetchone()[0]
        self.count_var.set(f"Remaining Orders: {count}")
        
        
        conn.commit()
        conn.close()
    
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
        self.item_list.delete(0, tk.END)
        self.barcode_var.set("")
        self.package_var.set("")
        self.package_label.config(text="")
        self.order_label.pack_forget()
        self.barcode_entry.pack(side=tk.TOP)
        self.barcode_label.pack(side=tk.TOP)
        self.barcode_entry.focus()
        if self.item_entry is not None:
            self.item_entry.destroy()
            self.item_entry = None
            
    def check_order(self, event):
        global order_num
        order_num = self.barcode_var.get()
        if order_num in self.order_dict:
            self.current_order = self.order_dict[order_num]
            self.item_list.delete(0, tk.END)
            for item in self.current_order:
                item_name = item[0]
                item_count = item[2]
                item_barcode = item[1]
                item_text = f"{item_barcode} - {item_name} ({item_count})"
                self.item_list.insert(tk.END, item_text.center(50))
            self.barcode_entry.pack_forget()
            self.barcode_label.pack_forget()
            self.barcode_entry.unbind("<Return>")
            self.barcode_entry.delete(0, tk.END)
            self.item_entry = tk.Entry(self.master)
            
            self.item_entry.pack(side=tk.TOP)
            self.item_entry.focus()
            self.item_entry.bind("<Return>", self.scan_item)
        else:
            self.item_list.delete(0, tk.END)
            self.item_list.insert(tk.END, "Order not found")

    def scan_item(self, event):
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
            self.current_order = [(item[0], item[1], count_remaining) if item[1] == item_barcode else item for item in self.current_order]
    
        self.item_entry.delete(0, tk.END)
        
        
        #Order has been filled 
        if len(self.current_order) == 0:
            
            self.item_entry.unbind("<Return>")
            self.item_entry.delete(0, tk.END)
          
            self.item_list.delete(0, tk.END)
            
            # Connect to the database
            conn = sqlite3.connect(path_to_db)
            cursor = conn.cursor()
            
            global packaged_order_dict
            # Get the current date and time
            now = datetime.datetime.now()
            
            # Format the date and time as mm/dd/yy and a time
            packedtimestamp = now
            global order_num
            
            # Execute the SQL statement to update the timestamp
            sql = "UPDATE orders SET packedtimestamp = ? WHERE order_num = ?"
            cursor.execute(sql, (packedtimestamp, order_num))
            
            # Execute the SQL statement to update the timestamp
            sql = "UPDATE ordersandtimestampsonly SET packedtimestamp = ? WHERE order_num = ?"
            cursor.execute(sql, (packedtimestamp, order_num))
            
            
            # Commit the changes and close the connection
            conn.commit()
            conn.close()
            
            self.update_count()

            self.reset_window()
            self.barcode_entry.focus()
            self.barcode_entry.bind("<Return>", self.check_order)
    
        else:
            self.item_list.delete(0, tk.END)
            for item in self.current_order:
                item_name = item[0]
                item_count = item[2]
                item_barcode = item[1]
                if item_barcode == 0:
                    remaining_items = count_remaining
                else:
                    remaining_items = remaining_counts.get(item_barcode, item_count)
                item_text = f"{item_barcode} - {item_name} ({item_count})"
                self.item_list.insert(tk.END, item_text.center(50))
            self.item_entry.focus()

    def run(self):
        #self.exit_btn.pack(side=tk.BOTTOM)
        self.master.mainloop()

def read_excel_file(file_path):
    """
    Reads an Excel file and returns a list of phrases where each phrase is "item_name barcode"
    """
    # Read the Excel file into a dataframe without headers
    df = pd.read_excel(file_path, engine='openpyxl', header=None, dtype=str)
    
    # If the file does not have named columns, assume the first column is "Item Name" and second is "UPC Code"
    if 0 not in df.columns and 1 not in df.columns:
        df.columns = ["Item Name", "UPC Code"]
    else:
        df.rename(columns={0: "Item Name", 1: "UPC Code"}, inplace=True)

    # Convert item names to lowercase
    df["Item Name"] = df["Item Name"].str.lower()

    # If UPC codes have ".0" at the end, remove it
    df["UPC Code"] = df["UPC Code"].str.rstrip('.0')

    phrases = []
    for _, row in df.iterrows():
        item_name = row["Item Name"]
        upc_code = row["UPC Code"]
        phrases.append(f"{item_name} {upc_code}")

    return phrases

def start_packing_sequence(order_dict):
    root = tk.Tk()
    root.title("Order Packer")
    root.geometry("1600x800")
    app = App(root, order_dict)
    app.run()

def split_data(text, keywords):
    # Copy the list of keywords
    repeat_keywords = list(keywords)
    
    # Initialize the dictionary to store the result
    results = []
    
    # While loop to handle multiple blocks of text
    while len(repeat_keywords) > 0:
        # Only consider keywords that are in the text
        present_keywords = [keyword for keyword in repeat_keywords if keyword in text]
        
        if not present_keywords:
            break
        
        # Sort keywords based on their order of appearance in the text
        present_keywords.sort(key=text.index)
        
        # Initialize a dictionary for this block of text
        block_results = {}
        
        # Loop through the keywords
        for i in range(len(present_keywords)):
            keyword = present_keywords[i]
            remaining_text = text.split(keyword, 1)[1].lstrip()

            # Check if there's another keyword or newline character after the current keyword
            end_indices = [idx for idx in [remaining_text.find(kw) for kw in present_keywords if kw != keyword] + [remaining_text.find('\n')] if idx != -1]
            end_index = min(end_indices) if end_indices else None
            
            # If there's another keyword or newline character after the current keyword, use it to split the text
            if end_index is not None:
                block_results[keyword] = remaining_text[:end_index].strip()  # remove leading/trailing whitespace
                text = keyword + remaining_text[end_index:]
            else:
                # If not, all the remaining text belongs to the current keyword
                block_results[keyword] = remaining_text.strip()  # remove leading/trailing whitespace
                text = ''
                break

        results.append(block_results)
        repeat_keywords = [kw for kw in repeat_keywords if kw not in block_results]
    
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
    # Split the text on the keyword
    parts = text.split(keyword)
    
    # The first part doesn't need the keyword prepended
    blocks = [parts[0]]
    
    # For all other parts, prepend the keyword
    for part in parts[1:]:
        blocks.append(keyword + part)
    
    return blocks

def enter_text(acceptable_phrases):
    with open(path_to_txt, "r") as file:
        text_input = file.read()
    order_dict = {}
    packaged_order_dict = {}
    
    text_input = (text_input.lower()).strip()
    
    # Connect to the database (create it if it doesn't exist)
    conn = sqlite3.connect(path_to_db)

    # Create a cursor to execute SQL commands
    c = conn.cursor()   
    
    c.execute("SELECT DISTINCT * FROM blockseperator WHERE keywordstype = 'blockseperator'")
    # Loop through the rows returned by the query
    rows = c.fetchall()
    # Loop through each row and append the values to the respective lists
    for row in rows:
        blockseperatorkeyword = (row[1]).lower()  # Assuming 'phrases' is the 2nd column
        packagenumberphrase = (row[2]).lower()
        
    conn.close()
    
    blocks = split_blocks(text_input, blockseperatorkeyword)
    print_blocks(blocks)


    for block in blocks:
        order_num = ""
        phrases = block.split("\n")
        modified_phrases = []  #\ list to store modified phrases
        phrases_length = len(phrases)
        phrasequantity = None

        print("PHRASES", phrases)
        print("PACKAGENUMBERPHRASE", packagenumberphrase)

        # Extract order number
        for phrase in phrases:
            match = re.search(f"{re.escape(packagenumberphrase)} ([^\s]+)", phrase)
            if match:
                order_num = ''.join([char for char in match.group(1) if char.isdigit()])
                print(f"Found order number: {order_num}")
                break
        
        pairKeywordFound = False
        global keywords_2d_list
        global removals_list 
        splitting_keywords_list = [None] * len(keywords_2d_list)
        keyword_positions = dict()

        for i, line in enumerate(block.split("\n")):
            for j, keywords_list in enumerate(keywords_2d_list):
                for keyword in keywords_list:
                    keyword_position = line.find(keyword)
                    if keyword_position != -1: # keyword is in line
                        # Store the line and position within line for this keyword
                        current_position = (i, keyword_position)
                        if keyword not in keyword_positions or current_position < keyword_positions[keyword]:
                            keyword_positions[keyword] = current_position
        
        # Now, find the keyword with the lowest index
        splitting_keywords_list = [None] * len(keywords_2d_list)
        for keyword, position in keyword_positions.items():
            for i, keywords_list in enumerate(keywords_2d_list):
                if keyword in keywords_list:
                    # If this is the first keyword we've seen for this keywords_list, or if the current keyword
                    # has a lower index than the stored keyword, update the stored keyword
                    if splitting_keywords_list[i] is None or position < keyword_positions[splitting_keywords_list[i]]:
                        splitting_keywords_list[i] = keyword
        
        #Pairing Logic
        processed_blocklist = []  # Define processed_blocklist before the loop
        stringtobeadded = ""
        
        for i, keywords_list in enumerate(keywords_2d_list):
            splittingkeyword = splitting_keywords_list[i]
            if splittingkeyword and splittingkeyword in block:  # if the first keyword is in the block
                # Split the block on the keyword
                sub_blocks = block.split(splittingkeyword)
                
                # re-add the keyword and filter out empty strings
                sub_blocks = [splittingkeyword + sub_block for sub_block in sub_blocks if sub_block.strip()] 
                prev_sub_block = ""
                print(f"Sub-blocks for keyword '{splittingkeyword}':")
                for sb in sub_blocks:
                    print(sb)
                    print("--------------------------------------------------------------------------------------------")

                for sub_block in sub_blocks: 
                    #We get back list of dicts    
                    processed_block = split_data(sub_block, keywords_list)  # process each sub block
                    for dict_ in processed_block:
                        for keyword in keywords_list:
                            if dict_ is not None and keyword in dict_:  # if keyword is in this dictionary
                                stringtobeadded += dict_[keyword] + " "
                    
                    #Find the most recent phrasequantity phrase and extract the phrasequantity
                    global quantityphrases
                    global quantitypositions
                    phrasequantity = 1
                    last_line = ""
                    last_phrase = ""

                    print("PREV_SUB_BLOCK", prev_sub_block)
                    
                    lines = prev_sub_block.splitlines()
                    for line in lines:
                        for quantityphrase in quantityphrases:
                            if quantityphrase in line:
                                last_line = line
                                last_phrase = quantityphrase
                    
                    for quantindex, phrase in enumerate(quantityphrases):
                        if phrase in last_phrase:
                            seperatedphrase = last_line.split()
                            phrasequantity = seperatedphrase[int(quantitypositions[quantindex])]
                                
                    prev_sub_block = sub_block#for phrasequantity
                    
                            
                    for removal in removals_list:
                        stringtobeadded = stringtobeadded.replace(removal, "")
                    stringtobeadded = ' '.join(stringtobeadded.split())
                    for _ in range(int(phrasequantity)):
                        modified_phrases.append(stringtobeadded)
                    stringtobeadded = ""

        
        quantityphrase = 1
        i = 0
        incompletekeyword = ""
        phrasequantity = 1
        while i < len(phrases):
            phrase = phrases[i].replace("’", "'")
            phrase = phrase.strip()
            packagenumberphrase = packagenumberphrase.strip()
            
            # print("PACKAGENUMBERPHRASE ", packagenumberphrase)
            # print("PHRASE ", phrase)

            # # Try to match "package: <number>"
            # match = re.search(r'package:\s*#?(\d+)', phrase, re.IGNORECASE)
            # if match:
            #     order_num = match.group(1)
            #     print(f"Found order number from 'package:': {order_num}")
            #     break

            # match = re.search(f"{re.escape(packagenumberphrase)} ([^\s]+)", phrase)
            # if match:
            #     order_num = ''.join([char for char in match.group(1) if char.isdigit()])
            #     print(f"Found order number: {order_num}")
            #     break
        
            for quantindex, quantphrase in enumerate(quantityphrases):
                if quantphrase.replace("’","'") in phrase:
                    seperatedphrase = phrase.split()
                    phrasequantity = int(seperatedphrase[int(quantitypositions[quantindex])])
            
            global incompletephrases
            global secondarykeywords
            
            for incompindex, incomp in enumerate(incompletephrases):
                if incomp in phrase:
                    incompletekeyword = secondarykeywords[incompindex]
            
            if incompletekeyword:
                print("phrasequantity " + str(phrasequantity))
                for _ in range(phrasequantity):
                    modified_phrases.append(phrase + " " + incompletekeyword)
                    print((phrase + " " + incompletekeyword).strip())
            
            global exactphrases
            global exactphraseitems_2d
            
            #Exact phrases may be multi-lined
            for exactindex, exactphrase in enumerate(exactphrases):
                linebreaks = exactphrase.count("\n")
                remaining_lines = phrases[i:i + linebreaks + 1]
                blockforexactcheck = "\n".join(remaining_lines)
                if exactphrase in blockforexactcheck:
                    for item in exactphraseitems_2d[exactindex]:
                        for _ in range(phrasequantity):
                            modified_phrases.append(item)
              
                
            else:
                for _ in range(phrasequantity):
                    modified_phrases.append(phrase)
            i +=1
           
         #Fix extra LASAGNE IN ORDER
           
        for modified_phrase in modified_phrases:
            for acceptable_phrase in acceptable_phrases:
                acceptable_words = acceptable_phrase.split()[:-1]  # exclude the last word (barcode)
                modified_words = modified_phrase.split()
                
                if set(modified_words) == set(acceptable_words):
                    barcode = acceptable_phrase.split()[-1]  # get the barcode
                    item = " ".join(acceptable_words)  # join the remaining words as the item
                    
                    if order_num not in order_dict:
                        order_dict[order_num] = [(item, barcode, 1)]
                    else:
                        item_found = False
                        for i, (item_, barcode_, count) in enumerate(order_dict[order_num]):
                            if item_ == item and barcode_ == barcode:
                                item_found = True
                                order_dict[order_num][i] = (item_, barcode_, count+1)
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
        
        # Get column names
        first_col = df.columns[0]  # First column
        second_col = df.columns[1]  # Second column
        
        # Create a view of valid rows (doesn't modify original Excel)
        valid_products = df[df[first_col].notna() & df[second_col].notna()].copy()
        
        # Filter out any rows that contain headers
        valid_products = valid_products[~valid_products[first_col].str.lower().isin(['item name', 'product name', 'name', 'upc', 'upc code'])]
        
        # Rename columns to our standard names
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

# Step 3: Define a function to prompt user to choose an option and perform that option
def choose_option():
    global acceptable_phrases
    global order_dict
    global packaged_order_dict
    acceptable_phrases = read_acceptable_phrases()
    while True:
        option = input("""Choose an option:
1. Upload today's orders from input.txt file
2. Start Packing Today's Orders
3. Check customer order numbers and related phrases
4. Delete today's orders
5. View UPC Codes
6. Exit
7. Today's Products at a Glance\n""")
        if option == "4":
            option = input("Are you sure you want to delete today's orders? Yes/No")
            if option.lower() == "yes" or option.lower() == "y":
                conn = sqlite3.connect(path_to_db)
                cursor = conn.cursor()
                
                # Get today's date
                today = datetime.date.today()
                
    
                # Execute the SQL query
                cursor.execute("""
                DELETE
                FROM orders
                WHERE DATE(generatedtimestamp) = ?
                """, (today,))
                
                # Execute the SQL query
                cursor.execute("""
                DELETE
                FROM ordersandtimestampsonly
                WHERE DATE(generatedtimestamp) = ?
                """, (today,))
                
                conn.commit()
                conn.close()
                print("Today's orders deleted")
           
        elif option == "1":
            
            # Connect to the SQLite database
            conn = sqlite3.connect(path_to_db)
            c = conn.cursor()
            
            # Query the database for records where the keywordstype is 'pairing'
            c.execute("SELECT DISTINCT * FROM pairings WHERE keywordstype = 'pairing'")
            
            # Initialize the lists to hold the keywords and removals
            
            global keywords_2d_list
            global removals_list
            
            # Loop through the rows returned by the query
            for row in c.fetchall():
                keywordstype, keywords_string, removals_string = row
                keywords_string = keywords_string.lower()  # Convert to lowercase
                removals_string = removals_string.lower()  # Convert to lowercase
                
                # Split the strings back into lists
                keywords_list = keywords_string.split('<')
                removals = removals_string.split('<')
            
                # Add the lists to the appropriate 2D list or list
                keywords_2d_list.append(keywords_list)
                removals_list.extend(removals)
                
            # Query the database for records where the keywordstype is 'exactphrase'
            c.execute("SELECT DISTINCT * FROM exactphrases WHERE keywordstype = 'exactphrase'")
            
            # Initialize the lists to hold the keywords and removals
            
            global exactphrases
            global exactphraseitems_2d
            
            # Loop through the rows returned by the query
            for row in c.fetchall():
                keywordstype, exactphrases_string, exactphraseitems_string = row
            
                exactphrases_string = exactphrases_string.lower()  # Convert to lowercase
                exactphraseitems_string = exactphraseitems_string.lower()  #    
            
                # Split the strings back into lists
                exactphrases_list = exactphrases_string.split('<')
                exactitems = exactphraseitems_string.split('<')
            
                # Add the lists to the appropriate 2D list or list
                exactphrases.extend(exactphrases_list)
                exactphraseitems_2d.append(exactitems)
            
            # Connect to the database (create it if it doesn't exist)
            conn = sqlite3.connect(path_to_db)

            # Create a cursor to execute SQL commands
            cursor = conn.cursor()   
            
            # Query the database for records where the keywordstype is 'pairing'
            c.execute("SELECT DISTINCT * FROM quantitys WHERE keywordstype = 'Quantity'")
            
            global quantityphrases
            global quantitypositions
            
            # Loop through the rows returned by the query
            rows = c.fetchall()
            # Loop through each row and append the values to the respective lists
            for row in rows:
                quantityphrase = row[1].lower()  # Convert to lowercase
                quantityphrases.append(quantityphrase)  # Assuming 'phrases' is the 2nd column
                quantitypositions.append(int(row[2]))  # Assuming 'positions' is the 3rd column

            print("quantity phrases", quantityphrases)
            print("quantitypositions", quantitypositions)
                
            
            c.execute("SELECT DISTINCT * FROM incompletephrases WHERE keywordstype = 'Incomplete Phrase'")
            
            global incompletephrases
            global secondarykeywords
            
            # Loop through the rows returned by the query
            rows = c.fetchall()
            # Loop through each row and append the values to the respective lists
            for row in rows:
                incompletephrase = (row[1].lower()).strip()  # Convert to lowercase
                secondarykeyword = (row[2].lower()).strip()  # Convert to lowercase
                incompletephrases.append(incompletephrase)  # Assuming 'phrases' is the 2nd column
                secondarykeywords.append(secondarykeyword)  # Assuming 'positions' is the 3rd column

                print(incompletephrases)
                print(secondarykeywords)

            order_dict = enter_text(acceptable_phrases)
            
            # Get the current date and time
            generatedtimestamp = datetime.datetime.now()
            
            # Format the date and time as mm/dd/yy and a time
            #generatedtimestamp = now.strftime("%m/%d/%Y %I:%M %p")
            
            for order_num, items in order_dict.items():
                for item, barcode, count in items:
                    cursor.execute("INSERT INTO orders (order_num, item, item_barcode, count, generatedtimestamp) VALUES (?, ?, ?, ?, ?)",
                                (order_num, item, barcode, count, generatedtimestamp))

            for order_num in order_dict:
                cursor.execute("INSERT INTO ordersandtimestampsonly (order_num, generatedtimestamp) VALUES (?, ?)",
                            (order_num, generatedtimestamp))
            conn.commit()
            conn.close()
        elif option == "3":
            print_orders(order_dict)
        elif option == "2":
            try:
                start_packing_sequence(order_dict)
            except NameError:
                print("You must enter text into the input file to begin packing")
           
        elif option == "6":
            break
        elif option == "7":
            display_todays_products()
        elif option == "5":
            display_phrases_as_table(acceptable_phrases)
        else:
            print("Invalid option. Try again.")

# Step 4: Call choose_option() function to start the program
choose_option()
