# -*- coding: utf-8 -*-
"""
Created on Wed Aug 30 12:11:38 2023

@author: Derek
"""

import tkinter as tk
from tkinter import ttk
import sqlite3
from tkinter import messagebox
import os
import sys

# Get the directory where the script or executable is located
if getattr(sys, 'frozen', False):
    # If the application is frozen with PyInstaller, use this path
    application_path = os.path.dirname(sys.executable)
else:
    # Otherwise use the path to the script file
    application_path = os.path.dirname(os.path.abspath(__file__))

# Construct full paths to your files
path_to_db = os.path.join(application_path, 'CopyPastePack.db')
path_to_xlsx = os.path.join(application_path, 'UPCCodes.xlsx')
path_to_txt = os.path.join(application_path, 'input.txt')

column_names_global = []
detail_windows = []  # List to keep track of open detail windows

def construct_sql_query_with_params(table_name, column_names, record_data):
    conditions = []
    params = []

    for col, value in zip(column_names, record_data):
        if value == 'None':
            conditions.append(f"{col} IS NULL")
        else:
            conditions.append(f"{col} = ?")
            params.append(value)

    query_conditions = " AND ".join(conditions)
    sql_query = f"SELECT * FROM {table_name} WHERE {query_conditions}"
    return sql_query, params

def delete_record_from_db(table_name, record_data):
    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()

    # Get column names for the table
    c.execute(f"PRAGMA table_info({table_name})")
    column_info = c.fetchall()
    column_names = [col[1] for col in column_info]

    # Construct a DELETE query that matches all column values
    delete_query, params = construct_sql_query_with_params(table_name, column_names, record_data)
    delete_query = delete_query.replace("SELECT *", "DELETE")  # Convert the select query to a delete query
    
    # Execute the DELETE query
    c.execute(delete_query, params)
    conn.commit()
    conn.close()
    
def on_delete_key_pressed(event, main_tree, search_tree):
    # Detect which treeview the event comes from
    origin_tree = event.widget

    item = origin_tree.selection()  # Get selected items from the originating treeview
    if not item:
        return

    record_data = origin_tree.item(item[0], 'values')

    # Confirm deletion
    msg = f"Are you sure you want to delete the selected record?"
    if not messagebox.askyesno("Confirm Deletion", msg):
        return

    table_name = table_combobox.get()
    delete_record_from_db(table_name, record_data)
    
    # Close all detail windows
    close_all_detail_windows()
    
    # Refresh data
    display_data()

    # Also remove the item from the originating Treeview
    if origin_tree.exists(item[0]):
        origin_tree.delete(item[0])

def on_search_tree_click(event):
    item = search_tree.selection()
    if not item:
        return
    record_data = search_tree.item(item[0], 'values')
    display_record_data(record_data)

def display_data():
    global tree
    global column_names_global
    table_name = table_combobox.get()
    if not table_name or table_name not in table_names:
        return

    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()

    c.execute(f"PRAGMA table_info({table_name})")
    column_info = c.fetchall()
    column_names = [col[1] for col in column_info]

    search_column_combobox['values'] = column_names
    search_column_combobox.set("")

    tree.delete(*tree.get_children())
    tree['columns'] = column_names

    for col in column_names:
        tree.column(col, width=100, anchor=tk.W, stretch=tk.NO)
        tree.heading(col, text=col)

    c.execute(f"SELECT * FROM {table_name}")
    data = c.fetchall()

    conn.close()

    for row in data:
        tree.insert('', tk.END, values=row)
    
    column_names_global = column_names


def on_tree_click(event):
    item = tree.selection()
    if not item:
        return
    record_data = tree.item(item[0], 'values')

    # Convert everything to string for consistency
    record_data_as_str = tuple(map(str, record_data))

    # Build the SQL query to fetch the selected record from the database
    columns = tree['columns']
    conditions = ' AND '.join([f"{col} = ?" for col in columns])
    sql_query = f"SELECT * FROM {table_combobox.get()} WHERE {conditions}"

    # Fetch the record from the database
    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()
    c.execute(sql_query, record_data_as_str)
    selected_record = c.fetchone()
    conn.close()

    display_record_data(record_data_as_str)

def display_record_data(data):
    # Create a new window to display the record details
    new_window = tk.Toplevel(root)
    new_window.title("Record Details")
    record_label_new_window = tk.Label(new_window, text=", ".join(map(str, data)), font=("Arial", 14, "bold"), wraplength=800)
    record_label_new_window.pack()
    
    # Add the window to our list of detail windows
    detail_windows.append(new_window)

# Add these lines to create empty lists
search_entries = []
search_columns = []

# Initialize column position for search bars
next_search_column_pos = 4

# This function will add another search row
def add_search_row():
    global next_search_column_pos
    global column_names_global  # Use global inside function

    new_search_entry = ttk.Entry(center_frame)
    new_search_entry.grid(row=4, column=next_search_column_pos, padx=5, pady=5)

    new_search_column_combobox = ttk.Combobox(center_frame, values=column_names_global)  # Use column names here
    new_search_column_combobox.grid(row=5, column=next_search_column_pos, padx=5, pady=5)

    search_entries.append(new_search_entry)
    search_columns.append(new_search_column_combobox)
    
    next_search_column_pos += 1  # Move to next column position for next search bar

def on_search():
    search_tree.delete(*search_tree.get_children())
    search_tree['columns'] = tree['columns']
    for col in search_tree['columns']:
        search_tree.heading(col, text=col)
        search_tree.column(col, width=100, anchor=tk.W, stretch=tk.NO)

    found_records = False

    conditions = []

    for search_entry, search_column in zip(search_entries, search_columns):
        search_query = search_entry.get()
        column_name = search_column.get()
        
        if search_query and column_name:
            conditions.append(f"{column_name} LIKE ?")

    if conditions:
        condition_str = " AND ".join(conditions)
        sql_query = f"SELECT * FROM {table_combobox.get()} WHERE {condition_str}"

        params = [f"%{search_entry.get()}%" for search_entry in search_entries if search_entry.get()]

        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()
        c.execute(sql_query, params)

        data = c.fetchall()
        conn.close()

        if data:
            for row in data:
                search_tree.insert('', tk.END, values=row)
            found_records = True
            search_record_label.config(text=f"Number of records found: {len(data)}")
        else:
            search_tree.insert('', tk.END, values=("No matching records found.",))
            search_record_label.config(text=f"Number of records found: 0")

def get_table_names():
    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()
    c.execute("SELECT name FROM sqlite_master WHERE type='table';")
    table_names = c.fetchall()
    conn.close()
    return [name[0] for name in table_names]

def on_closing():
    # Here, we can add any logic to handle window closing, such as closing database connections or saving data.
    # But since you're opening and closing connections within each function, there's nothing specific to add here.
    root.destroy()

def close_all_detail_windows():
    global detail_windows
    for window in detail_windows:
        if window.winfo_exists():
            window.destroy()
    detail_windows = []
    
def clear_dashboard():
    # Clear the main Treeview
    tree.delete(*tree.get_children())
    
    # Clear the search Treeview
    search_tree.delete(*search_tree.get_children())
    
    # Clear search entries and remove additional search rows
    for entry, column in zip(search_entries, search_columns):
        entry.grid_forget()
        column.grid_forget()
        entry.delete(0, tk.END)
        column.set("")
    search_entries.clear()
    search_columns.clear()
    
    # Close all detail windows
    close_all_detail_windows()
    
    # Reset the table Combobox
    table_combobox.set("")
    
    # Reset the search record label
    search_record_label.config(text="")
    
    # Reset global column position for search bars
    global next_search_column_pos
    next_search_column_pos = 4
    
    # Re-add the original search entry and dropdown
    search_entry.grid(row=4, column=1, padx=5, pady=5)
    search_column_combobox.grid(row=5, column=1, padx=5, pady=5)
    search_entries.append(search_entry)
    search_columns.append(search_column_combobox)
    
    # Reset these to their default state
    search_entry.delete(0, tk.END)
    search_column_combobox.set("")

def bulk_delete_records():
    # Get current search conditions
    conditions = []
    params = []

    for search_entry, search_column in zip(search_entries, search_columns):
        search_query = search_entry.get()
        column_name = search_column.get()
        
        if search_query and column_name:
            conditions.append(f"{column_name} LIKE ?")
            params.append(f"%{search_query}%")

    if not conditions:
        messagebox.showwarning("Warning", "Please enter search criteria before attempting bulk delete.")
        return

    # Construct and execute a SELECT query first to show how many records will be affected
    condition_str = " AND ".join(conditions)
    select_query = f"SELECT COUNT(*) FROM {table_combobox.get()} WHERE {condition_str}"

    conn = sqlite3.connect(path_to_db)
    c = conn.cursor()
    c.execute(select_query, params)
    record_count = c.fetchone()[0]
    conn.close()

    if record_count == 0:
        messagebox.showinfo("Info", "No records match the search criteria.")
        return

    # Create a custom confirmation dialog
    confirm_dialog = tk.Toplevel(root)
    confirm_dialog.title("Confirm Bulk Delete")
    confirm_dialog.geometry("400x150")
    confirm_dialog.transient(root)  # Make dialog modal
    confirm_dialog.grab_set()
    
    # Center the dialog on the screen
    confirm_dialog.geometry(f"+{root.winfo_x() + 150}+{root.winfo_y() + 150}")

    # Warning message
    message = f"Are you sure you want to delete {record_count} records?\nThis action cannot be undone!"
    warning_label = tk.Label(confirm_dialog, text=message, wraplength=350, pady=20)
    warning_label.pack()

    # Button frame
    button_frame = tk.Frame(confirm_dialog)
    button_frame.pack(pady=10)

    def on_yes():
        # Execute the delete query
        delete_query = f"DELETE FROM {table_combobox.get()} WHERE {condition_str}"
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()
        c.execute(delete_query, params)
        conn.commit()
        conn.close()
        
        # Close all detail windows
        close_all_detail_windows()
        
        # Refresh the display
        display_data()
        
        # Clear the search results
        search_tree.delete(*search_tree.get_children())
        search_record_label.config(text="")
        
        messagebox.showinfo("Success", f"{record_count} records have been deleted.")
        confirm_dialog.destroy()

    def on_no():
        confirm_dialog.destroy()

    # Create buttons - No button is focused by default
    no_button = tk.Button(button_frame, text="No", command=on_no, width=10)
    yes_button = tk.Button(button_frame, text="Yes", command=on_yes, width=10)
    
    no_button.pack(side=tk.LEFT, padx=10)
    yes_button.pack(side=tk.LEFT, padx=10)
    
    # Set focus to "No" button and bind Enter key
    no_button.focus_set()
    
    # Bind Enter key to "No" button and Escape key to dialog
    confirm_dialog.bind('<Return>', lambda e: on_no())
    confirm_dialog.bind('<Escape>', lambda e: on_no())


root = tk.Tk()
root.title("SQL Dashboard")

center_frame = tk.Frame(root)
center_frame.pack()

table_names = get_table_names()

# Add direction label here
table_direction_label = tk.Label(center_frame, text="Enter table:")
table_direction_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)

table_combobox = ttk.Combobox(center_frame, values=table_names)
table_combobox.grid(row=0, column=0, padx=5, pady=5, columnspan=3)

# Create a separate frame for Treeview
treeview_frame = tk.Frame(center_frame, width=1000, height=200)
treeview_frame.grid(row=1, column=0, columnspan=20, padx=5, pady=5)
treeview_frame.grid_propagate(False)  # Stops frame from resizing to fit treeview

tree = ttk.Treeview(treeview_frame, height=10, show='headings')
tree.place(relheight=1, relwidth=1)  # Here, you place it and control its size

# Add a horizontal scrollbar
scroll_x = tk.Scrollbar(treeview_frame, orient='horizontal', command=tree.xview)
scroll_x.place(relx=0, rely=0.95, relwidth=1)  # Changed pack to place
tree.configure(xscrollcommand=scroll_x.set)

refresh_button = tk.Button(center_frame, text="Refresh Data", command=display_data)
refresh_button.grid(row=2, column=0, padx=5, pady=5, columnspan=3)

record_label = tk.Label(center_frame, font=("Arial", 14, "bold"), wraplength=400)
record_label.grid(row=3, column=0, padx=5, pady=5, columnspan=3)

direction_label2 = tk.Label(center_frame, text="Enter search query:")
direction_label2.grid(row=4, column=0, padx=5, pady=5)
search_entry = ttk.Entry(center_frame)
search_entry.grid(row=4, column=1, padx=5, pady=5)

direction_label3 = tk.Label(center_frame, text="Select column:")
direction_label3.grid(row=5, column=0, padx=5, pady=5)
search_column_combobox = ttk.Combobox(center_frame)
search_column_combobox.grid(row=5, column=1, padx=5, pady=5)

# Add these lines to add the first search entry and dropdown to the lists
search_entries.append(search_entry)
search_columns.append(search_column_combobox)

# Create a frame for the search and bulk delete buttons
button_frame = tk.Frame(center_frame)
button_frame.grid(row=4, column=2, rowspan=2, padx=5, pady=5)

search_button = tk.Button(button_frame, text="Search", command=on_search)
search_button.pack(pady=2)

bulk_delete_button = tk.Button(button_frame, text="Bulk Delete", command=bulk_delete_records)
bulk_delete_button.pack(pady=2)

# Add this button to create new search bars
add_search_button = tk.Button(center_frame, text="Add Search", command=add_search_row)
add_search_button.grid(row=4, column=3, rowspan=2, padx=5, pady=5)

# Replace these lines to pack search_treeview similar to treeview
search_treeview_frame = tk.Frame(center_frame, width=1000, height=200)
search_treeview_frame.grid(row=6, column=0, padx=5, pady=5, columnspan=3)
search_treeview_frame.grid_propagate(False)  # Disable resizing of frame

search_tree = ttk.Treeview(search_treeview_frame, height=10, show='headings')
search_tree.place(relheight=1, relwidth=1)  # Use the place method to control its size

# Add a horizontal scrollbar for search Treeview
search_scroll_x = tk.Scrollbar(search_treeview_frame, orient='horizontal', command=search_tree.xview)
search_scroll_x.place(relx=0, rely=0.95, relwidth=1)  # Changed pack to place
search_tree.configure(xscrollcommand=search_scroll_x.set)

search_record_label = tk.Label(center_frame, font=("Arial", 14, "bold"), wraplength=400)
search_record_label.grid(row=7, column=0, padx=5, pady=5, columnspan=3)

tree.bind('<ButtonRelease-1>', on_tree_click)
search_tree.bind('<ButtonRelease-1>', on_search_tree_click)

tree.bind('<Delete>', lambda e: on_delete_key_pressed(e, tree, search_tree))
search_tree.bind('<Delete>', lambda e: on_delete_key_pressed(e, tree, search_tree))

root.protocol("WM_DELETE_WINDOW", on_closing)

# Add this button to clear the dashboard
clear_button = tk.Button(center_frame, text="Clear", command=clear_dashboard)
clear_button.grid(row=2, column=3, padx=5, pady=5)

root.mainloop()