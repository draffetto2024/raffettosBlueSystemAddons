import tkinter as tk
from tkinter import ttk
import re
import sqlite3
import sys
import os

# Get the directory where the script or executable is located
if getattr(sys, 'frozen', False):
    # If the application is frozen with PyInstaller, use this path
    application_path = os.path.dirname(sys.executable)
else:
    # Otherwise use the path to the script file
    application_path = os.path.dirname(os.path.abspath(__file__))

path_to_db = os.path.join(application_path, 'CopyPastePack.db')

#the curly version of an apostrophe can cause issues with matching, fixed for now only in phrasequantity

directions = [
    "Step 1: Select and highlight the desired text.",
    "Step 2: Click the 'Print Highlighted Text' button to print the highlighted text.",
    "Step 3: Use the 'Next Step' and 'Previous Step' buttons to navigate through the directions."
]

selected_text = ""

current_direction_index = 0

highlightbuttonpressed = False

def on_close():
    # Perform any cleanup operations here
    # For example, saving data, closing files, etc.
    root.destroy()

def on_selection():
    global selected_text
    global highlightbuttonpressed
    try:
        selected_text = text_widget.selection_get()
        highlightbuttonpressed = True  # Button has been pressed
    except:
        print("Nothing was highlighted")

def update_directions(new_text):
    directions_text.set(new_text)
    global selected_text
    selected_text = ""
    results_text.set("")

def next_step():
    global current_direction_index
    current_direction_index = (current_direction_index + 1) % len(directions)
    update_directions(directions[current_direction_index])

def previous_step():
    global current_direction_index
    current_direction_index = (current_direction_index - 1) % len(directions)
    update_directions(directions[current_direction_index])
    
def find_matching_starting_words(block1, block2):
    # Split the strings by spaces
    words1 = block1.split()
    words2 = block2.split()
    
    index = 0
    
    keyword = ""
    
    if len(words1) < len(words2):
        maxindex = len(words1)
    else:
        maxindex = len(words2)
    
    for i in range(0, maxindex):
        if words1[i] == words2[i]:
            keyword = keyword + words1[i] + " ";
        else:
            break
        
    return keyword


def return_to_main_menu():
    next_button.pack_forget()
    previous_button.pack_forget()
    button.pack_forget()
    next_button.pack_forget()
    directions_label.pack_forget()
    results_label.pack_forget()
    text_widget.pack_forget()
    return_to_main_menu_button.pack_forget()
    
    get_block_keyword_button.pack()
    pair_setup_button.pack()
    phrasequantity_setup_button.pack()
    incompletephrase_setup_button.pack()
    exactphrase_setup_button.pack()

def selection_presetup():
    directions_label.pack(fill='x', padx=10, pady=10)
    results_label.pack(side = 'right', fill='x', padx = 200, pady = 10)
    text_widget.pack()
    
    button.pack(side=tk.LEFT, anchor=tk.S)
    next_button.pack(side=tk.LEFT, anchor=tk.S)
    previous_button.pack(side=tk.LEFT, anchor=tk.S)
    return_to_main_menu_button.pack(side=tk.LEFT, anchor=tk.S)
    
    #Menu buttons should be deleted
    get_block_keyword_button.pack_forget()
    pair_setup_button.pack_forget()
    phrasequantity_setup_button.pack_forget()
    incompletephrase_setup_button.pack_forget()
    exactphrase_setup_button.pack_forget()
    
def exactphrase_setup():
    selection_presetup()
    global highlightbuttonpressed
    
    global directions
    directions  = [
        "Step 1: Highlight One Exact Order Phrase. Make sure to EXCLUDE data that can change (Like Quantity)",
        "Step 2: Highlight each item to be added when this phrase is seen. These items should match the your UPC items and the website item format",
        "Step 3: Check results"
        ]
    global current_direction_index
    current_direction_index = 0
    update_directions(directions[current_direction_index])
    
    global selected_text
    
    exactphrase = ""
    itemslist = []
    phrasesstring = ""
    itemsstring = ""
    goingtofinalstep = False
    
    while True:
        if current_direction_index == 0:
            if selected_text and selected_text != exactphrase:
                exactphrase = selected_text
                phrasesstring = "Phrase: " + exactphrase + "\n"
                results_text.set(phrasesstring)
                
        if current_direction_index == 1:
            goingtofinalstep = True
            if selected_text and highlightbuttonpressed:
                itemslist.append(selected_text)
                itemsstring = ""
                for i, p in enumerate(itemslist, 1):  # Start numbering from 1
                    itemsstring += "Item " + str(i) + ": " + p + "\n"  # Add the number and the string to itemsstring
                results_text.set(itemsstring)
                highlightbuttonpressed = False
                
        if current_direction_index == 2:
            if goingtofinalstep:
                # Connect to the SQLite database (it will be created if it doesn't exist)
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                # Your keyword type, keywords list, and removals list
                keyword_type = 'exactphrase'
                
                # Insert a record into the pairings table
                c.execute("INSERT INTO exactphrases (keywordstype, exactphrases, items) VALUES (?, ?, ?)",
                          (keyword_type, exactphrase, '<'.join(itemslist)))
                
                
                # Commit the changes and close the connection
                conn.commit()
                conn.close()  
                goingtofinalstep = False
                
        if root.winfo_exists():
            root.update()

def incompletephrase_setup():
    selection_presetup()
    global highlightbuttonpressed
    
    global directions
    directions  = [
        "Step 1: Highlight Each Incomplete Phrase. Maintain order for the next step",
        "Step 2: Highlight the word that will be appended to the item describers below",
        "Step 3: Check results"
        ]
    global current_direction_index
    current_direction_index = 0
    update_directions(directions[current_direction_index])
    
    global selected_text
    
    phraseslist = []
    phrasesstring = ""
    data = []
    datastring = ""
    
    while True:
        if(current_direction_index == 0):
            if selected_text and selected_text not in phraseslist:
                phraseslist.append(selected_text)
                phrasesstring = ""
                for i, p in enumerate(phraseslist, 1):  # Start numbering from 1
                    phrasesstring += "Phrase " + str(i) + ": " + p + "\n"  # Add the number and the string to keywordsstring
                results_text.set(phrasesstring)
        if(current_direction_index == 1):
           goingtofinalstep = True
           if selected_text and highlightbuttonpressed:
               data.append(selected_text)
               datastring = ""
               for i, p in enumerate(data, 1):  # Start numbering from 1
                   datastring += "SecondaryPiece " + str(i) + ": " + p + "\n"  # Add the number and the string to keywordsstring
               results_text.set(datastring)
               highlightbuttonpressed = False  # Reset the flag
        if(current_direction_index == 2):
            if(goingtofinalstep):
                # Connect to the SQLite database (it will be created if it doesn't exist)
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()              
                
                # Your keyword type, keywords list, and removals list
                keyword_type = 'Incomplete Phrase'
                
                while phraseslist and data:
                    phrase = phraseslist.pop(0)  # Take the first item from phraseslist
                    dat = data.pop(0)  # Take the first item from datalocations
                    
                
                    # Insert a record into the pairings table
                    c.execute("INSERT INTO incompletephrases (keywordstype, phrases, secondarykeywords) VALUES (?, ?, ?)",
                              (keyword_type, phrase, dat))
                
                # Commit the changes and close the connection
                conn.commit()
                conn.close()  
                goingtofinalstep = False
            
        if root.winfo_exists():
                root.update()

def phrasequantity_setup():
    selection_presetup()
    global highlightbuttonpressed
    
    global directions
    directions  = [
        "Step 1: Highlight Each full phrase containing the quantity for an item. If you plan to do multiple at a time, maintain the order for the next step",
        "Step 2: Highlight the quantity in each phrase. Again if doing multiple, highlight the first phrase's quantity first",
        "Step 3: Check results"
        ]
    global current_direction_index
    current_direction_index = 0
    update_directions(directions[current_direction_index])
    
    global selected_text
    
    phraseslist = []
    phrasesstring = ""
    data = []
    datalocations = []
    datastring = ""
    goingtofinalstep = False
    
    while True:
        if(current_direction_index == 0):
            if selected_text and selected_text not in phraseslist:
                phraseslist.append(selected_text)
                phrasesstring = ""
                for i, p in enumerate(phraseslist, 1):  # Start numbering from 1
                    phrasesstring += "Phrase " + str(i) + ": " + p + "\n"  # Add the number and the string to keywordsstring
                results_text.set(phrasesstring)
        if(current_direction_index == 1):
           goingtofinalstep = True
           if selected_text and highlightbuttonpressed:
               data.append(selected_text)
               datastring = ""
               for i, p in enumerate(data, 1):  # Start numbering from 1
                   datastring += "QuantityPiece " + str(i) + ": " + p + "\n"  # Add the number and the string to keywordsstring
               results_text.set(datastring)
               highlightbuttonpressed = False  # Reset the flag
        if(current_direction_index == 2):
            if(goingtofinalstep):
                # Connect to the SQLite database (it will be created if it doesn't exist)
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                for i, phrase in enumerate(phraseslist):
                    try:
                        datalocations.append(phrase.split().index(data[i]))
                    except:
                        pass
                    
                
                # Your keyword type, keywords list, and removals list
                keyword_type = 'Quantity'
                
                while phraseslist and datalocations:
                    phrase = phraseslist.pop(0)  # Take the first item from phraseslist
                    dat = datalocations.pop(0)  # Take the first item from datalocations
                    
                    words = phrase.split()  # Splits the string into a list of words.
                    del words[dat]  # Removes the word at the specific index.
                    phrase = " ".join(words)  # Joins the words back into a string.
                    
                
                    # Insert a record into the pairings table
                    c.execute("INSERT INTO quantitys (keywordstype, phrases, positions) VALUES (?, ?, ?)",
                              (keyword_type, phrase, str(dat)))
                
                # Commit the changes and close the connection
                conn.commit()
                conn.close()  
                goingtofinalstep = False
            
        if root.winfo_exists():
                root.update()
    
    
def pair_setup():
    selection_presetup()
    
    global directions
    directions  = [
        "Step 1: Highlight Each Keyword/Identifier in Each Phrase. The order in which you highlighted should match way the item would be read",
        "Step 2: Highlight the Data You DO NOT Care About in Each Phrase. Any data that isn't a keyword or removal will be considered part of the item",
        "Step 3: Highlight An Entire Order Containing the Pairing"
        ]
    global current_direction_index
    current_direction_index = 0
    update_directions(directions[current_direction_index])
    
    global selected_text
    
    phraseslist = []
    keywordslist = []
    unwantedlist = []
    phrasesstring = ""
    keywordsstring = ""
    unwantedstring = ""
    goingtofinalstep = False
    while True:
        if(current_direction_index == 0):
            if selected_text and selected_text not in keywordslist:
                keywordslist.append(selected_text)
                keywordsstring = ""
                for i, p in enumerate(keywordslist, 1):  # Start numbering from 1
                    keywordsstring += "Keyword " + str(i) + ": " + p + "\n"  # Add the number and the string to keywordsstring
                results_text.set(keywordsstring)
        if(current_direction_index == 1):
           goingtofinalstep = True
           if selected_text and selected_text not in unwantedlist:
               unwantedlist.append(selected_text)
               unwantedstring = ""
               for i, p in enumerate(unwantedlist, 1):  # Start numbering from 1
                   unwantedstring += "Unwanted " + str(i) + ": " + p + "\n"  # Add the number and the string to keywordsstring
               results_text.set(unwantedstring)
        if(current_direction_index == 2):
            if(goingtofinalstep):
                # Connect to the SQLite database (it will be created if it doesn't exist)
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                # Your keyword type, keywords list, and removals list
                keyword_type = 'pairing'
                
                # Convert lists into strings with < symbol as the separator
                keywords = '<'.join(keywordslist)
                removals = '<'.join(unwantedlist)
                
                # Insert a record into the pairings table
                c.execute("INSERT INTO pairings (keywordstype, keywords, removals) VALUES (?, ?, ?)",
                          (keyword_type, keywords, removals))
                
                # Commit the changes and close the connection
                conn.commit()
                conn.close()  
                goingtofinalstep = False
            
        if root.winfo_exists():
                root.update()
    
    
def get_block_keyword():
    selection_presetup()
    
    global directions
    directions  = [
        "Step 1: Highlight one order as a block of text including any important words before or after and the hit the commit button to verify it",
        "Step 2: Highlight a different order as a block of text",
        "Step 3: Check if blocks are correct on the right side of your screen and continue to the next step for package number setup",
        "Step 4: Highlight phrase where package number will appear right after",
        "Step 5: Highlight package number phrase including theh package number and check that the package number is correct",
        "Step 6: Check the database and make sure things look right"
        ]
    global current_direction_index
    current_direction_index = 0
    update_directions(directions[current_direction_index])
    get_block_keyword_button.pack_forget()
    
    block1 = ""
    block2 = ""
    packagenumberphrase = ""
    
    global selected_text
    global highlightbuttonpressed
    goingtofinalstep = False
    
    while True:
        if(current_direction_index == 0): #Step 1
            if selected_text:
                block1 = selected_text
                selected_text = ""
                results_text.set("Order 1: \n" + block1)
        if(current_direction_index == 1):
            if selected_text:
                block2 = selected_text
                selected_text = ""
                results_text.set("Order 2: \n" + block2)
                goingtofinalstep = True
        if current_direction_index == 2 and highlightbuttonpressed:
                blockseperator = find_matching_starting_words(block1, block2)
                results_text.set("Keyword: \n" + blockseperator)
                
        if(current_direction_index == 3 and highlightbuttonpressed):
            packagenumberphrase = selected_text
            results_text.set("Package Phrase: " + packagenumberphrase)
            select_text = ""
        if(current_direction_index == 4 and highlightbuttonpressed):
            
            # Use a regular expression to find the word immediately following the phrase
            match = re.search(f"{re.escape(packagenumberphrase)} (\w+)", selected_text)
            try:
                results_text.set("Package Number: " + match.group(1))
            except: 
                pass
        
            goingtofinalstep = True
        
        if (current_direction_index == 5):
            if goingtofinalstep:
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                # Your keyword type, keywords list, and removals list
                keyword_type = 'blockseperator'
                
                # Insert a record into the pairings table
                c.execute("INSERT INTO blockseperator (keywordstype, blockseperator, packagenumberphrase) VALUES (?, ?, ?)",
                          (keyword_type, blockseperator, packagenumberphrase))
                
                # Commit the changes and close the connection
                conn.commit()
                conn.close()  
                goingtofinalstep = False
        if root.winfo_exists():
                root.update()
        
    
root = tk.Tk()

directions_text = tk.StringVar()
directions_label = tk.Label(root, textvariable=directions_text, justify='left', anchor='w', wraplength=400, font=("Arial", 20))

results_text = tk.StringVar()
results_label = tk.Label(root, textvariable=results_text, justify= 'right', wraplength = 400, font=("Arial", 20))

text_widget = tk.Text(root)

button = tk.Button(root, text="Commit Highlighted Text", command=on_selection)

next_button = tk.Button(root, text="Next Step", command=next_step)

previous_button = tk.Button(root, text="Previous Step", command=previous_step)

get_block_keyword_button = tk.Button(root, text="Order Block Setup", command=get_block_keyword)
get_block_keyword_button.pack()

pair_setup_button = tk.Button(root, text="Pair Setup", command=pair_setup)
pair_setup_button.pack()

phrasequantity_setup_button = tk.Button(root, text="Quantity Per Item Setup", command=phrasequantity_setup)
phrasequantity_setup_button.pack()

incompletephrase_setup_button = tk.Button(root, text="Incomplete/Split Phrase Setup", command= incompletephrase_setup)
incompletephrase_setup_button.pack()

exactphrase_setup_button = tk.Button(root, text="Exact phrase setup", command= exactphrase_setup)
exactphrase_setup_button.pack()

return_to_main_menu_button = tk.Button(root, text="Return to Main Menu", command=return_to_main_menu)


update_directions(directions[current_direction_index])

root.protocol("WM_DELETE_WINDOW", on_close)

root.mainloop()
