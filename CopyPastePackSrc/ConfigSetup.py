import tkinter as tk
from tkinter import ttk, messagebox
import re
import sqlite3
import sys
import os
import pandas as pd

# Get the directory where the script or executable is located
if getattr(sys, 'frozen', False):
    # If the application is frozen with PyInstaller, use this path
    application_path = os.path.dirname(sys.executable)
else:
    # Otherwise use the path to the script file
    application_path = os.path.dirname(os.path.abspath(__file__))

path_to_db = os.path.join(application_path, 'CopyPastePack.db')

class MatchingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Matching Setup Application")
        self.root.geometry("1200x800")  # Adjust size as needed
        
        self.selected_text = ""
        self.current_function = None
        self.highlightbuttonpressed = False
        self.current_direction_index = 0
        self.directions = []
        self.active_button = None
        self.current_step = 0
        self.max_steps = 3  # Default max steps
        
        self.setup_ui()
        self.load_upc_codes()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Modify step indicator setup
        self.step_frame = ttk.Frame(main_frame)
        self.step_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
        self.step_labels = []

        # for i in range(3):
        #     label = ttk.Label(self.step_frame, text=f"Step {i+1}", font=("Arial", 12))
        #     label.grid(row=0, column=i, padx=10, pady=5)
        #     self.step_labels.append(label)
        # self.update_step_indicator()
        
        # Left third: Directions
        directions_frame = ttk.Frame(main_frame, padding="10")
        directions_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.directions_text = tk.StringVar()
        self.directions_label = ttk.Label(directions_frame, textvariable=self.directions_text, wraplength=300, justify=tk.LEFT, font=("Arial", 14))
        self.directions_label.pack(fill=tk.BOTH, expand=True)
        
        # Middle third: Text widget
        self.text_widget = tk.Text(main_frame, wrap=tk.WORD, width=50, height=20, font=("Arial", 12))
        self.text_widget.grid(row=1, column=1, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.text_widget.bind('<Control-l>', self.on_selection)
        
        # Right third: Results
        results_frame = ttk.Frame(main_frame, padding="10")
        results_frame.grid(row=1, column=2, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.results_text = tk.StringVar()
        self.results_label = ttk.Label(results_frame, textvariable=self.results_text, wraplength=300, justify=tk.LEFT, font=("Arial", 14))
        self.results_label.pack(fill=tk.BOTH, expand=True)
        
        # Bottom: Option buttons
        options_frame = ttk.Frame(main_frame, padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        self.buttons = []
        button_configs = [
            ("Order Block Setup", self.get_block_keyword),
            ("Pair Setup", self.pair_setup),
            ("Quantity Per Item Setup", self.phrasequantity_setup),
            ("Incomplete/Split Phrase Setup", self.incompletephrase_setup),
            ("Exact phrase setup", self.exactphrase_setup)
        ]
        
        for i, (text, command) in enumerate(button_configs):
            btn_frame = tk.Frame(options_frame, borderwidth=2, relief='flat', highlightthickness=2, highlightbackground='#E6F3FF')
            btn_frame.grid(row=0, column=i, padx=5, pady=5)
            
            btn = ttk.Button(btn_frame, text=text, style='TButton')
            btn.pack(expand=True, fill='both')
            btn.configure(command=lambda b=btn_frame, c=command: self.on_button_click(b, c))
            
            self.buttons.append((btn_frame, btn))
        
        self.next_button = ttk.Button(options_frame, text="Next Step", command=self.next_step)
        self.next_button.grid(row=0, column=len(self.buttons), padx=5, pady=5)
        
        # Configure grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(1, weight=1)
        options_frame.columnconfigure(tuple(range(len(self.buttons) + 1)), weight=1)

    def set_active_function(self, function):
        if self.active_button:
            self.active_button.state(['!active'])
        
        self.current_function = function
        
        for btn, cmd in self.buttons:
            if cmd == function:
                self.active_button = btn
                self.active_button.state(['active'])
                break
        
        function()

    def update_step_indicator(self):
        # Clear existing labels
        for label in self.step_labels:
            label.destroy()
        self.step_labels.clear()

        # Create new labels
        for i in range(self.max_steps):
            label = ttk.Label(self.step_frame, text=f"Step {i+1}", font=("Arial", 12))
            label.grid(row=0, column=i, padx=10, pady=5)
            self.step_labels.append(label)

        # Update label styles
        for i, label in enumerate(self.step_labels):
            if i == self.current_step:
                label.configure(font=("Arial", 12, "bold"), foreground="blue")
            else:
                label.configure(font=("Arial", 12), foreground="black")

    def on_button_click(self, clicked_frame, function):
        if self.active_button:
            self.active_button[0].config(highlightbackground='#E6F3FF', highlightthickness=2)
        
        clicked_frame.config(highlightbackground='#4682B4', highlightthickness=2)
        self.active_button = (clicked_frame, function)
        self.current_function = function
        self.current_step = 0  # Reset to first step when new function is selected
        function()  # This will set self.max_steps and update the step indicator

    def next_step(self):
        if self.current_step == self.max_steps - 1:  # If we're on the last step
            messagebox.showinfo("End of Process", "You've reached the end of this matching process. Please select a new matching type or close the application.")
            return

        self.current_step = (self.current_step + 1) % self.max_steps
        self.update_step_indicator()
        self.current_direction_index = (self.current_direction_index + 1) % len(self.directions)
        self.update_directions(self.directions[self.current_direction_index])
        if self.current_function:
            self.current_function()

    def update_directions(self, new_text):
        self.directions_text.set(new_text)
        self.selected_text = ""
        self.results_text.set("")

    def on_selection(self, event=None):
        try:
            self.selected_text = self.text_widget.selection_get()
            self.highlightbuttonpressed = True
            print("Selected text committed:", self.selected_text)  # For debugging
            if self.current_function:
                self.current_function()
        except:
            print("Nothing was highlighted")


    def load_upc_codes(self):
        path_to_xlsx = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'UPCCodes.xlsx')
        df = pd.read_excel(path_to_xlsx, engine='openpyxl', header=None, dtype=str)
        df.columns = ["Item Name", "UPC Code"]
        df["Item Name"] = df["Item Name"].str.lower()
        df["UPC Code"] = df["UPC Code"].str.rstrip('.0')
        self.upc_codes = dict(zip(df["Item Name"], df["UPC Code"]))

    def process_matching(self, text_input, matching_type):
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()

        try:
            if matching_type == 'block_keyword':
                c.execute("SELECT * FROM blockseperator WHERE keywordstype = 'blockseperator'")
                row = c.fetchone()
                if row and len(row) >= 4:
                    blockseperatorkeyword = row[2].lower() if row[2] else ""
                    packagenumberphrase = row[3].lower() if row[3] else ""

                    if not blockseperatorkeyword or not packagenumberphrase:
                        return ["Incomplete block keyword configuration. Please set up both the block separator and package number phrase."]

                    blocks = re.split(f'({re.escape(blockseperatorkeyword)})', text_input.lower())
                    blocks_with_keyword = [blocks[i] + (blocks[i+1] if i+1 < len(blocks) else '') for i in range(0, len(blocks), 2)]

                    results = []
                    for block in blocks_with_keyword:
                        match = re.search(f"{re.escape(packagenumberphrase)} ([^\s]+)", block)
                        if match:
                            order_num = match.group(1)
                            results.append(f"Order Number: {order_num}")
                            for item in self.upc_codes:
                                if item in block:
                                    results.append(f"  Item: {item}, UPC: {self.upc_codes[item]}")
                    return results if results else ["No matches found with the current configuration."]
                else:
                    return ["Incomplete block keyword configuration. Please complete the setup."]

            elif matching_type == 'pair':
                c.execute("SELECT * FROM pairings WHERE keywordstype = 'pairing'")
                row = c.fetchone()
                if row and len(row) >= 4:
                    keywords = row[2].lower().split('<') if row[2] else []
                    removals = row[3].lower().split('<') if row[3] else []

                    if not keywords:
                        return ["No keywords defined in the pairing configuration. Please set up at least one keyword."]

                    results = []
                    for keyword in keywords:
                        if keyword in text_input.lower():
                            parts = text_input.lower().split(keyword)
                            if len(parts) > 1:
                                item = parts[1].split()[0]
                                for removal in removals:
                                    item = item.replace(removal, '')
                                if item in self.upc_codes:
                                    results.append(f"Item: {item}, UPC: {self.upc_codes[item]}")
                    return results if results else ["No matches found with the current configuration."]
                else:
                    return ["Incomplete pairing configuration. Please complete the setup."]

            elif matching_type == 'quantity':
                c.execute("SELECT * FROM quantitys WHERE keywordstype = 'Quantity'")
                rows = c.fetchall()
                if rows:
                    results = []
                    for row in rows:
                        if len(row) >= 4:
                            phrase = row[2].lower() if row[2] else ""
                            try:
                                position = int(row[3]) if row[3] else 0
                            except ValueError:
                                return [f"Invalid position value in quantity configuration: {row[3]}"]

                            if phrase and phrase in text_input.lower():
                                words = text_input.lower().split()
                                phrase_start = words.index(phrase.split()[0])
                                if phrase_start + position < len(words):
                                    quantity = words[phrase_start + position]
                                    item = ' '.join(words[phrase_start + position + 1:])
                                    if item in self.upc_codes:
                                        results.append(f"Quantity: {quantity}, Item: {item}, UPC: {self.upc_codes[item]}")
                    return results if results else ["No matches found with the current configuration."]
                else:
                    return ["No quantity configuration found. Please set up the quantity matching."]

            elif matching_type == 'incomplete_phrase':
                c.execute("SELECT * FROM incompletephrases WHERE keywordstype = 'Incomplete Phrase'")
                rows = c.fetchall()
                if rows:
                    results = []
                    for row in rows:
                        if len(row) >= 4:
                            phrase = row[2].lower() if row[2] else ""
                            secondary = row[3].lower() if row[3] else ""

                            if phrase and secondary and phrase in text_input.lower():
                                parts = text_input.lower().split(phrase)
                                if len(parts) > 1:
                                    item = parts[1].split()[0] + ' ' + secondary
                                    if item in self.upc_codes:
                                        results.append(f"Item: {item}, UPC: {self.upc_codes[item]}")
                    return results if results else ["No matches found with the current configuration."]
                else:
                    return ["No incomplete phrase configuration found. Please set up the incomplete phrase matching."]

            elif matching_type == 'exact_phrase':
                c.execute("SELECT * FROM exactphrases WHERE keywordstype = 'exactphrase'")
                rows = c.fetchall()
                if rows:
                    results = []
                    for row in rows:
                        if len(row) >= 4:
                            phrase = row[2].lower() if row[2] else ""
                            items = row[3].lower().split('<') if row[3] else []

                            if phrase and phrase in text_input.lower():
                                for item in items:
                                    if item in self.upc_codes:
                                        results.append(f"Item: {item}, UPC: {self.upc_codes[item]}")
                    return results if results else ["No matches found with the current configuration."]
                else:
                    return ["No exact phrase configuration found. Please set up the exact phrase matching."]

            else:
                return ["Unknown matching type."]

        except sqlite3.Error as e:
            return [f"An error occurred while accessing the database: {str(e)}"]
        except Exception as e:
            return [f"An unexpected error occurred: {str(e)}"]
        finally:
            conn.close()
    
    def display_matching_results(self, matching_type):
        text_input = self.text_widget.get("1.0", tk.END)
        results = self.process_matching(text_input, matching_type)
        result_text = "\n".join(results)
        self.results_text.set(result_text)


    def find_matching_starting_words(self, block1, block2):
        words1 = block1.split()
        words2 = block2.split()
        keyword = ""
        maxindex = min(len(words1), len(words2))
        
        for i in range(maxindex):
            if words1[i] == words2[i]:
                keyword += words1[i] + " "
            else:
                break
        
        return keyword.strip()

    def get_block_keyword(self):
        self.max_steps = 6  # Set max steps for this function
        self.update_step_indicator()
        
        self.directions = [
            "Step 1: Highlight one order as a block of text including any important words before or after and press Ctrl+L to commit it",
            "Step 2: Highlight a different order as a block of text and press Ctrl+L to commit it",
            "Step 3: Check if blocks are correct on the right side of your screen and continue to the next step for package number setup",
            "Step 4: Highlight phrase where package number will appear right after and press Ctrl+L to commit it",
            "Step 5: Highlight package number phrase including the package number and press Ctrl+L to commit it",
            "Step 6: Check the database and make sure things look right"
        ]
        self.current_direction_index = 0
        self.update_directions(self.directions[self.current_direction_index])
        
        block1 = ""
        block2 = ""
        packagenumberphrase = ""
        blockseperator = ""
        
        def process_step():
            nonlocal block1, block2, packagenumberphrase, blockseperator
            
            if self.current_direction_index == 0 and self.selected_text:
                block1 = self.selected_text
                self.results_text.set("Order 1: \n" + block1)
            elif self.current_direction_index == 1 and self.selected_text:
                block2 = self.selected_text
                self.results_text.set("Order 2: \n" + block2)
            elif self.current_direction_index == 2 and self.highlightbuttonpressed:
                blockseperator = self.find_matching_starting_words(block1, block2)
                self.results_text.set("Keyword: \n" + blockseperator)
            elif self.current_direction_index == 3 and self.highlightbuttonpressed:
                packagenumberphrase = self.selected_text
                self.results_text.set("Package Phrase: " + packagenumberphrase)
            elif self.current_direction_index == 4 and self.highlightbuttonpressed:
                match = re.search(f"{re.escape(packagenumberphrase)} (\w+)", self.selected_text)
                if match:
                    self.results_text.set("Package Number: " + match.group(1))
            elif self.current_direction_index == 5:
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                # Create table if it doesn't exist
                c.execute('''CREATE TABLE IF NOT EXISTS blockseperator
                             (id INTEGER PRIMARY KEY, keywordstype TEXT, blockseperator TEXT, packagenumberphrase TEXT)''')
                
                keyword_type = 'blockseperator'
                c.execute("INSERT INTO blockseperator (keywordstype, blockseperator, packagenumberphrase) VALUES (?, ?, ?)",
                          (keyword_type, blockseperator, packagenumberphrase))
                conn.commit()
                conn.close()
                self.display_matching_results('block_keyword')
            
            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def pair_setup(self):
        self.max_steps = 3  # Set max steps for this function
        self.update_step_indicator()
        
        self.directions = [
            "Step 1: Highlight Each Keyword/Identifier in Each Phrase. The order in which you highlighted should match way the item would be read. Press Ctrl+L after each highlight.",
            "Step 2: Highlight the Data You DO NOT Care About in Each Phrase. Any data that isn't a keyword or removal will be considered part of the item. Press Ctrl+L after each highlight.",
            "Step 3: Check results and press 'Next Step' to save to database."
        ]
        
        keywordslist = []
        unwantedlist = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text not in keywordslist:
                keywordslist.append(self.selected_text)
                self.results_text.set("\n".join(f"Keyword {i+1}: {kw}" for i, kw in enumerate(keywordslist)))
            elif self.current_direction_index == 1 and self.selected_text and self.selected_text not in unwantedlist:
                unwantedlist.append(self.selected_text)
                self.results_text.set("\n".join(f"Unwanted {i+1}: {uw}" for i, uw in enumerate(unwantedlist)))
            elif self.current_direction_index == 2:
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                # Create table if it doesn't exist
                c.execute('''CREATE TABLE IF NOT EXISTS pairings
                             (id INTEGER PRIMARY KEY, keywordstype TEXT, keywords TEXT, removals TEXT)''')
                
                keyword_type = 'pairing'
                keywords = '<'.join(keywordslist)
                removals = '<'.join(unwantedlist)
                c.execute("INSERT INTO pairings (keywordstype, keywords, removals) VALUES (?, ?, ?)",
                          (keyword_type, keywords, removals))
                conn.commit()
                conn.close()
                self.display_matching_results('pair')

            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def phrasequantity_setup(self):

        self.max_steps = 3  # Set max steps for this function
        self.update_step_indicator()

        self.directions = [
            "Step 1: Highlight Each full phrase containing the quantity for an item. If you plan to do multiple at a time, maintain the order for the next step. Press Ctrl+L after each highlight.",
            "Step 2: Highlight the quantity in each phrase. Again if doing multiple, highlight the first phrase's quantity first. Press Ctrl+L after each highlight.",
            "Step 3: Check results and press 'Next Step' to save to database."
        ]
        self.current_direction_index = 0
        self.update_directions(self.directions[self.current_direction_index])
        
        phraseslist = []
        data = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text not in phraseslist:
                phraseslist.append(self.selected_text)
                self.results_text.set("\n".join(f"Phrase {i+1}: {p}" for i, p in enumerate(phraseslist)))
            elif self.current_direction_index == 1 and self.highlightbuttonpressed:
                data.append(self.selected_text)
                self.results_text.set("\n".join(f"QuantityPiece {i+1}: {d}" for i, d in enumerate(data)))
            elif self.current_direction_index == 2:
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                # Create table if it doesn't exist
                c.execute('''CREATE TABLE IF NOT EXISTS quantitys
                             (id INTEGER PRIMARY KEY, keywordstype TEXT, phrases TEXT, positions TEXT)''')
                
                keyword_type = 'Quantity'
                
                for i, phrase in enumerate(phraseslist):
                    try:
                        dat = phrase.split().index(data[i])
                        words = phrase.split()
                        del words[dat]
                        phrase = " ".join(words)
                        c.execute("INSERT INTO quantitys (keywordstype, phrases, positions) VALUES (?, ?, ?)",
                                  (keyword_type, phrase, str(dat)))
                    except:
                        pass
                
                conn.commit()
                conn.close()
                self.display_matching_results('quantity')

            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def incompletephrase_setup(self):
        self.max_steps = 3  # Set max steps for this function
        self.update_step_indicator()

        self.directions = [
            "Step 1: Highlight Each Incomplete Phrase. Maintain order for the next step. Press Ctrl+L after each highlight.",
            "Step 2: Highlight the word that will be appended to the item describers below. Press Ctrl+L after each highlight.",
            "Step 3: Check results and press 'Next Step' to save to database."
        ]
        self.current_direction_index = 0
        self.update_directions(self.directions[self.current_direction_index])
        
        phraseslist = []
        data = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text not in phraseslist:
                phraseslist.append(self.selected_text)
                self.results_text.set("\n".join(f"Phrase {i+1}: {p}" for i, p in enumerate(phraseslist)))
            elif self.current_direction_index == 1 and self.highlightbuttonpressed:
                data.append(self.selected_text)
                self.results_text.set("\n".join(f"SecondaryPiece {i+1}: {d}" for i, d in enumerate(data)))
            elif self.current_direction_index == 2:
                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                # Create table if it doesn't exist
                c.execute('''CREATE TABLE IF NOT EXISTS incompletephrases
                             (id INTEGER PRIMARY KEY, keywordstype TEXT, phrases TEXT, secondarykeywords TEXT)''')
                
                keyword_type = 'Incomplete Phrase'
                
                while phraseslist and data:
                    phrase = phraseslist.pop(0)
                    dat = data.pop(0)
                    c.execute("INSERT INTO incompletephrases (keywordstype, phrases, secondarykeywords) VALUES (?, ?, ?)",
                              (keyword_type, phrase, dat))
                
                conn.commit()
                conn.close()
                self.display_matching_results('incomplete_phrase')

            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def exactphrase_setup(self):
            self.max_steps = 3  # Set max steps for this function
            self.update_step_indicator()

            self.directions = [
                "Step 1: Highlight One Exact Order Phrase. Make sure to EXCLUDE data that can change (Like Quantity). Press Ctrl+L to commit it.",
                "Step 2: Highlight each item to be added when this phrase is seen. These items should match your UPC items and the website item format. Press Ctrl+L after each highlight.",
                "Step 3: Check results and press 'Next Step' to save to database."
            ]
            self.current_direction_index = 0
            self.update_directions(self.directions[self.current_direction_index])
            
            exactphrase = ""
            itemslist = []
            
            def process_step():
                nonlocal exactphrase
                
                if self.current_direction_index == 0 and self.selected_text and self.selected_text != exactphrase:
                    exactphrase = self.selected_text
                    self.results_text.set("Phrase: " + exactphrase)
                elif self.current_direction_index == 1 and self.highlightbuttonpressed:
                    itemslist.append(self.selected_text)
                    self.results_text.set("\n".join(f"Item {i+1}: {item}" for i, item in enumerate(itemslist)))
                elif self.current_direction_index == 2:
                    conn = sqlite3.connect(path_to_db)
                    c = conn.cursor()
                    
                    # Create table if it doesn't exist
                    c.execute('''CREATE TABLE IF NOT EXISTS exactphrases
                                (id INTEGER PRIMARY KEY, keywordstype TEXT, exactphrases TEXT, items TEXT)''')
                    
                    keyword_type = 'exactphrase'
                    c.execute("INSERT INTO exactphrases (keywordstype, exactphrases, items) VALUES (?, ?, ?)",
                            (keyword_type, exactphrase, '<'.join(itemslist)))
                    conn.commit()
                    conn.close()
                self.display_matching_results('exact_phrase')

                self.highlightbuttonpressed = False
            
            self.current_function = process_step

def main():
    root = tk.Tk()
    app = MatchingApp(root)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

if __name__ == "__main__":
    main()
