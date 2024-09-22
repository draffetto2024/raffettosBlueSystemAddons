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

        self.last_order_dict = {}  # New attribute to store the last matched order
        self.results_text = tk.StringVar()
        
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
        self.load_database_configurations()

        # Add these lines to bind Ctrl+L and Ctrl+N
        self.root.bind('<Control-l>', self.on_selection)
        self.root.bind('<Control-n>', self.on_next_step)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Modify step indicator setup
        self.step_frame = ttk.Frame(main_frame)
        self.step_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
        self.step_labels = []
        
        # Left third: Directions
        directions_frame = ttk.Frame(main_frame, padding="10")
        directions_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.directions_text = tk.StringVar()
        self.directions_label = ttk.Label(directions_frame, textvariable=self.directions_text, wraplength=300, justify=tk.LEFT, font=("Arial", 14))
        self.directions_label.pack(fill=tk.BOTH, expand=True)
        
        # Middle third: Text widget (single text box for orders)
        middle_frame = ttk.Frame(main_frame, padding="10")
        middle_frame.grid(row=1, column=1, sticky=(tk.N, tk.S, tk.W, tk.E))
        
        # Add header for the middle entry box
        entry_header = ttk.Label(middle_frame, text="Entry box for orders", font=("Arial", 12, "bold"))
        entry_header.pack(pady=(0, 5))
        
        self.text_widget = tk.Text(middle_frame, wrap=tk.WORD, width=50, height=20, font=("Arial", 12))
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        # self.text_widget.bind('<Control-l>', self.on_selection)
        # self.root.bind('<Control-n>', self.next_step)
        
        # Right third: Results (using a Text widget)
        results_frame = ttk.Frame(main_frame, padding="10")
        results_frame.grid(row=1, column=2, sticky=(tk.N, tk.S, tk.W, tk.E))
        
        # Add header for the output display
        output_header = ttk.Label(results_frame, text="Output display", font=("Arial", 12, "bold"))
        output_header.pack(pady=(0, 5))
        
        self.results_text = tk.Text(results_frame, wrap=tk.WORD, width=50, height=20, font=("Arial", 12), state='disabled')
        self.results_text.pack(fill=tk.BOTH, expand=True)

        # Configure tags for coloring
        self.results_text.tag_configure("normal", foreground="black")
        self.results_text.tag_configure("blue", foreground="blue")
        
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
        
        # Modify the Next Step button to use the on_next_step method
        self.next_button = ttk.Button(options_frame, text="Next Step", command=self.on_next_step)
        self.next_button.grid(row=0, column=len(self.buttons), padx=5, pady=5)
        
        # Add Full Matching button
        self.full_matching_button = ttk.Button(options_frame, text="Full Matching", command=self.perform_full_matching)
        self.full_matching_button.grid(row=1, column=len(self.buttons), padx=5, pady=5)
        
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
        
        # Clear previous results
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.config(state='disabled')
                
        # Force update of the UI
        self.root.update_idletasks()
        
        function()  # This will set self.max_steps and update the step indicator
        
        # Update step indicator
        self.update_step_indicator()
        
        # Reset and update directions
        self.current_direction_index = 0
        self.update_directions(self.directions[self.current_direction_index])

    def on_next_step(self, event=None):
        # This method will be called by both the Next Step button and Ctrl+N
        self.next_step()
        return "break"  # Prevent the event from propagating

    def next_step(self):
        if self.current_step == self.max_steps - 1:  # If we're on the last step
            messagebox.showinfo("End of Process", "You've reached the end of this matching process. Please select a new matching type or close the application.")
            return

        self.current_step = (self.current_step + 1) % self.max_steps
        self.update_step_indicator()
        if len(self.directions) > 0:
            self.current_direction_index = (self.current_direction_index + 1) % len(self.directions)
            self.update_directions(self.directions[self.current_direction_index])
            if self.current_function:
                self.current_function()
        else:
            messagebox.showinfo("No Mode Selected", "You must select a mode before using the next step button")
            return

    def update_directions(self, new_text):
        self.directions_text.set(new_text)
        self.selected_text = ""
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.config(state='disabled')

    def on_selection(self, event=None):
        try:
            self.selected_text = self.text_widget.selection_get()
            print(f"Selected text: '{self.selected_text}'")
            
            if self.current_function:
                print(f"Current function: {self.current_function.__name__}")
                print(f"Current direction index: {self.current_direction_index}")
                self.highlightbuttonpressed = True
                self.current_function()
            else:
                print("No current function set.")
            
        except tk.TclError:
            print("No text currently selected.")
        except Exception as e:
            print(f"An error occurred: {str(e)}")
        finally:
            # Reset the selection to avoid processing the same selection multiple times
            self.text_widget.tag_remove(tk.SEL, "1.0", tk.END)


    def load_upc_codes(self):
        path_to_xlsx = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'UPCCodes.xlsx')
        df = pd.read_excel(path_to_xlsx, engine='openpyxl', header=None, dtype=str)
        df.columns = ["Item Name", "UPC Code"]
        df["Item Name"] = df["Item Name"].str.lower()
        df["UPC Code"] = df["UPC Code"].str.rstrip('.0')
        # Remove any rows with NaN values
        df = df.dropna()
        self.upc_codes = dict(zip(df["Item Name"], df["UPC Code"]))
        print(f"Loaded {len(self.upc_codes)} UPC codes")
        print("Sample UPC codes:")
        for item, upc in list(self.upc_codes.items())[:5]:
            print(f"  {item}: {upc}")

    def load_database_configurations(self):
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()

        # Load block separator configuration
        c.execute("SELECT DISTINCT * FROM blockseperator WHERE keywordstype = 'blockseperator'")
        row = c.fetchone()
        if row and len(row) >= 3:
            self.blockseperatorkeyword = row[2].lower() if row[2] else ""
            self.packagenumberphrase = row[3].lower() if len(row) > 3 and row[3] else ""
        else:
            self.blockseperatorkeyword = ""
            self.packagenumberphrase = ""
        
        print(f"Block separator keyword: '{self.blockseperatorkeyword}'")
        print(f"Package number phrase: '{self.packagenumberphrase}'")

        # Load pairing configuration
        c.execute("SELECT DISTINCT * FROM pairings WHERE keywordstype = 'pairing'")
        self.keywords_2d_list = []
        self.removals_list = []
        for row in c.fetchall():
            if len(row) >= 2:
                keywords_string = row[1].lower() if row[1] else ""
                removals_string = row[2].lower() if len(row) > 2 and row[2] else ""
                
                self.keywords_2d_list.append(keywords_string.split('<'))
                self.removals_list.extend(removals_string.split('<'))

        # Load quantity configuration
        c.execute("SELECT DISTINCT * FROM quantitys WHERE keywordstype = 'Quantity'")
        self.quantityphrases = []
        self.quantitypositions = []
        for row in c.fetchall():
            if len(row) >= 2:
                self.quantityphrases.append(row[1].lower() if row[1] else "")
                self.quantitypositions.append(int(row[2]) if len(row) > 2 and row[2] and row[2].isdigit() else 0)

        # Load incomplete phrase configuration
        c.execute("SELECT DISTINCT * FROM incompletephrases WHERE keywordstype = 'Incomplete Phrase'")
        self.incompletephrases = []
        self.secondarykeywords = []
        for row in c.fetchall():
            if len(row) >= 2:
                self.incompletephrases.append(row[1].lower().strip() if row[1] else "")
                self.secondarykeywords.append(row[2].lower().strip() if len(row) > 2 and row[2] else "")

        # Load exact phrase configuration
        c.execute("SELECT DISTINCT * FROM exactphrases WHERE keywordstype = 'exactphrase'")
        self.exactphrases = []
        self.exactphraseitems_2d = []
        for row in c.fetchall():
            if len(row) >= 2:
                phrases = row[1].lower().split('<') if row[1] else []
                items = row[2].lower().split('<') if len(row) > 2 and row[2] else []
                self.exactphrases.extend(phrases)
                self.exactphraseitems_2d.extend([items] * len(phrases))

        # Ensure exactphrases and exactphraseitems_2d have the same length
        min_length = min(len(self.exactphrases), len(self.exactphraseitems_2d))
        self.exactphrases = self.exactphrases[:min_length]
        self.exactphraseitems_2d = self.exactphraseitems_2d[:min_length]

        print("Loaded configurations:")
        print(f"  Keywords: {self.keywords_2d_list}")
        print(f"  Removals: {self.removals_list}")
        print(f"  Quantity phrases: {self.quantityphrases}")
        print(f"  Quantity positions: {self.quantitypositions}")
        print(f"  Incomplete phrases: {self.incompletephrases}")
        print(f"  Secondary keywords: {self.secondarykeywords}")
        print(f"  Exact phrases: {self.exactphrases}")
        print(f"  Exact phrase items: {self.exactphraseitems_2d}")

        conn.close()
        print("Database configurations loaded successfully.")


    def perform_full_matching(self):
        print("Starting perform_full_matching")
        text_input = self.text_widget.get("1.0", tk.END).lower().strip()
        print(f"Input text: {text_input}")
        order_dict = {}
        
        blocks = re.split(f'({re.escape(self.blockseperatorkeyword)})', text_input)
        blocks_with_keyword = [blocks[i] + (blocks[i + 1] if i + 1 < len(blocks) else '') for i in range(0, len(blocks), 2)]
        
        print(f"Number of blocks: {len(blocks_with_keyword)}")
        
        for block in blocks_with_keyword:
            order_num = ""
            phrases = block.split("\n")
            modified_phrases = []  # list to store modified phrases
            phrases_length = len(phrases)
            phrasequantity = 1
            
            # Extract order number
            for phrase in phrases:
                match = re.search(f"{re.escape(self.packagenumberphrase)} ([^\s]+)", phrase)
                if match:
                    order_num = ''.join([char for char in match.group(1) if char.isdigit()])
                    print(f"Found order number: {order_num}")
                    break
            
            if not order_num:
                print("Warning: No order number found for this block. Skipping...")
                continue

            splitting_keywords_list = [None] * len(self.keywords_2d_list)
            keyword_positions = dict()

            for i, line in enumerate(block.split("\n")):
                for j, keywords_list in enumerate(self.keywords_2d_list):
                    for keyword in keywords_list:
                        keyword_position = line.find(keyword)
                        if keyword_position != -1:  # keyword is in line
                            current_position = (i, keyword_position)
                            if keyword not in keyword_positions or current_position < keyword_positions[keyword]:
                                keyword_positions[keyword] = current_position

            print("keywords_2d_list", self.keywords_2d_list)
            print("removals_list", self.removals_list)
            
            # Now, find the keyword with the lowest index
            for keyword, position in keyword_positions.items():
                for i, keywords_list in enumerate(self.keywords_2d_list):
                    if keyword in keywords_list:
                        if splitting_keywords_list[i] is None or position < keyword_positions[splitting_keywords_list[i]]:
                            splitting_keywords_list[i] = keyword
            
            # Pairing Logic
            stringtobeadded = ""
            
            for i, keywords_list in enumerate(self.keywords_2d_list):
                splittingkeyword = splitting_keywords_list[i]
                if splittingkeyword and splittingkeyword in block:  # if the first keyword is in the block
                    sub_blocks = block.split(splittingkeyword)
                    sub_blocks = [splittingkeyword + sub_block for sub_block in sub_blocks if sub_block.strip()]
                    prev_sub_block = ""
                    print(f"Sub-blocks for keyword '{splittingkeyword}':")
                    for sb in sub_blocks:
                        print(sb)
                        print("--------------------------------------------------------------------------------------------")

                    for sub_block in sub_blocks:
                        processed_block = self.split_data(sub_block, keywords_list)  # process each sub block
                        for dict_ in processed_block:
                            for keyword in keywords_list:
                                if dict_ is not None and keyword in dict_:  # if keyword is in this dictionary
                                    stringtobeadded += dict_[keyword] + " "
                        
                        # Find the most recent phrasequantity phrase and extract the phrasequantity
                        phrasequantity = 1
                        last_line = ""
                        last_phrase = ""

                        print("PREV_SUB_BLOCK", prev_sub_block)
                        
                        lines = prev_sub_block.splitlines()
                        for line in lines:
                            for quantityphrase in self.quantityphrases:
                                if quantityphrase in line:
                                    last_line = line
                                    last_phrase = quantityphrase
                        
                        for quantindex, phrase in enumerate(self.quantityphrases):
                            if phrase in last_phrase:
                                seperatedphrase = last_line.split()
                                phrasequantity = int(seperatedphrase[int(self.quantitypositions[quantindex])])
                                    
                        prev_sub_block = sub_block  # for phrasequantity
                        
                        for removal in self.removals_list:
                            stringtobeadded = stringtobeadded.replace(removal, "")
                        stringtobeadded = ' '.join(stringtobeadded.split())
                        for _ in range(int(phrasequantity)):
                            modified_phrases.append(stringtobeadded)
                        stringtobeadded = ""

            i = 0
            incompletekeyword = ""
            while i < len(phrases):
                phrase = phrases[i].replace("'", "'")
                
                for quantindex, quantphrase in enumerate(self.quantityphrases):
                    if quantphrase.replace("'","'") in phrase:
                        seperatedphrase = phrase.split()
                        phrasequantity = int(seperatedphrase[int(self.quantitypositions[quantindex])])
                
                for incompindex, incomp in enumerate(self.incompletephrases):
                    if incomp in phrase:
                        incompletekeyword = self.secondarykeywords[incompindex]
                
                if incompletekeyword:
                    print("phrasequantity " + str(phrasequantity))
                    for _ in range(phrasequantity):
                        modified_phrases.append(phrase + " " + incompletekeyword)
                        print((phrase + " " + incompletekeyword).strip())
                
                # Exact phrases may be multi-lined
                for exactindex, exactphrase in enumerate(self.exactphrases):
                    linebreaks = exactphrase.count("\n")
                    remaining_lines = phrases[i:i + linebreaks + 1]
                    blockforexactcheck = "\n".join(remaining_lines)
                    if exactphrase in blockforexactcheck:
                        for item in self.exactphraseitems_2d[exactindex]:
                            for _ in range(phrasequantity):
                                modified_phrases.append(item)
                else:
                    for _ in range(phrasequantity):
                        modified_phrases.append(phrase)
                i += 1
            
            # Match with UPC codes
            print(f"Matching with UPC codes for order {order_num}")
            for modified_phrase in modified_phrases:
                for item, upc in self.upc_codes.items():
                    acceptable_words = item.split()
                    modified_words = modified_phrase.split()
                    
                    if set(modified_words) == set(acceptable_words):
                        barcode = upc  # Use the UPC code directly
                        item = " ".join(acceptable_words)  # join the remaining words as the item
                        
                        if order_num not in order_dict:
                            order_dict[order_num] = [(item, barcode, 1)]
                        else:
                            item_found = False
                            for i, (item_, barcode_, count) in enumerate(order_dict[order_num]):
                                if item_ == item and barcode_ == barcode:
                                    item_found = True
                                    order_dict[order_num][i] = (item_, barcode_, count + 1)
                                    print(f"Updated existing item quantity: {item_}, New quantity: {count + 1}")
                                    break
                            if not item_found:
                                order_dict[order_num].append((item, barcode, 1))
                                print(f"Added new item to order {order_num}: {item}")
        
        print("Final order_dict:")
        print(order_dict)
        if order_dict:
            self.display_matched_order(order_dict)
        else:
            self.display_results("No matches found.")
        
        print("Finished perform_full_matching")

    def display_matched_order(self, order_dict):
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        
        self.results_text.insert(tk.END, "Matched Orders:\n\n", "normal")
        
        for order_num, items in order_dict.items():
            self.results_text.insert(tk.END, f"Order Number: {order_num}\n", "normal")
            
            if order_num in self.last_order_dict:
                last_items = self.last_order_dict[order_num]
                last_items_dict = {item[0]: (item[1], item[2]) for item in last_items}
                
                for i, (item, barcode, count) in enumerate(items, start=1):
                    if item in last_items_dict:
                        last_barcode, last_count = last_items_dict[item]
                        if count != last_count or barcode != last_barcode:
                            self.results_text.insert(tk.END, f"Item {i}: {count} {item}\n", "blue")
                        else:
                            self.results_text.insert(tk.END, f"Item {i}: {count} {item}\n", "normal")
                    else:
                        self.results_text.insert(tk.END, f"Item {i}: {count} {item}\n", "blue")
            else:
                for i, (item, barcode, count) in enumerate(items, start=1):
                    self.results_text.insert(tk.END, f"Item {i}: {count} {item}\n", "blue")
            
            self.results_text.insert(tk.END, "\n", "normal")
        
        self.results_text.config(state='disabled')
        self.results_text.see("1.0")
        
        self.last_order_dict = order_dict.copy()

    def open_item_input_window(self):
        input_window = tk.Toplevel(self.root)
        input_window.title("Enter Items")
        input_window.geometry("400x300")

        instruction_label = ttk.Label(input_window, text="Enter items (one per line):", wraplength=380)
        instruction_label.pack(pady=10)

        text_area = tk.Text(input_window, height=10, width=50)
        text_area.pack(pady=10)

        def submit_items():
            items = text_area.get("1.0", tk.END).strip().split("\n")
            self.itemslist = [item.strip() for item in items if item.strip()]
            input_window.destroy()
            self.display_results(f"Phrase: {self.exactphrase}\n\nItems:\n" + "\n".join(f"Item {i+1}: {item}" for i, item in enumerate(self.itemslist)))
            self.next_step()  # Move to the next step automatically

        submit_button = ttk.Button(input_window, text="Submit", command=submit_items)
        submit_button.pack(pady=10)

        input_window.transient(self.root)
        input_window.grab_set()
        self.root.wait_window(input_window)

    def convert_to_rtf(self, text_list):
        rtf = r'{\rtf1\ansi\deff0'
        rtf += r'{\colortbl;\red0\green0\blue0;\red0\green0\blue255;}'  # Define colors: 1=black, 2=blue
        
        for item in text_list:
            if isinstance(item, tuple):
                color, text = item
                if color == 'blue':
                    rtf += r'\cf2 ' + text + r'\cf1 '
                else:
                    rtf += text
            else:
                rtf += item
        
        rtf += '}'
        return rtf

    def split_data(self, text, keywords):
        # Filter out empty keywords
        keywords = [kw for kw in keywords if kw.strip()]
        
        # Initialize the list to store the result
        results = []
        
        # While loop to handle multiple blocks of text
        while keywords:
            # Only consider keywords that are in the text
            present_keywords = [keyword for keyword in keywords if keyword in text]
            
            if not present_keywords:
                break
            
            # Sort keywords based on their order of appearance in the text
            present_keywords.sort(key=text.index)
            
            # Initialize a dictionary for this block of text
            block_results = {}
            
            # Loop through the keywords
            for keyword in present_keywords:
                parts = text.split(keyword, 1)
                if len(parts) > 1:
                    before, after = parts
                    block_results[keyword] = after.split('\n', 1)[0].strip()
                    text = after
                else:
                    break

            if block_results:
                results.append(block_results)
            keywords = [kw for kw in keywords if kw not in block_results]
        
        return results

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
            "Step 3: Check if the keyword is at the start of each block",
            "Step 4: Highlight phrase where package number will appear right after and press Ctrl+L to commit it",
            "Step 5: Highlight package number phrase including the package number and press Ctrl+L to commit it",
            "Step 6: Check the right and see if it has matched orders"
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
                self.display_results("Order 1: \n" + block1)
            elif self.current_direction_index == 1 and self.selected_text:
                block2 = self.selected_text
                self.display_results("Order 2: \n" + block2)
            elif self.current_direction_index == 2:
                blockseperator = self.find_matching_starting_words(block1, block2)
                self.display_results("Keyword: \n" + blockseperator)
            elif self.current_direction_index == 3 and self.selected_text:
                packagenumberphrase = self.selected_text
                self.display_results("Package Phrase: " + packagenumberphrase)
            elif self.current_direction_index == 4:
                match = re.search(f"{re.escape(packagenumberphrase)} (\w+)", self.selected_text)
                if match:
                    self.display_results("Package Number: " + match.group(1))
            elif self.current_direction_index == 5:
                if not blockseperator or not packagenumberphrase:
                    self.display_results("Error: Block separator or package number phrase is missing. Please complete all steps.")
                    return

                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                c.execute('''CREATE TABLE IF NOT EXISTS blockseperator
                            (id INTEGER PRIMARY KEY, keywordstype TEXT, blockseperator TEXT, packagenumberphrase TEXT)''')
                
                keyword_type = 'blockseperator'
                c.execute("INSERT INTO blockseperator (keywordstype, blockseperator, packagenumberphrase) VALUES (?, ?, ?)",
                        (keyword_type, blockseperator, packagenumberphrase))
                conn.commit()
                conn.close()

                # Load new configuration and run test
                self.load_database_configurations()
                self.perform_full_matching()
            
            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def pair_setup(self):
        self.max_steps = 3  # Set max steps for this function
        self.update_step_indicator()
        
        self.directions = [
            "Step 1: Highlight Each Keyword/Identifier in Each Phrase. The order in which you highlighted should match way the item would be read. Press Ctrl+L after each highlight.",
            "Step 2: Highlight the Data You DO NOT Care About in Each Phrase. Any data that isn't a keyword or removal will be considered part of the item. Press Ctrl+L after each highlight.",
            "Step 3: Check results on the right. New items are highlighted in blue"
        ]
        
        keywordslist = []
        unwantedlist = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text not in keywordslist:
                keywordslist.append(self.selected_text)
                self.display_results("\n".join(f"Keyword {i+1}: {kw}" for i, kw in enumerate(keywordslist)))
            elif self.current_direction_index == 1 and self.selected_text and self.selected_text not in unwantedlist:
                unwantedlist.append(self.selected_text)
                self.display_results("\n".join(f"Unwanted {i+1}: {uw}" for i, uw in enumerate(unwantedlist)))
            elif self.current_direction_index == 2:
                if not keywordslist:
                    self.display_results("Error: No keywords entered. Please complete all steps.")
                    return

                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                c.execute('''CREATE TABLE IF NOT EXISTS pairings
                            (id INTEGER PRIMARY KEY, keywordstype TEXT, keywords TEXT, removals TEXT)''')
                
                keyword_type = 'pairing'
                keywords = '<'.join(keywordslist)
                removals = '<'.join(unwantedlist)
                c.execute("INSERT INTO pairings (keywordstype, keywords, removals) VALUES (?, ?, ?)",
                        (keyword_type, keywords, removals))
                conn.commit()
                conn.close()

                self.load_database_configurations()
                self.perform_full_matching()

            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def phrasequantity_setup(self):
        self.max_steps = 3  # Set max steps for this function
        self.update_step_indicator()

        self.directions = [
            "Step 1: Highlight Each full phrase containing the quantity for an item. If you plan to do multiple at a time, maintain the order for the next step. Press Ctrl+L after each highlight.",
            "Step 2: Highlight the quantity in each phrase. Again if doing multiple, highlight the first phrase's quantity first. Press Ctrl+L after each highlight.",
            "Step 3: Check results on the right. New items are highlighted in blue"
        ]
        self.current_direction_index = 0
        self.update_directions(self.directions[self.current_direction_index])
        
        phraseslist = []
        data = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text not in phraseslist:
                phraseslist.append(self.selected_text)
                self.display_results("\n".join(f"Phrase {i+1}: {p}" for i, p in enumerate(phraseslist)))
            elif self.current_direction_index == 1 and self.highlightbuttonpressed:
                data.append(self.selected_text)
                self.display_results("\n".join(f"QuantityPiece {i+1}: {d}" for i, d in enumerate(data)))
            elif self.current_direction_index == 2:
                if not phraseslist or not data or len(phraseslist) != len(data):
                    self.display_results("Error: Incomplete quantity setup. Please ensure all phrases and quantities are entered.")
                    return

                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
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
                self.load_database_configurations()
                self.perform_full_matching()

            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def incompletephrase_setup(self):
        self.max_steps = 3  # Set max steps for this function
        self.update_step_indicator()

        self.directions = [
            "Step 1: Highlight Each Incomplete Phrase. Maintain order for the next step. Press Ctrl+L after each highlight.",
            "Step 2: Highlight the word that will be appended to the item describers below. Press Ctrl+L after each highlight.",
            "Step 3: Check results on the right. New items are highlighted in blue"
        ]
        self.current_direction_index = 0
        self.update_directions(self.directions[self.current_direction_index])
        
        phraseslist = []
        data = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text not in phraseslist:
                phraseslist.append(self.selected_text)
                self.display_results("\n".join(f"Phrase {i+1}: {p}" for i, p in enumerate(phraseslist)))
            elif self.current_direction_index == 1 and self.highlightbuttonpressed:
                data.append(self.selected_text)
                self.display_results("\n".join(f"SecondaryPiece {i+1}: {d}" for i, d in enumerate(data)))
            elif self.current_direction_index == 2:
                if not phraseslist or not data or len(phraseslist) != len(data):
                    self.display_results("Error: Incomplete phrase setup. Please ensure all phrases and secondary keywords are entered.")
                    return

                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
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
                self.load_database_configurations()
                self.perform_full_matching()

            self.highlightbuttonpressed = False
        
        self.current_function = process_step

    def exactphrase_setup(self):
        self.max_steps = 3
        self.update_step_indicator()

        self.directions = [
            "Step 1: Highlight One Exact Order Phrase. Make sure to EXCLUDE data that can change (Like Quantity). Press Ctrl+L to commit it.",
            "Step 2: A new window will open. Enter each item to be added when this phrase is seen. These items should match your UPC items and the website item format.",
            "Step 3: Check results on the right. New items are highlighted in blue"
        ]
        self.current_direction_index = 0
        self.update_directions(self.directions[self.current_direction_index])
        
        self.exactphrase = ""
        self.itemslist = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text != self.exactphrase:
                self.exactphrase = self.selected_text
                self.display_results(f"Phrase: {self.exactphrase}")
                self.next_step()  # Automatically move to the next step
            elif self.current_direction_index == 1:
                self.open_item_input_window()
                # The next_step() is called within the open_item_input_window method
            elif self.current_direction_index == 2:
                if not self.exactphrase or not self.itemslist:
                    self.display_results("Error: Exact phrase or items are missing. Please complete all steps.")
                    return

                conn = sqlite3.connect(path_to_db)
                c = conn.cursor()
                
                c.execute('''CREATE TABLE IF NOT EXISTS exactphrases
                            (id INTEGER PRIMARY KEY, keywordstype TEXT, exactphrases TEXT, items TEXT)''')
                
                keyword_type = 'exactphrase'
                c.execute("INSERT INTO exactphrases (keywordstype, exactphrases, items) VALUES (?, ?, ?)",
                        (keyword_type, self.exactphrase, '<'.join(self.itemslist)))
                conn.commit()
                conn.close()
                
                self.load_database_configurations()
                self.perform_full_matching()
            
            self.highlightbuttonpressed = False
        
        self.current_function = process_step


    def display_results(self, text, tag="normal"):
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert(tk.END, text, tag)
        self.results_text.config(state='disabled')
        self.results_text.see("1.0")  # Scroll to the top

def main():
    root = tk.Tk()
    app = MatchingApp(root)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

if __name__ == "__main__":
    main()
