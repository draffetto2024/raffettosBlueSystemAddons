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
    # Add these helper functions at the start of the MatchingApp class

    def check_block_separator_duplicate(self, block_separator, package_phrase):
        """Check if block separator or package phrase already exists"""
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()
        
        # Check for either matching block separator or package phrase
        c.execute("""
            SELECT blockseperator, packagenumberphrase 
            FROM blockseperator 
            WHERE blockseperator = ? OR packagenumberphrase = ?
        """, (block_separator, package_phrase))
        
        result = c.fetchone()
        conn.close()
        
        if result:
            message = f"Cannot add new entry - duplicate found:\n"
            if result[0] == block_separator:
                message += f"Block separator '{block_separator}' already exists"
            if result[1] == package_phrase:
                message += f"\nPackage phrase '{package_phrase}' already exists"
            messagebox.showerror("Duplicate Entry", message)
            return True
        return False

    def check_quantity_duplicate(self, phrase):
        """Check if quantity phrase already exists"""
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()
        
        c.execute("SELECT phrases FROM quantitys WHERE phrases = ?", (phrase,))
        result = c.fetchone()
        conn.close()
        
        if result:
            messagebox.showerror("Duplicate Entry", 
                f"Cannot add new entry - the quantity phrase:\n'{phrase}'\nalready exists")
            return True
        return False

    def check_exact_phrase_duplicate(self, phrase):
        """Check if exact phrase already exists"""
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()
        
        c.execute("SELECT exactphrases FROM exactphrases WHERE exactphrases = ?", (phrase,))
        result = c.fetchone()
        conn.close()
        
        if result:
            messagebox.showerror("Duplicate Entry", 
                f"Cannot add new entry - the exact phrase:\n'{phrase}'\nalready exists")
            return True
        return False

    def check_pairing_duplicate(self, keywords):
        """Check if keyword combination already exists"""
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()
        
        # Get all existing keywords
        c.execute("SELECT keywords FROM pairings")
        existing_keywords = c.fetchall()
        conn.close()
        
        # Convert new keywords to a set for comparison
        new_keywords_set = set(keywords.split('<'))
        
        # Check against each existing keyword combination
        for (existing,) in existing_keywords:
            if existing and set(existing.split('<')) == new_keywords_set:
                messagebox.showerror("Duplicate Entry", 
                    f"Cannot add new entry - the keyword combination:\n'{keywords}'\nalready exists")
                return True
        return False

    def check_incomplete_phrase_duplicate(self, phrase):
        """Check if incomplete phrase already exists"""
        conn = sqlite3.connect(path_to_db)
        c = conn.cursor()
        
        c.execute("SELECT phrases FROM incompletephrases WHERE phrases = ?", (phrase,))
        result = c.fetchone()
        conn.close()
        
        if result:
            messagebox.showerror("Duplicate Entry", 
                f"Cannot add new entry - the incomplete phrase:\n'{phrase}'\nalready exists")
            return True
        return False

    # Add this new method to the MatchingApp class:
    def reload_upc_codes(self):
        """Reload the UPC codes from the Excel file"""
        try:
            self.load_upc_codes()  # Reload the UPC codes
            messagebox.showinfo("Success", "UPC Codes have been reloaded successfully!")
            print("UPC Codes reloaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to reload UPC Codes: {str(e)}")
            print(f"Error reloading UPC codes: {str(e)}")

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
        self.root.bind('<Control-m>', lambda e: self.perform_full_matching())


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
        
        # In setup_ui, replace all the text widget setup code with this:
        # Middle third: Text widget (single text box for orders)
        middle_frame = ttk.Frame(main_frame, padding="10")
        middle_frame.grid(row=1, column=1, sticky=(tk.N, tk.S, tk.W, tk.E))
        
        # Add header for the middle entry box
        entry_header = ttk.Label(middle_frame, text="Entry box for orders", font=("Arial", 12, "bold"))
        entry_header.pack(pady=(0, 5))
        
        # Create and store scrollbar for left text widget
        self.left_scrollbar = ttk.Scrollbar(middle_frame)
        self.left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.text_widget = tk.Text(middle_frame, wrap=tk.WORD, width=50, height=20, 
                                font=("Arial", 12), yscrollcommand=self.left_scrollbar.set)
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        self.left_scrollbar.config(command=self.text_widget.yview)
        
        # Right third: Results
        results_frame = ttk.Frame(main_frame, padding="10")
        results_frame.grid(row=1, column=2, sticky=(tk.N, tk.S, tk.W, tk.E))
        
        # Add header for the output display
        output_header = ttk.Label(results_frame, text="Output display", font=("Arial", 12, "bold"))
        output_header.pack(pady=(0, 5))
        
        # Create and store scrollbar for right text widget
        self.right_scrollbar = ttk.Scrollbar(results_frame)
        self.right_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.results_text = tk.Text(results_frame, wrap=tk.WORD, width=50, height=20,
                                font=("Arial", 12), yscrollcommand=self.right_scrollbar.set)
        self.results_text.pack(fill=tk.BOTH, expand=True)
        self.right_scrollbar.config(command=self.results_text.yview)

        # Configure tags for coloring
        self.results_text.tag_configure("normal", foreground="black")
        self.results_text.tag_configure("blue", foreground="blue")
        
        # Bind mousewheel scrolling
        self.text_widget.bind("<MouseWheel>", self.on_mousewheel)
        self.results_text.bind("<MouseWheel>", self.on_mousewheel)
        
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

        # Add Reload UPC Codes button
        self.reload_upc_button = ttk.Button(options_frame, text="Reload UPC Codes", command=self.reload_upc_codes)
        self.reload_upc_button.grid(row=0, column=len(self.buttons), padx=5, pady=5)
        
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

    def enable_scroll_sync(self):
        """Enable synchronized scrolling between text boxes"""
        def combined_scroll(*args):
            # Update both text widgets and scrollbars
            self.text_widget.yview_moveto(args[0])
            self.results_text.yview_moveto(args[0])
            
        # Configure both text widgets to use combined scroll
        self.text_widget.config(yscrollcommand=combined_scroll)
        self.results_text.config(yscrollcommand=combined_scroll)
        
        # Configure scrollbars to move both text widgets
        def on_scrollbar(*args):
            self.text_widget.yview(*args)
            self.results_text.yview(*args)
        
        self.left_scrollbar.config(command=on_scrollbar)
        self.right_scrollbar.config(command=on_scrollbar)

    def disable_scroll_sync(self):
        """Disable synchronized scrolling between text boxes"""
        # Restore original scrollbar configurations
        self.text_widget.config(yscrollcommand=self.left_scrollbar.set)
        self.results_text.config(yscrollcommand=self.right_scrollbar.set)
        self.left_scrollbar.config(command=self.text_widget.yview)
        self.right_scrollbar.config(command=self.results_text.yview)

    def on_mousewheel(self, event):
        # Get the widget that triggered the event
        widget = event.widget

        # Determine scroll direction (Windows)
        delta = -1 * (event.delta // 120)

        # Scroll both text widgets
        self.text_widget.yview_scroll(delta, "units")
        self.results_text.yview_scroll(delta, "units")

        self.sync_scrolling(event)
        return "break"  # Prevent default scrolling

    def sync_scrolling(self, event=None):
        """Synchronize scrolling position between text widgets"""
        # Get the widget that triggered the event
        widget = event.widget
        
        # Get the first visible line of the widget that was scrolled
        fraction = widget.yview()[0]
        
        # Update both text widgets to the same position
        self.text_widget.yview_moveto(fraction)
        self.results_text.yview_moveto(fraction)
        return "break"  # Prevent default scrolling

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
        print("\n=== Starting button click ===")
        current_position = self.text_widget.yview()[0]
        print("Initial current_position:", current_position)
        
        if self.active_button:
            self.active_button[0].config(highlightbackground='#E6F3FF', highlightthickness=2)
        
        clicked_frame.config(highlightbackground='#4682B4', highlightthickness=2)
        self.active_button = (clicked_frame, function)
        self.current_function = function
        self.current_step = 0
        
        # Clear and update the results text
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.config(state='disabled')
        
        # Run the function first - this sets up the directions list
        function()
        
        # Now update UI elements after directions are populated
        self.update_step_indicator()
        self.current_direction_index = 0
        
        if self.directions:
            self.update_directions(self.directions[self.current_direction_index])
        
        # Use after_idle to ensure all content is loaded before restoring position
        def restore_positions():
            print("Restoring positions after content load")
            print("Attempting to restore to:", current_position)
            self.results_text.yview_moveto(current_position)
            self.text_widget.yview_moveto(current_position)
            print("Final positions:")
            print("Text widget position:", self.text_widget.yview()[0])
            print("Results text position:", self.results_text.yview()[0])
        
        self.root.after_idle(restore_positions)
        print("=== Finished button click setup ===\n")

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

    # First modify update_directions to not reset the scroll position:
    def update_directions(self, new_text):
        # Store current position
        print("\n=== Starting update_directions ===")
        current_position = self.text_widget.yview()[0]
        print("Position at start of update_directions:", current_position)
        
        self.directions_text.set(new_text)
        # self.selected_text = ""
        # self.results_text.config(state='normal')
        # self.results_text.delete('1.0', tk.END)
        # self.results_text.config(state='disabled')
    
        print("Position after updates:", self.text_widget.yview()[0])
        print("=== Finished update_directions ===\n")

    def on_selection(self, event=None):
        try:
            # Store current scroll position
            current_position = self.text_widget.yview()[0]
            
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
            # Store position again in case it changed
            current_position = self.text_widget.yview()[0]
            # Remove selection
            self.text_widget.tag_remove(tk.SEL, "1.0", tk.END)
            # Restore scroll position
            self.text_widget.yview_moveto(current_position)
            self.results_text.yview_moveto(current_position)

    def load_upc_codes(self):
        path_to_xlsx = "./UPCCodes.xlsx"  # Changed to use relative path
        df = pd.read_excel(path_to_xlsx, engine='openpyxl', header=None, dtype=str)
        df.columns = ["Item Name", "UPC Code"]
        df["Item Name"] = df["Item Name"].str.lower()
        df["UPC Code"] = df["UPC Code"].str.rstrip('.0')
        # Remove any rows with NaN values
        df = df.dropna()
        self.upc_codes = dict(zip(df["Item Name"], df["UPC Code"]))
        print(f"Loaded {len(self.upc_codes)} UPC codes")
        print("Sample UPC codes:")
        for item, upc in list(self.upc_codes.items()):
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

    def split_blocks(self, text, keyword):
        # Split the text on the keyword
        parts = text.split(keyword)
        
        # The first part doesn't need the keyword prepended
        blocks = [parts[0]]
        
        # For all other parts, prepend the keyword
        for part in parts[1:]:
            blocks.append(keyword + part)
        
        return blocks

    def print_blocks(self, blocks, block_separator="=" * 50):
        print("\nBlocks:")
        for i, block in enumerate(blocks, 1):
            print(f"\n{block_separator}")
            print(f"Block {i}:")
            print(f"{block_separator}")
            
            # Split the block into lines and print each line
            lines = block.split('\n')
            for line in lines:
                print(line.strip())

    def perform_full_matching(self):
        print("\n=== Starting full matching ===")
    # Store current scroll position of both text widgets
        self.initial_text_position = self.text_widget.yview()
        current_position = self.text_widget.yview()
        self.initial_results_position = self.results_text.yview()        
        # Enable sync just for matching
        self.enable_scroll_sync()


        print("Full matching - Initial position:", current_position)
        print("Starting perform_full_matching")
        text_input = self.text_widget.get("1.0", tk.END).lower().strip()
        print(f"Input text: {text_input}")
        order_dict = {}

        blocks = self.split_blocks(text_input, self.blockseperatorkeyword)
        # self.print_blocks(blocks)        

        for block in blocks:
            order_num = ""
            phrases = block.split("\n")
            modified_phrases = []  # list to store modified phrases
            phrases_length = len(phrases)
            phrasequantity = 1
            
            # print("---------------------------------------------------------------------------------")

            # print("PHRASES", phrases)
            # print("PACKAGENUMBERPHRASE", self.packagenumberphrase)

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

            # Initialize the order in the dictionary, even if no items are found
            if order_num not in order_dict:
                order_dict[order_num] = []

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

            # print("keywords_2d_list", self.keywords_2d_list)
            # print("removals_list", self.removals_list)
            
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
                    # print(f"Sub-blocks for keyword '{splittingkeyword}':")
                    # for sb in sub_blocks:
                        # print(sb)
                        # print("--------------------------------------------------------------------------------------------")

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

                        # print("PREV_SUB_BLOCK", prev_sub_block)
                        
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
                    # print("phrasequantity " + str(phrasequantity))
                    for _ in range(phrasequantity):
                        modified_phrases.append(phrase + " " + incompletekeyword)
                        # print((phrase + " " + incompletekeyword).strip())
                
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
        self.display_matched_order(order_dict)

        # # Always disable sync when we're done
        # self.disable_scroll_sync()

         # Restore the initial scroll positions
        self.text_widget.yview_moveto(self.initial_text_position[0])
        self.results_text.yview_moveto(self.initial_results_position[0])

        print("Full matching - After display, position:", self.text_widget.yview()[0])
        print("=== Finished full matching ===\n")
        
        print("Finished perform_full_matching")

    def display_matched_order(self, order_dict):
        # Store current scroll position
        current_position = self.results_text.yview()[0]

        # print("\n=== Starting display_matched_order ===")
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        
        # Get the text from the left text box and split into blocks
        left_text = self.text_widget.get('1.0', tk.END)
        # print("BLOCKSEP", self.blockseperatorkeyword)
        blocks = self.split_blocks((left_text.strip()).lower(), self.blockseperatorkeyword)
        # print(blocks)
        # print(f"Number of blocks: {len(blocks)}")
        
        for block in blocks:
            if not block.strip():
                continue
                
            # Find the order number for this block
            order_num = None
            for line in block.split('\n'):
                match = re.search(f"{re.escape(self.packagenumberphrase)} ([^\s]+)", line)
                if match:
                    order_num = ''.join([char for char in match.group(1) if char.isdigit()])
                    # print(f"Found order number: {order_num}")
                    break
            
            if order_num:
                # Count lines in THIS SPECIFIC block
                block_lines = len([line for line in block.split('\n') if line.strip()])
                # print(f"\nProcessing order {order_num} with {block_lines} lines in its block")
                lines_used = 0
                
                # Insert order header
                self.results_text.insert(tk.END, f"Order Number: {order_num}\n", "normal")
                lines_used += 1
                
                # Insert items
                if order_num in order_dict:
                    items = order_dict[order_num]
                    # print(f"Number of items for order {order_num}: {len(items)}")
                    
                    if not items:
                        self.results_text.insert(tk.END, "No items found for this order\n", "blue")
                        lines_used += 1
                    else:
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
                                lines_used += 1
                        else:
                            for i, (item, barcode, count) in enumerate(items, start=1):
                                self.results_text.insert(tk.END, f"Item {i}: {count} {item}\n", "blue")
                                lines_used += 1
                
                # print(f"Lines used in display: {lines_used}")
                # print(f"Lines in original block: {block_lines}")
                
                # Calculate and add padding to align with left side
                padding_needed = max(0, block_lines - lines_used)
                # print(f"Adding {padding_needed} padding lines for alignment")
                
                if padding_needed > 0:
                    self.results_text.insert(tk.END, '\n' * padding_needed)
                
                # Add a separator line between orders
                self.results_text.insert(tk.END, "-" * 50 + "\n", "normal")
        
        # print("=== Finished display_matched_order ===")


        self.results_text.config(state='disabled')
        # After all content is added and widget is configured, restore scroll position
        self.results_text.yview_moveto(current_position)
        # Synchronize the left text widget to match
        self.text_widget.yview_moveto(current_position)

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
            "Step 1: Highlight one order as a block of text including any important words before or after and press Ctrl+L to commit it. Ctrl+N to go to next step",
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

                # Add duplicate check
                if self.check_block_separator_duplicate(blockseperator, packagenumberphrase):
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
            "Step 1: Highlight Each Keyword/Identifier, NOT the actual data part, of a singlular pairing schema. The order in which you highlighted should match way the item would be read. Press Ctrl+L after each highlight. Ctrl+N to go to next step",
            "Step 2: Highlight the Data You DO NOT Care About in Each Phrase. Any data that isn't a keyword or removal will be considered part of the item. Press Ctrl+L after each highlight.",
            "Step 3: Check results on the right. New items are highlighted in blue"
        ]
        
        keywordslist = []
        unwantedlist = []
        
        def process_step():
            if self.current_direction_index == 0 and self.selected_text and self.selected_text not in keywordslist:
                keywordslist.append(self.selected_text)
                print("About to call display_results from pair_setup step 0")
                self.display_results("\n".join(f"Keyword {i+1}: {kw}" for i, kw in enumerate(keywordslist)))
            elif self.current_direction_index == 1 and self.selected_text and self.selected_text not in unwantedlist:
                unwantedlist.append(self.selected_text)
                self.display_results("\n".join(f"Unwanted {i+1}: {uw}" for i, uw in enumerate(unwantedlist)))
            elif self.current_direction_index == 2:
                if not keywordslist:
                    self.display_results("Error: No keywords entered. Please complete all steps.")
                    return

                keywords = '<'.join(keywordslist)
                
                # Add duplicate check
                if self.check_pairing_duplicate(keywords):
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
            "Step 1: Highlight Each full phrase containing the quantity for an item. If you plan to do multiple at a time, maintain the order for the next step. Press Ctrl+L after each highlight. Ctrl+N to go to next step",
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
                        
                        # Add duplicate check
                        if self.check_quantity_duplicate(phrase):
                            continue
                        
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
            "Step 1: Highlight Each Incomplete Phrase. Make sure to EXCLUDE things like quantity. Maintain order for the next step. Press Ctrl+L after each highlight. Ctrl+N to go to next step",
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
                    
                    # Add duplicate check
                    if self.check_incomplete_phrase_duplicate(phrase):
                        continue
                    
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
            "Step 1: Highlight One Exact Order Phrase. Make sure to EXCLUDE data that can change (Like Quantity). Press Ctrl+L to commit it. Ctrl+N to go to next step",
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

                # Add duplicate check
                if self.check_exact_phrase_duplicate(self.exactphrase):
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
        # Store the results text position independently
        current_results_position = self.results_text.yview()[0]
        
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert(tk.END, text, tag)
        self.results_text.config(state='disabled')
        
        # Only restore results text position
        self.results_text.yview_moveto(current_results_position)

def main():
    try:
        root = tk.Tk()
        app = MatchingApp(root)
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        root.mainloop()
    except Exception as e:
        print("An error occurred:")
        print(str(e))
        print("\nTraceback:")
        traceback.print_exc()

if __name__ == "__main__":
    main()
