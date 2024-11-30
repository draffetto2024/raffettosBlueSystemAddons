import os
import re
import cv2
import shutil
import smtplib
import pandas as pd
import fitz  # PyMuPDF
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.cloud import vision
from google.cloud.vision_v1 import types
from datetime import datetime
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from collections import defaultdict




# Simplified path definitions using relative paths
service_account_key = "./caramel-compass-429017-h3-c2d4e157e809.json"
invoice_folder = "./Invoices/InvoicePictures"
destination_base_folder = "./Invoices/SortedInvoices"
unsorted_base_folder = "./Invoices/UnsortedInvoices"
customer_emails_file = "./customer_emails.xlsx"

# Set the Google Application Credentials environment variable
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = service_account_key

# Initialize the Vision API client
client = vision.ImageAnnotatorClient()

# Email details
sender_email = "gingoso2@gmail.com"
app_password = "soiz avjw bdtu hmtn"

class ReportManager:
    def __init__(self):
        self.reports_folder = "./Reports"
        self.current_session_invoices = []
        os.makedirs(self.reports_folder, exist_ok=True)
        
    def add_invoice(self, invoice_num, route_num):
        """Add an invoice to the current session tracking."""
        self.current_session_invoices.append({
            'invoice_num': invoice_num.strip() if invoice_num else invoice_num,
            'route_num': route_num.strip() if route_num else 'Unassigned'
        })
    
    def generate_report(self):
        """Generate or update the daily report for processed invoices."""
        if not self.current_session_invoices:
            print("No invoices to report.")
            return
            
        today = datetime.now().strftime("%Y-%m-%d")
        report_path = os.path.join(self.reports_folder, f"invoice_report_{today}.txt")
        
        # Read existing invoices if report exists
        existing_invoices = []
        if os.path.exists(report_path):
            with open(report_path, 'r') as f:
                # Skip header lines
                lines = f.readlines()[3:]
                for line in lines:
                    if line.strip():
                        # Split on whitespace and combine all but the last part as invoice number
                        parts = line.strip().split()
                        if len(parts) >= 2:
                            route_num = parts[-1]  # Last part is route number
                            invoice_num = ' '.join(parts[:-1])  # All other parts form invoice number
                            existing_invoices.append({
                                'invoice_num': invoice_num,
                                'route_num': route_num
                            })
        
        # Combine existing and new invoices
        all_invoices = existing_invoices + self.current_session_invoices
        
        # Sort by route number first, then by invoice number
        sorted_invoices = sorted(all_invoices, 
                               key=lambda x: (x['route_num'], x['invoice_num']))
        
        # Generate report content
        report_content = f"Invoice Report - Generated {today}\n"
        report_content += "=" * 50 + "\n\n"
        
        current_route = None
        for invoice in sorted_invoices:
            # Add route header if route changes
            if invoice['route_num'] != current_route:
                current_route = invoice['route_num']
                report_content += f"\nRoute Number: {current_route}\n"
                report_content += "-" * 20 + "\n"
            
            # Add invoice with proper spacing for alignment
            report_content += f"{invoice['invoice_num']}    {invoice['route_num']}\n"
        
        # Write report
        with open(report_path, 'w') as f:
            f.write(report_content)
        
        print(f"Report generated: {report_path}")
        
        # Clear current session invoices after report generation
        self.current_session_invoices = []

class RouteManager:
    def __init__(self, base_folder):
        self.base_folder = base_folder
        self.route_map = defaultdict(list)
    
    def extract_route_number(self, texts):
        """Extract route number using existing OCR text detection pattern."""
        return find_text_near_keyphrase(texts, "Driver/Route #", "below", 50)
    
    def add_invoice_to_route(self, route_number, invoice_number, file_path, date):
        """Add an invoice to the route mapping."""
        self.route_map[route_number].append({
            'invoice_number': invoice_number,
            'path': file_path,
            'date': date
        })
        # Sort invoices within this route by invoice number
        self.route_map[route_number].sort(key=lambda x: x['invoice_number'])
    
    def get_sorted_routes(self):
        """Get all routes sorted with their invoices."""
        return dict(sorted(self.route_map.items()))

class CustomerTracker:
    def __init__(self, base_folder, inactivity_threshold_days=21):
        self.base_folder = base_folder
        self.inactivity_threshold = inactivity_threshold_days
        self.last_order_dates = {}
        self.load_customer_history()
    
    def load_customer_history(self):
        """Scan through sorted invoices to build customer order history."""
        print("Loading customer order history...")
        for year_dir in os.listdir(self.base_folder):
            year_path = os.path.join(self.base_folder, year_dir)
            if not os.path.isdir(year_path):
                continue
                
            for month_dir in os.listdir(year_path):
                month_path = os.path.join(year_path, month_dir)
                if not os.path.isdir(month_path):
                    continue
                    
                for day_dir in os.listdir(month_path):
                    day_path = os.path.join(month_path, day_dir)
                    if not os.path.isdir(day_path):
                        continue
                    
                    # Process invoices in this day's folder
                    for invoice in os.listdir(day_path):
                        if not invoice.endswith(('.jpg', '.png', '.pdf')):
                            continue
                        
                        # Extract customer ID from filename (assumed format: CUSTID_INVNUM.ext)
                        customer_id = invoice.split('_')[0]
                        
                        # Create date object from folder structure
                        try:
                            order_date = datetime(
                                int(f"20{year_dir}"),  # Assuming 20xx year format
                                int(month_dir),
                                int(day_dir)
                            )
                            
                            # Update last order date if more recent
                            if customer_id not in self.last_order_dates or \
                               order_date > self.last_order_dates[customer_id]:
                                self.last_order_dates[customer_id] = order_date
                                
                        except ValueError as e:
                            print(f"Error parsing date for {invoice}: {e}")
    
    def update_customer_activity(self, customer_id, date_str):
        """Update customer's last order date when processing new invoice."""
        try:
            # Convert date string (MM/DD/YY) to datetime
            date_parts = date_str.split('/')
            order_date = datetime(
                2000 + int(date_parts[2]),  # Assuming 20xx year
                int(date_parts[0]),
                int(date_parts[1])
            )
            
            # Update last order date if more recent
            if customer_id not in self.last_order_dates or \
               order_date > self.last_order_dates[customer_id]:
                self.last_order_dates[customer_id] = order_date
                
        except (ValueError, IndexError) as e:
            print(f"Error updating customer activity for {customer_id}: {e}")
    
    def get_inactive_customers(self):
        """Return list of customers who haven't ordered in threshold days."""
        current_date = datetime.now()
        inactive_customers = []
        
        for customer_id, last_order in self.last_order_dates.items():
            days_since_order = (current_date - last_order).days
            if days_since_order >= self.inactivity_threshold:
                inactive_customers.append({
                    'customer_id': customer_id,
                    'last_order_date': last_order.strftime('%m/%d/%y'),
                    'days_inactive': days_since_order
                })
        
        # Sort by days inactive (descending)
        inactive_customers.sort(key=lambda x: x['days_inactive'], reverse=True)
        return inactive_customers
    
    def generate_inactivity_report(self):
        """Generate a detailed report of inactive customers."""
        inactive = self.get_inactive_customers()
        if not inactive:
            return "No inactive customers found."
        
        report = "Customer Inactivity Report\n"
        report += "=" * 50 + "\n"
        report += f"Generated on: {datetime.now().strftime('%m/%d/%y %H:%M')}\n"
        report += f"Inactivity threshold: {self.inactivity_threshold} days\n"
        report += "=" * 50 + "\n\n"
        
        for customer in inactive:
            report += f"Customer ID: {customer['customer_id']}\n"
            report += f"Last Order: {customer['last_order_date']}\n"
            report += f"Days Inactive: {customer['days_inactive']}\n"
            report += "-" * 30 + "\n"
        
        return report

# Add this function near your other utility functions
def check_inactive_customers():
    inactive = customer_tracker.get_inactive_customers()
    if inactive:
        print("\nInactive Customer Alert:")
        for customer in inactive:
            print(f"Customer {customer['customer_id']} hasn't ordered in {customer['days_inactive']} days "
                  f"(Last order: {customer['last_order_date']})")

def display_image_and_get_input(image_path):
    """Display image preview and get user input for invoice details."""
    root = tk.Tk()
    root.title("Invoice Preview and Input")
    
    # Store user input in a mutable object that can be accessed by the validate_and_submit function
    user_input = {'customer_id': None, 'invoice_num': None, 'date': None, 'route_num': None}
    
    try:
        # Load and display image
        img = Image.open(image_path)
        # Resize image if too large while maintaining aspect ratio
        max_size = (800, 800)
        img.thumbnail(max_size, Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        
        # Create and pack widgets
        image_label = ttk.Label(root, image=photo)
        image_label.image = photo  # Keep a reference
        image_label.pack(pady=10)
    except Exception as e:
        print(f"Error loading image: {e}")
        image_label = ttk.Label(root, text="[Image preview unavailable]")
        image_label.pack(pady=10)
    
    # Input frame
    input_frame = ttk.Frame(root)
    input_frame.pack(pady=10, padx=10)
    
    # Customer ID input
    ttk.Label(input_frame, text="Customer ID (6 alphanumeric):").grid(row=0, column=0, sticky='e', pady=5)
    customer_id_entry = ttk.Entry(input_frame)
    customer_id_entry.grid(row=0, column=1, pady=5)
    
    # Invoice number input
    ttk.Label(input_frame, text="Invoice Number (6 alphanumeric):").grid(row=1, column=0, sticky='e', pady=5)
    invoice_num_entry = ttk.Entry(input_frame)
    invoice_num_entry.grid(row=1, column=1, pady=5)
    
    # Date input
    ttk.Label(input_frame, text="Date (MM/DD/YY or MMDDYY):").grid(row=2, column=0, sticky='e', pady=5)
    date_entry = ttk.Entry(input_frame)
    date_entry.grid(row=2, column=1, pady=5)
    
    # Route number input
    ttk.Label(input_frame, text="Route Number (optional):").grid(row=3, column=0, sticky='e', pady=5)
    route_num_entry = ttk.Entry(input_frame)
    route_num_entry.grid(row=3, column=1, pady=5)
    
    # Error label
    error_label = ttk.Label(input_frame, text="", foreground="red")
    error_label.grid(row=4, column=0, columnspan=2)
    
    def focus_next_empty(event=None):
        """Focus next empty field or submit if all required fields are filled."""
        current = root.focus_get()
        entries = [customer_id_entry, invoice_num_entry, date_entry, route_num_entry]
        required_entries = [customer_id_entry, invoice_num_entry, date_entry]  # Route is optional
        
        if current in entries:
            current_idx = entries.index(current)
            # Check remaining fields
            for idx in range(current_idx + 1, len(entries)):
                if entries[idx] in required_entries and not entries[idx].get().strip():
                    entries[idx].focus()
                    return "break"
            # If we get here and all required fields are filled, try to submit
            if all(entry.get().strip() for entry in required_entries):
                submit_wrapper()
        return "break"
    
    def submit_wrapper(event=None):
        customer_id = customer_id_entry.get().strip()
        invoice_num = invoice_num_entry.get().strip()
        date = date_entry.get().strip()
        route_num = route_num_entry.get().strip()
        
        # Validate customer ID
        if not is_six_alphanumeric(customer_id):
            error_label.config(text="Invalid Customer ID format - must be 6 alphanumeric characters")
            customer_id_entry.focus()
            customer_id_entry.selection_range(0, tk.END)
            return
            
        # Validate invoice number
        if not is_six_alphanumeric(invoice_num):
            error_label.config(text="Invalid Invoice Number format - must be 6 alphanumeric characters")
            invoice_num_entry.focus()
            invoice_num_entry.selection_range(0, tk.END)
            return
        
        # Validate and possibly convert date format
        date_result = is_date_format(date)
        if isinstance(date_result, str):
            date = date_result  # Date was converted from MMDDYY to MM/DD/YY
        elif not date_result:
            error_label.config(text="Invalid Date format - use MM/DD/YY or MMDDYY")
            date_entry.focus()
            date_entry.selection_range(0, tk.END)
            return
        
        # Store all values
        user_input['customer_id'] = customer_id
        user_input['invoice_num'] = invoice_num
        user_input['date'] = date
        user_input['route_num'] = route_num if route_num else None
        
        root.quit()
        root.destroy()
    
    # Bind Enter key for each entry
    customer_id_entry.bind('<Return>', focus_next_empty)
    invoice_num_entry.bind('<Return>', focus_next_empty)
    date_entry.bind('<Return>', focus_next_empty)
    route_num_entry.bind('<Return>', focus_next_empty)
    
    # Center the window
    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - root.winfo_width()) // 2
    y = (screen_height - root.winfo_height()) // 2
    root.geometry(f"+{x}+{y}")
    
    # Force focus to the first entry field immediately
    root.lift()  # Bring window to front
    root.focus_force()  # Force focus to window
    customer_id_entry.focus_set()  # Force focus to first entry
    root.update()  # Force update of window
    
    try:
        root.mainloop()
    except Exception as e:
        print(f"Error in mainloop: {e}")
    
    return user_input

# Validate paths
def validate_paths():
    paths = {
        "Service Account Key": service_account_key,
        "Invoice Folder": invoice_folder,
        "Destination Base Folder": destination_base_folder,
        "Unsorted Base Folder": unsorted_base_folder,
        "Customer Emails File": customer_emails_file
    }
    for name, path in paths.items():
        if not os.path.exists(path):
            print(f"ERROR: Path does not exist: {name} - {path}")
        else:
            print(f"Path exists: {name} - {path}")
    return all(os.path.exists(path) for path in paths.values())

if not validate_paths():
    print("Exiting due to invalid paths.")
    exit(1)

# Add this line right after it
customer_tracker = CustomerTracker(destination_base_folder)

def load_customer_emails(excel_file):
    df = pd.read_excel(excel_file, header=None)  # No header row
    customer_emails = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))  # First column as keys, second as values
    return customer_emails

customer_emails = load_customer_emails(customer_emails_file)

def extract_text_and_positions(image):
    _, encoded_image = cv2.imencode('.jpg', image)
    content = encoded_image.tobytes()
    image = types.Image(content=content)
    response = client.document_text_detection(image=image)
    if response.error.message:
        raise Exception(f'{response.error.message}')
    
    return response.text_annotations

def print_bounding_box(description, box):
    print(f"{description}:")
    print(f"  Top-left: ({box[0].x}, {box[0].y})")
    print(f"  Bottom-right: ({box[2].x}, {box[2].y})")


def find_text_near_keyphrase(texts, keyphrase, position, threshold):
    """Find text elements that appear in sequence within a threshold distance."""
    logging.info(f"\nSearching for text near '{keyphrase}' in '{position}' position:")
    
    # Split keyphrase into individual characters/words
    keyphrase_parts = []
    current_part = ""
    
    # Split into parts while preserving special characters
    for char in keyphrase:
        if char.isalpha():
            current_part += char
        else:
            if current_part:
                keyphrase_parts.append(current_part)
                current_part = ""
            if not char.isspace():  # Add special characters as their own parts
                keyphrase_parts.append(char)
    if current_part:  # Add final part if exists
        keyphrase_parts.append(current_part)
    
    keyphrase_parts = [part.lower() for part in keyphrase_parts]
    keyphrase_texts = []
    keyphrase_boxes = []
    
    # Search for first part
    for i, text in enumerate(texts):
        if keyphrase_parts[0].lower() in text.description.lower():
            keyphrase_texts = [text]
            keyphrase_boxes = [text.bounding_poly.vertices]
            current_box = text.bounding_poly.vertices
            
            # Look for remaining parts
            last_found_idx = i
            for part in keyphrase_parts[1:]:
                found = False
                # Look at next few texts (reduced threshold to 25)
                for j, next_text in enumerate(texts[last_found_idx + 1:last_found_idx + 5]):
                    if part.lower() in next_text.description.lower():
                        next_box = next_text.bounding_poly.vertices
                        if is_within_threshold(current_box, next_box, 25):  # Reduced threshold
                            logging.info(f"Found next part '{part}' in: '{next_text.description}'")
                            keyphrase_texts.append(next_text)
                            keyphrase_boxes.append(next_box)
                            current_box = combine_boxes(current_box, next_box)
                            last_found_idx = last_found_idx + j + 1
                            found = True
                            break
                if not found:
                    break
            
            if len(keyphrase_texts) == len(keyphrase_parts):
                break
    
    if len(keyphrase_texts) != len(keyphrase_parts):
        logging.info(f"Could not find complete keyphrase '{keyphrase}'")
        return None
    
    # Log the individual boxes for debugging
    logging.info("\nIndividual part bounding boxes:")
    for i, box in enumerate(keyphrase_boxes):
        logging.info(f"Part {i} ('{keyphrase_texts[i].description}'):")
        logging.info(f"  Top-left: ({box[0].x}, {box[0].y})")
        logging.info(f"  Bottom-right: ({box[2].x}, {box[2].y})")
    
    # Create bounding box for complete keyphrase
    keyphrase_box = keyphrase_boxes[0]
    for box in keyphrase_boxes[1:]:
        keyphrase_box = combine_boxes(keyphrase_box, box)
    
    logging.info(f"\nFinal combined keyphrase bounding box:")
    logging.info(f"  Top-left: ({keyphrase_box[0].x}, {keyphrase_box[0].y})")
    logging.info(f"  Bottom-right: ({keyphrase_box[2].x}, {keyphrase_box[2].y})")
    
    # Search for text near the keyphrase
    for adjacent_text in texts:
        if adjacent_text in keyphrase_texts:
            continue
        adjacent_box = adjacent_text.bounding_poly.vertices
        
        if is_in_position(keyphrase_box, adjacent_box, position, threshold):
            if re.search(r'[a-zA-Z0-9]', adjacent_text.description):
                logging.info(f"\nMATCH FOUND: '{adjacent_text.description}'")
                logging.info(f"Match bounding box:")
                logging.info(f"  Top-left: ({adjacent_box[0].x}, {adjacent_box[0].y})")
                logging.info(f"  Bottom-right: ({adjacent_box[2].x}, {adjacent_box[2].y})")
                return adjacent_text.description
    
    logging.info(f"No matching text found near '{keyphrase}'")
    return None

def is_within_threshold(box1, box2, threshold):
    # Check if box2 is within threshold distance of box1
    return (abs(box2[0].x - box1[2].x) < threshold and
            abs(box2[0].y - box1[0].y) < threshold)

def combine_boxes(box1, box2):
    """Combine two bounding boxes more precisely."""
    # Create a minimal box that contains both boxes
    return [
        types.Vertex(x=min(box1[0].x, box2[0].x), y=min(box1[0].y, box2[0].y)),  # Top-left
        types.Vertex(x=max(box1[1].x, box2[1].x), y=min(box1[1].y, box2[1].y)),  # Top-right
        types.Vertex(x=max(box1[2].x, box2[2].x), y=max(box1[2].y, box2[2].y)),  # Bottom-right
        types.Vertex(x=min(box1[3].x, box2[3].x), y=max(box1[3].y, box2[3].y))   # Bottom-left
    ]




# def print_all_text_elements(texts):
#     print("\nAll detected text elements:")
#     for i, text in enumerate(texts):
#         print(f"Text {i + 1}: '{text.description}'")
#         box = text.bounding_poly.vertices
#         print(f"  Top-left: ({box[0].x}, {box[0].y})")
#         print(f"  Bottom-right: ({box[2].x}, {box[2].y})")
#         print()

def print_all_text_elements(texts):
    logging.info("\n=== COMPLETE OCR RESULTS ===")
    for i, text in enumerate(texts[1:]):  # Skip first element which contains all text
        box = text.bounding_poly.vertices
        logging.info(f"\nText Element {i}:")
        logging.info(f"Content: '{text.description}'")
        logging.info(f"Bounding Box:")
        logging.info(f"  Top-left: ({box[0].x}, {box[0].y})")
        logging.info(f"  Top-right: ({box[1].x}, {box[1].y})")
        logging.info(f"  Bottom-right: ({box[2].x}, {box[2].y})")
        logging.info(f"  Bottom-left: ({box[3].x}, {box[3].y})")
    logging.info("\n=== END OCR RESULTS ===\n")


def is_in_position(box1, box2, position, threshold):
    box1_top = min(v.y for v in box1)
    box1_bottom = max(v.y for v in box1)
    box1_left = min(v.x for v in box1)
    box1_right = max(v.x for v in box1)

    box2_top = min(v.y for v in box2)
    box2_bottom = max(v.y for v in box2)
    box2_left = min(v.x for v in box2)
    box2_right = max(v.x for v in box2)

    # Check for partial horizontal overlap
    horizontal_overlap = (
        (box2_left < box1_right and box2_right > box1_left)
    )

    if position == 'below':
        vertical_check = (box2_top > box1_top) and (box2_top - box1_bottom <= threshold)
        return vertical_check and horizontal_overlap
    elif position == 'above':
        vertical_check = (box2_bottom < box1_bottom) and (box1_top - box2_bottom <= threshold)
        return vertical_check and horizontal_overlap
    elif position == 'left':
        horizontal_check = (box2_right < box1_right) and (box1_left - box2_right <= threshold)
        vertical_overlap = (box2_bottom > box1_top and box2_top < box1_bottom)
        return horizontal_check and vertical_overlap
    elif position == 'right':
        horizontal_check = (box2_left > box1_left) and (box2_left - box1_right <= threshold)
        vertical_overlap = (box2_bottom > box1_top and box2_top < box1_bottom)
        return horizontal_check and vertical_overlap
    return False


import re

def validate_and_submit(event=None):  # Added event parameter for Enter key binding
    try:
        customer_id = customer_id_entry.get().strip()
        invoice_num = invoice_num_entry.get().strip()
        date = date_entry.get().strip()
        
        print(f"Validating input - Customer ID: {customer_id}, Invoice: {invoice_num}, Date: {date}")
        
        # Validate customer ID
        if not is_six_alphanumeric(customer_id):
            error_label.config(text="Invalid Customer ID format - must be 6 alphanumeric characters")
            return
            
        # Validate invoice number
        if not is_six_alphanumeric(invoice_num):
            error_label.config(text="Invalid Invoice Number format - must be 6 alphanumeric characters")
            return
        
        # Validate and possibly convert date format
        date_result = is_date_format(date)
        if isinstance(date_result, str):
            # Date was in MMDDYY format and was converted
            date = date_result
            print(f"DEBUG: Converted date format to: {date}")
        elif not date_result:
            error_label.config(text="Invalid Date format - use MM/DD/YY or MMDDYY")
            return
        
        # If all validations pass, store the values and close window
        user_input['customer_id'] = customer_id
        user_input['invoice_num'] = invoice_num
        user_input['date'] = date
        print(f"DEBUG: All validations passed. Input values: {user_input}")
        root.quit()  # This will exit the mainloop
        root.destroy()  # This will destroy the window
        
    except Exception as e:
        print(f"Error in validate_and_submit: {e}")
        error_label.config(text=f"An error occurred: {str(e)}")

def is_six_alphanumeric(s):
    """Check if string is exactly 6 alphanumeric characters."""
    if not s:
        return False
    s = s.strip()
    print(f"DEBUG: Checking if '{s}' is six alphanumeric characters")
    result = bool(re.match(r'^[A-Za-z0-9]{6}$', s))
    print(f"DEBUG: Result of alphanumeric check: {result}")
    return result


def is_date_format(s):
    """Check if string matches MM/DD/YY format."""
    if not s:
        return False
    s = s.strip()
    
    # First try MM/DD/YY format
    pattern1 = re.compile(r'^(0[1-9]|1[0-2])/(0[1-9]|[12][0-9]|3[01])/\d{2}$')
    # Also allow MMDDYY format and convert it
    pattern2 = re.compile(r'^(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])(\d{2})$')
    
    print(f"DEBUG: Checking date format for '{s}'")
    
    if pattern1.match(s):
        print(f"DEBUG: Date matches MM/DD/YY format")
        return True
    elif pattern2.match(s):
        # Convert MMDDYY to MM/DD/YY format
        matches = pattern2.match(s)
        month, day, year = matches.groups()
        return f"{month}/{day}/{year}"
    
    print(f"DEBUG: Date format check failed")
    return False

def format_date(date_str):
    """Convert date string to MM/DD/YY format if needed."""
    if '/' in date_str:
        return date_str
    # Convert MMDDYY to MM/DD/YY
    return f"{date_str[:2]}/{date_str[2:4]}/{date_str[4:]}"

def extract_year(date_str):
    return date_str.split('/')[2]

def extract_month(date_str):
    return date_str.split('/')[0]

def extract_day(date_str):
    return date_str.split('/')[1]

# Update copy_image_to_sorted_folder to handle the complete filename base properly
def copy_image_to_sorted_folder(image_path, filename_base, customer_id, date):
    year, month, day = extract_year(date), extract_month(date), extract_day(date)
    destination_folder = os.path.join(destination_base_folder, year, month, day)
    os.makedirs(destination_folder, exist_ok=True)
    
    # Ensure we're using the complete filename_base
    destination_path = os.path.join(destination_folder, f"{filename_base}.jpg")
    print(f"DEBUG: Attempting to copy from {image_path} to {destination_path}")
    try:
        shutil.copy2(image_path, destination_path)
        print(f"Copied and renamed image to: {destination_path}")
        customer_tracker.update_customer_activity(customer_id, date)
    except Exception as e:
        print(f"DEBUG: Error copying file: {e}")
    
    # Extract invoice number from filename_base for email
    invoice_num = filename_base.split('_')[1]
    send_email_with_attachment(destination_path, customer_id, invoice_num, date)

def copy_image_to_unsorted_folder(image_path):
    today = datetime.now().strftime("%Y-%m-%d")
    destination_folder = os.path.join(unsorted_base_folder, today)
    os.makedirs(destination_folder, exist_ok=True)
    destination_path = os.path.join(destination_folder, os.path.basename(image_path))
    print(f"DEBUG: Attempting to copy unsorted invoice from {image_path} to {destination_path}")
    try:
        shutil.copy2(image_path, destination_path)
        print(f"Copied unsorted invoice to: {destination_path}")
    except Exception as e:
        print(f"DEBUG: Error copying unsorted file: {e}")

def send_email_with_attachment(file_path, customer_id, invoice_number, date):
    receiver_email = customer_emails.get(customer_id, None)
    if receiver_email is None:
        print(f"No email found for customer {customer_id}. Skipping email.")
        return

    subject = f"Invoice {invoice_number} for Customer {customer_id} dated {date}"
    body = f"Attached is the sorted invoice {invoice_number} for customer {customer_id} dated {date}."

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(file_path)}")
    message.attach(part)

    try:
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(sender_email, app_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print(f"Email sent successfully to {receiver_email}!")
    except Exception as e:
        print(f"Failed to send email: {e}")
    finally:
        server.quit()

import fitz
import tempfile

def convert_pdf_to_images(pdf_path):
    doc = fitz.open(pdf_path)
    image_paths = []
    try:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Increase resolution
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                pix.save(tmp.name)
                image_paths.append(tmp.name)
            print(f"Converted page {i+1} to image: {tmp.name}")
    finally:
        doc.close()
    return image_paths

# def extract_invoice_data(texts):
#     invoice_num = find_text_near_keyphrase(texts, "Invoice", "below", 10)
#     customer_num = find_text_near_keyphrase(texts, "Account", "below", 100)
#     date_num = find_text_near_keyphrase(texts, "Date", "below", 10)
    
#     return invoice_num, customer_num, date_num

import logging

# Update the log folder path
log_folder = "./Logs"
os.makedirs(log_folder, exist_ok=True)
log_file = os.path.join(log_folder, f"ocr_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(filename=log_file, level=logging.DEBUG, format='%(asctime)s - %(message)s')

def process_invoice_image(image_path, manual_mode=False):
    """Modified process_invoice_image to include better route number logging."""
    logging.info(f"\nDEBUG: Processing image: {image_path}")
    processed_successfully = False
    
    if not manual_mode:
        image = cv2.imread(image_path)
        texts = extract_text_and_positions(image)
        
        # Debug print all detected text
        print("\nAll detected text elements:")
        for i, text in enumerate(texts[1:]):  # Skip first element which contains all text
            print(f"Text {i}: '{text.description}'")
            box = text.bounding_poly.vertices
            print(f"  Position: ({box[0].x}, {box[0].y}) to ({box[2].x}, {box[2].y})")
        
        logging.info("\nStarting field extraction...")
        invoice_num = find_text_near_keyphrase(texts, "INVOICE NO", "below", 50)
        customer_num = find_text_near_keyphrase(texts, "ACCOUNT NO", "below", 50)
        date_num = find_text_near_keyphrase(texts, "INVOICE DATE", "below", 50)
        
        logging.info("\nAttempting to find route number...")
        route_num = find_text_near_keyphrase(texts, "DRIVER/RT", "below", 50)
        if not route_num:
            logging.info("First attempt failed, trying alternative format...")
            route_num = find_text_near_keyphrase(texts, "Driver/Route #", "below", 50)
        
        logging.info("\nExtracted field values:")
        logging.info(f"Invoice Number: {invoice_num}")
        logging.info(f"Customer Number: {customer_num}")
        logging.info(f"Date: {date_num}")
        logging.info(f"Route Number: {route_num}")
        
        if (invoice_num and is_six_alphanumeric(invoice_num) and
            customer_num and is_six_alphanumeric(customer_num) and
            date_num and is_date_format(date_num)):
            
            filename_base = f"{customer_num}_{invoice_num}"
            if route_num:
                filename_base += f"_{route_num}"
            
            copy_image_to_sorted_folder(image_path, filename_base, customer_num, date_num)
            
            # Add to report tracking
            report_manager.add_invoice(invoice_num, route_num)
            
            processed_successfully = True
            try:
                os.remove(image_path)
                logging.info(f"Deleted original file: {image_path}")
            except Exception as e:
                logging.error(f"Error deleting original file {image_path}: {e}")
    
    return processed_successfully

# def process_invoices(invoice_folder):
#     """Process all invoices in the folder."""
#     print(f"DEBUG: Processing invoices in folder: {invoice_folder}")
#     if not os.path.exists(invoice_folder):
#         print(f"ERROR: Invoice folder does not exist: {invoice_folder}")
#         return

#     # Store failed invoices for manual processing
#     failed_invoices = []

#     # First pass: Automatic processing
#     print("\nStarting automatic processing of invoices...")
#     for filename in os.listdir(invoice_folder):
#         file_path = os.path.join(invoice_folder, filename)
        
#         if filename.lower().endswith('.pdf'):
#             print(f"Processing PDF file: {filename}")
#             try:
#                 # Convert PDF to images and process each page
#                 doc = fitz.open(file_path)
#                 for page_num in range(len(doc)):
#                     pix = doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2))
#                     temp_image_path = save_pixmap_with_retry(pix)
#                     try:
#                         if not process_invoice_image(temp_image_path):
#                             # Store the failed page for manual processing
#                             failed_invoices.append(temp_image_path)
#                     except Exception as e:
#                         print(f"Error processing PDF page {page_num + 1}: {str(e)}")
#                 doc.close()
#                 # Delete original PDF after processing all pages
#                 os.remove(file_path)
#                 print(f"Deleted original PDF: {file_path}")
#             except Exception as e:
#                 print(f"Error processing PDF {filename}: {str(e)}")
        
#         elif filename.lower().endswith(('.png', '.jpg', '.jpeg')):
#             print(f"Processing image file: {filename}")
#             try:
#                 if not process_invoice_image(file_path):
#                     failed_invoices.append(file_path)
#             except Exception as e:
#                 print(f"Error processing image {filename}: {str(e)}")
        
#         else:
#             print(f"Skipped non-invoice file: {filename}")

#     # Second pass: Manual processing of failed invoices
#     if failed_invoices:
#         print(f"\nFound {len(failed_invoices)} invoices that need manual processing.")
#         print("Starting manual processing...\n")
        
#         for failed_invoice in failed_invoices:
#             try:
#                 process_invoice_image(failed_invoice, manual_mode=True)
#             except Exception as e:
#                 print(f"Error during manual processing of {os.path.basename(failed_invoice)}: {str(e)}")
            
#             # Clean up temporary files (PDF pages)
#             if failed_invoice.startswith(tempfile.gettempdir()):
#                 try:
#                     if os.path.exists(failed_invoice):  # Check if file still exists
#                         os.remove(failed_invoice)
#                         print(f"Removed temporary file: {failed_invoice}")
#                 except Exception as e:
#                     print(f"Error removing temporary file: {str(e)}")
    
#     else:
#         print("\nAll invoices were processed automatically. No manual input needed.")

#     print("\nInvoice processing completed.")

#     # Do not delete the original invoice after processing
#     # os.remove(invoice_path)
#     # print(f"DEBUG: Deleted original invoice: {invoice_path}")

#     print("="*40)  # Separator for readability

import tempfile
import time
import random
import string

def generate_unique_filename(suffix='.png'):
    timestamp = int(time.time() * 1000)
    random_string = ''.join(random.choices(string.ascii_lowercase + string.digits, k=6))
    return f"temp_{timestamp}_{random_string}{suffix}"

def save_pixmap_with_retry(pix, max_retries=5):
    for attempt in range(max_retries):
        try:
            temp_dir = tempfile.gettempdir()
            temp_filename = generate_unique_filename()
            temp_path = os.path.join(temp_dir, temp_filename)
            pix.save(temp_path)
            return temp_path
        except Exception as e:
            print(f"Error saving pixmap (attempt {attempt + 1}): {str(e)}")
            if attempt == max_retries - 1:
                raise
            time.sleep(0.5)  # Wait for half a second before retrying

def process_invoices(invoice_folder):
    """Process all invoices in the folder."""
    print(f"DEBUG: Processing invoices in folder: {invoice_folder}")
    if not os.path.exists(invoice_folder):
        print(f"ERROR: Invoice folder does not exist: {invoice_folder}")
        return

    # Store failed invoices for manual processing
    failed_invoices = []

    # First pass: Automatic processing
    print("\nStarting automatic processing of invoices...")
    for filename in os.listdir(invoice_folder):
        file_path = os.path.join(invoice_folder, filename)
        
        if filename.lower().endswith('.pdf'):
            print(f"Processing PDF file: {filename}")
            try:
                # Convert PDF to images and process each page
                doc = fitz.open(file_path)
                for page_num in range(len(doc)):
                    pix = doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2))
                    temp_image_path = save_pixmap_with_retry(pix)
                    try:
                        if not process_invoice_image(temp_image_path):
                            # Store the failed page for manual processing
                            failed_invoices.append(temp_image_path)
                        else:
                            # Remove temp file if successful
                            os.remove(temp_image_path)
                    except Exception as e:
                        print(f"Error processing PDF page {page_num + 1}: {str(e)}")
                doc.close()
            except Exception as e:
                print(f"Error processing PDF {filename}: {str(e)}")
        
        elif filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            print(f"Processing image file: {filename}")
            try:
                if not process_invoice_image(file_path):
                    failed_invoices.append(file_path)
            except Exception as e:
                print(f"Error processing image {filename}: {str(e)}")
        
        else:
            print(f"Skipped non-invoice file: {filename}")

    # Second pass: Manual processing of failed invoices
    if failed_invoices:
        print(f"\nFound {len(failed_invoices)} invoices that need manual processing.")
        print("Starting manual processing...\n")
        
        for failed_invoice in failed_invoices:
            try:
                process_invoice_image(failed_invoice, manual_mode=True)
            except Exception as e:
                print(f"Error during manual processing of {os.path.basename(failed_invoice)}: {str(e)}")
            
            # Clean up temporary files
            if failed_invoice.startswith(tempfile.gettempdir()):
                try:
                    os.remove(failed_invoice)
                    print(f"Removed temporary file: {failed_invoice}")
                except Exception as e:
                    print(f"Error removing temporary file: {str(e)}")
    
    else:
        print("\nAll invoices were processed automatically. No manual input needed.")

    print("\nInvoice processing completed.")

    # Generate report for this session's invoices
    report_manager.generate_report()

def process_pdf_invoice(pdf_path):
    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc):
            print(f"Processing page {i+1} of PDF: {pdf_path}")
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            try:
                temp_image_path = save_pixmap_with_retry(pix)
                print(f"Saved page {i+1} as temporary image: {temp_image_path}")
                if process_invoice_image(temp_image_path):
                    print(f"Successfully processed page {i+1}")
                else:
                    print(f"Failed to process page {i+1}")
            finally:
                if 'temp_image_path' in locals():
                    try:
                        os.remove(temp_image_path)
                        print(f"Removed temporary image: {temp_image_path}")
                    except Exception as e:
                        print(f"Error removing temporary file: {str(e)}")
    finally:
        doc.close()

def add_customer_email(excel_file, customer_id, email):
    df = pd.read_excel(excel_file, header=None)  # No header row
    new_entry = pd.DataFrame({0: [customer_id], 1: [email]})
    df = pd.concat([df, new_entry], ignore_index=True)
    df.to_excel(excel_file, index=False, header=False)
    print(f"Added {customer_id} with email {email} to the list.")

def update_customer_email(excel_file, customer_id, new_email):
    df = pd.read_excel(excel_file, header=None)  # No header row
    df.loc[df[0] == customer_id, 1] = new_email
    df.to_excel(excel_file, index=False, header=False)
    print(f"Updated {customer_id}'s email to {new_email}.")

def remove_customer_email(excel_file, customer_id):
    df = pd.read_excel(excel_file, header=None)  # No header row
    df = df[df[0] != customer_id]
    df.to_excel(excel_file, index=False, header=False)
    print(f"Removed {customer_id} from the list.")

if __name__ == "__main__":
    print("DEBUG: Starting invoice processing")

    # Initialize report manager
    report_manager = ReportManager()

    process_invoices(invoice_folder)
    check_inactive_customers()  # Add this line


    print("Program execution completed.")

    input("Press Enter to exit...")
    
    print("DEBUG: Finished invoice processing")
