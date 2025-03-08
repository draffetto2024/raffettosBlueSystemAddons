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

import numpy as np
import pandas as pd
import sqlite3




# Simplified path definitions using relative paths
# service_account_key = "./caramel-compass-429017-h3-c2d4e157e809.json"
service_account_key = "./invoicescanner-446017-e1844c3df524.json"
invoice_folder = "./Invoices/InvoicePictures"
destination_base_folder = "./Invoices/SortedInvoices"
unsorted_base_folder = "./Invoices/UnsortedInvoices"
customer_emails_file = "./customer_emails.xlsx"

# Set the Google Application Credentials environment variable
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = service_account_key

# Initialize the Vision API client
client = vision.ImageAnnotatorClient()


# Email details
# sender_email = "gingoso2@gmail.com"
sender_email = "derekraffetto@gmail.com"
# app_password = "soiz avjw bdtu hmtn"
app_password = "sqha dkre bten tgpe"


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
    user_input = {'customer_id': None, 'invoice_num': None, 'date': None}
    
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
    
    # Error label
    error_label = ttk.Label(input_frame, text="", foreground="red")
    error_label.grid(row=4, column=0, columnspan=2)
    
    def focus_next_empty(event=None):
        """Focus next empty field or submit if all are filled."""
        current = root.focus_get()
        entries = [customer_id_entry, invoice_num_entry, date_entry]
        
        if current in entries:
            current_idx = entries.index(current)
            # Check remaining fields
            for idx in range(current_idx + 1, len(entries)):
                if not entries[idx].get().strip():
                    entries[idx].focus()
                    return "break"
            # If we get here and current field is not empty, try to submit
            if current.get().strip():
                submit_wrapper()
        return "break"
    
    def submit_wrapper(event=None):
        customer_id = customer_id_entry.get().strip()
        invoice_num = invoice_num_entry.get().strip()
        date = date_entry.get().strip()
        
        print(f"Validating input - Customer ID: {customer_id}, Invoice: {invoice_num}, Date: {date}")
        
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
            # Date was in MMDDYY format and was converted
            date = date_result
        elif not date_result:
            error_label.config(text="Invalid Date format - use MM/DD/YY or MMDDYY")
            date_entry.focus()
            date_entry.selection_range(0, tk.END)
            return
        
        # If all validations pass, store the values and close window
        user_input['customer_id'] = customer_id
        user_input['invoice_num'] = invoice_num
        user_input['date'] = date
        print(f"DEBUG: All validations passed. Input values: {user_input}")
        root.quit()
        root.destroy()
    
    # Bind Enter key for each entry
    customer_id_entry.bind('<Return>', focus_next_empty)
    invoice_num_entry.bind('<Return>', focus_next_empty)
    date_entry.bind('<Return>', focus_next_empty)
    
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
    
    print(f"Returning user input: {user_input}")
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
    logging.info(f"\nSearching for text near '{keyphrase}' in '{position}' position:")
    
    keyphrase_parts = keyphrase.lower().split()
    keyphrase_texts = []
    
    # Find the keyphrase first
    for i, text in enumerate(texts):
        if keyphrase_parts[0] in text.description.lower():
            keyphrase_texts = [text]
            current_box = text.bounding_poly.vertices
            
            # Look for remaining parts within the threshold
            for part in keyphrase_parts[1:]:
                found = False
                for next_text in texts[i+1:]:
                    next_box = next_text.bounding_poly.vertices
                    if part in next_text.description.lower() and is_within_threshold(current_box, next_box, threshold):
                        keyphrase_texts.append(next_text)
                        current_box = combine_boxes(current_box, next_box)
                        found = True
                        break
                if not found:
                    break
            
            if len(keyphrase_texts) == len(keyphrase_parts):
                break
    
    if len(keyphrase_texts) != len(keyphrase_parts):
        logging.info(f"Complete keyphrase '{keyphrase}' not found within threshold.")
        return None
    
    # Combine bounding boxes of keyphrase parts
    keyphrase_box = keyphrase_texts[0].bounding_poly.vertices
    for text in keyphrase_texts[1:]:
        keyphrase_box = combine_boxes(keyphrase_box, text.bounding_poly.vertices)
    
    logging.info(f"Keyphrase '{' '.join(text.description for text in keyphrase_texts)}' bounding box:")
    logging.info(f"  Top-left: ({keyphrase_box[0].x}, {keyphrase_box[0].y})")
    logging.info(f"  Bottom-right: ({keyphrase_box[2].x}, {keyphrase_box[2].y})")
    
    # Store valid matches to pick the best one
    valid_matches = []
    
    # Now search for text near the keyphrase
    for adjacent_text in texts:
        if adjacent_text in keyphrase_texts:
            continue  # Skip the keyphrase itself
            
        adjacent_box = adjacent_text.bounding_poly.vertices
        
        # Skip if the bounding box is too large (likely entire document)
        box_width = adjacent_box[2].x - adjacent_box[0].x
        box_height = adjacent_box[2].y - adjacent_box[0].y
        if box_width > 500 or box_height > 500:  # Adjust these thresholds as needed
            continue
            
        if is_in_position(keyphrase_box, adjacent_box, position, threshold):
            text_value = adjacent_text.description.strip()
            
            # Validate based on keyphrase type
            is_valid = False
            if "INVOICE NO" in keyphrase.upper():
                is_valid = bool(re.match(r'^[A-Za-z0-9]{6}$', text_value))
            elif "ACCOUNT NO" in keyphrase.upper():
                is_valid = bool(re.match(r'^[A-Za-z0-9]{6}$', text_value))
            elif "INVOICE DATE" in keyphrase.upper():
                is_valid = bool(re.match(r'^(0[1-9]|1[0-2])/(0[1-9]|[12][0-9]|3[01])/\d{2}$', text_value))
            
            if is_valid:
                distance = calculate_distance(keyphrase_box, adjacent_box)
                valid_matches.append((text_value, distance, adjacent_box))
    
    if valid_matches:
        # Sort by distance and take the closest match
        valid_matches.sort(key=lambda x: x[1])
        best_match = valid_matches[0]
        
        logging.info(f"\nMATCH FOUND: '{best_match[0]}'")
        logging.info(f"Match bounding box:")
        logging.info(f"  Top-left: ({best_match[2][0].x}, {best_match[2][0].y})")
        logging.info(f"  Bottom-right: ({best_match[2][2].x}, {best_match[2][2].y})")
        return best_match[0]
    
    logging.info(f"No valid matching text found near '{keyphrase}'")
    return None

def is_within_threshold(box1, box2, threshold):
    # Check if box2 is within threshold distance of box1
    return (abs(box2[0].x - box1[2].x) < threshold and
            abs(box2[0].y - box1[0].y) < threshold)

def combine_boxes(box1, box2):
    # Combine two bounding boxes
    return [
        types.Vertex(x=min(box1[0].x, box2[0].x), y=min(box1[0].y, box2[0].y)),
        types.Vertex(x=max(box1[1].x, box2[1].x), y=min(box1[1].y, box2[1].y)),
        types.Vertex(x=max(box1[2].x, box2[2].x), y=max(box1[2].y, box2[2].y)),
        types.Vertex(x=min(box1[3].x, box2[3].x), y=max(box1[3].y, box2[3].y))
    ]




def print_all_text_elements(texts):
    print("\nAll detected text elements:")
    for i, text in enumerate(texts):
        print(f"Text {i + 1}: '{text.description}'")
        box = text.bounding_poly.vertices
        print(f"  Top-left: ({box[0].x}, {box[0].y})")
        print(f"  Bottom-right: ({box[2].x}, {box[2].y})")
        print()


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

def calculate_distance(box1, box2):
    """Calculate the minimum distance between two bounding boxes."""
    # Calculate center points
    center1_x = (box1[0].x + box1[2].x) / 2
    center1_y = (box1[0].y + box1[2].y) / 2
    center2_x = (box2[0].x + box2[2].x) / 2
    center2_y = (box2[0].y + box2[2].y) / 2
    
    # Calculate Euclidean distance
    return ((center1_x - center2_x) ** 2 + (center1_y - center2_y) ** 2) ** 0.5

def copy_image_to_sorted_folder(image_path, six_digit_number, customer_id, date):
    year, month, day = extract_year(date), extract_month(date), extract_day(date)
    destination_folder = os.path.join(destination_base_folder, year, month, day)
    os.makedirs(destination_folder, exist_ok=True)
    
    base_name = f"{customer_id}_{six_digit_number}"
    extension = os.path.splitext(image_path)[1]
    
    # Try original filename first
    destination_path = os.path.join(destination_folder, f"{base_name}{extension}")
    counter = 1
    
    # If file exists, add incremental number until we find unused name
    while os.path.exists(destination_path):
        destination_path = os.path.join(destination_folder, f"{base_name}_{counter}{extension}")
        counter += 1
    
    try:
        shutil.copy2(image_path, destination_path)
        print(f"Copied and renamed image to: {destination_path}")
        customer_tracker.update_customer_activity(customer_id, date)
    except Exception as e:
        print(f"DEBUG: Error copying file: {e}")
        
    send_email_with_attachment(destination_path, customer_id, six_digit_number, date)

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
    """Process a single invoice image."""
    logging.info(f"DEBUG: Processing image: {image_path}")
    processed_successfully = False
    
    if manual_mode:
        print(f"\nProcessing failed invoice: {os.path.basename(image_path)}")
        user_input = display_image_and_get_input(image_path)
        
        if all(user_input.values()):
            logging.info(f"DEBUG: Manual input successful, copying image to sorted folder")
            print(f"Processing manual input: {user_input}")

            copy_image_to_sorted_folder(
                image_path,
                user_input['invoice_num'],
                user_input['customer_id'],
                user_input['date']
            )
            processed_successfully = True
        else:
            logging.info(f"DEBUG: Manual input cancelled or invalid, copying to unsorted folder")
            print("Manual input failed or cancelled, moving to unsorted folder")
            copy_image_to_unsorted_folder(image_path)
    else:
        image = cv2.imread(image_path)
        texts = extract_text_and_positions(image)
        
        invoice_num = None
        customer_num = None
        date_num = None
        
        # Try increasingly larger thresholds for each field
        for threshold in range(5, 501, 5):
            if not invoice_num or not is_six_alphanumeric(invoice_num):
                invoice_num = find_text_near_keyphrase(texts, "INVOICE NO", "below", threshold)
            
            if not customer_num or not is_six_alphanumeric(customer_num):
                customer_num = find_text_near_keyphrase(texts, "ACCOUNT NO", "below", threshold)
            
            if not date_num or not is_date_format(date_num):
                date_num = find_text_near_keyphrase(texts, "INVOICE DATE", "below", threshold)
            
            if (invoice_num and is_six_alphanumeric(invoice_num) and
                customer_num and is_six_alphanumeric(customer_num) and
                date_num and is_date_format(date_num)):
                break
                
        if (invoice_num and is_six_alphanumeric(invoice_num) and
            customer_num and is_six_alphanumeric(customer_num) and
            date_num and is_date_format(date_num)):
            logging.info(f"DEBUG: Automatic detection successful, copying image to sorted folder")
            
            # Add this line to extract and display the invoice data
            process_invoice_data_extraction(image_path, invoice_num, customer_num, date_num)
            
            copy_image_to_sorted_folder(image_path, invoice_num, customer_num, date_num)
            processed_successfully = True
        else:
            logging.info(f"DEBUG: Automatic detection failed")
    
    if processed_successfully:
        try:
            os.remove(image_path)
            print(f"Deleted original file: {image_path}")
        except Exception as e:
            print(f"Error deleting original file {image_path}: {e}")
    
    return processed_successfully

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

    failed_invoices = []

    print("\nStarting automatic processing of invoices...")
    for filename in os.listdir(invoice_folder):
        file_path = os.path.join(invoice_folder, filename)
        
        if filename.lower().endswith('.pdf'):
            print(f"Processing PDF file: {filename}")
            try:
                doc = fitz.open(file_path)
                all_pages_processed = True
                temp_files = []
                
                for page_num in range(len(doc)):
                    pix = doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2))
                    temp_image_path = save_pixmap_with_retry(pix)
                    temp_files.append(temp_image_path)
                    try:
                        if not process_invoice_image(temp_image_path):
                            failed_invoices.append(temp_image_path)
                            all_pages_processed = False
                        else:
                            os.remove(temp_image_path)
                    except Exception as e:
                        print(f"Error processing PDF page {page_num + 1}: {str(e)}")
                        all_pages_processed = False
                
                doc.close()
                
                # Only delete original PDF if all pages were processed successfully
                if all_pages_processed:
                    os.remove(file_path)
                    print(f"Deleted original PDF: {file_path}")
                    
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

    if failed_invoices:
        print(f"\nFound {len(failed_invoices)} invoices that need manual processing.")
        print("Starting manual processing...\n")
        
        for failed_invoice in failed_invoices:
            try:
                manual_result = process_invoice_image(failed_invoice, manual_mode=True)
                original_file = failed_invoice
                
                # If it's a temp file from PDF, find the original PDF file
                if failed_invoice.startswith(tempfile.gettempdir()):
                    # Clean up the temp file
                    try:
                        os.remove(failed_invoice)
                        print(f"Removed temporary file: {failed_invoice}")
                    except Exception as e:
                        print(f"Error removing temporary file: {str(e)}")
                else:
                    # For direct image files, delete original after processing
                    try:
                        if os.path.exists(failed_invoice):
                            os.remove(failed_invoice)
                            print(f"Deleted original file after manual processing: {failed_invoice}")
                    except Exception as e:
                        print(f"Error deleting original file: {str(e)}")
                        
            except Exception as e:
                print(f"Error during manual processing of {os.path.basename(failed_invoice)}: {str(e)}")
    
    else:
        print("\nAll invoices were processed automatically. No manual input needed.")

    print("\nInvoice processing completed.")

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



#********************************************************************************************#********************************************************************************************
#********************************************************************************************#********************************************************************************************
#********************************************************************************************#********************************************************************************************

import cv2

def print_tabular_data(texts):
    """
    Print the OCR data in a tabular format to visualize the detected grid.
    This is a simplified visualization that doesn't require drawing on the image.
    """
    # Skip the first text annotation which is the entire text
    texts = texts[1:] if len(texts) > 1 else texts
    
    print("\n===== DETECTED TEXT ELEMENTS IN TABULAR FORM =====")
    
    # Create a dictionary to store text by row position
    rows = defaultdict(list)
    row_threshold = 10  # Pixels within which elements are considered the same row
    
    for text in texts:
        content = text.description.strip()
        if not content:
            continue
            
        box = text.bounding_poly.vertices
        center_y = (box[0].y + box[2].y) / 2
        center_x = (box[0].x + box[2].x) / 2
        
        # Find or create appropriate row
        matched_row = None
        for row_y in rows.keys():
            if abs(center_y - row_y) < row_threshold:
                matched_row = row_y
                break
                
        if matched_row is None:
            matched_row = center_y
            
        rows[matched_row].append((content, center_x))
    
    # Sort rows by Y position (top to bottom)
    sorted_rows = sorted(rows.items(), key=lambda x: x[0])
    
    # Print each row
    for i, (y, elements) in enumerate(sorted_rows):
        # Sort elements by X position (left to right)
        sorted_elements = sorted(elements, key=lambda x: x[1])
        elements_text = [f"{text} ({x:.1f})" for text, x in sorted_elements]
        print(f"Row {i+1} (y={y:.1f}): {' | '.join(elements_text)}")

def extract_invoice_items(texts):
    """
    Extract line items from Raffetto's invoices which have a consistent format.
    This function uses hardcoded positions based on the known invoice layout.
    """
    print("\nDEBUG: Starting line item extraction for Raffetto's invoice format...")
    
    # Hardcoded column positions based on Raffetto's invoice layout
    column_positions = {
        "item_no": 65.0,      # Position of item numbers like 01LASW, P1FETT
        "qty_ordered": 155.0, # Position of ordered quantities
        "qty_shipped": 215.0, # Position of shipped quantities
        "unit": 290.0,        # Position of unit (CS)
        "description": 570.0, # Position of item descriptions
        "case_pack": 910.0,   # Position of case pack details
        "unit_price": 1020.0, # Position of unit prices
        "extension": 1135.0   # Position of extension (total) prices
    }
    
    # Hardcoded y-coordinate range for the header row
    header_y_range = (340.0, 370.0)
    
    # Hardcoded y-coordinate range for the data rows (typical line items area)
    data_y_min = 380.0
    data_y_max = 900.0  # Set this to exclude footer content
    
    print(f"DEBUG: Using hardcoded column positions: {column_positions}")
    print(f"DEBUG: Using hardcoded data row range: y={data_y_min} to y={data_y_max}")
    
    # Group text elements by their y-coordinate to form rows
    rows = defaultdict(list)
    row_threshold = 20  # Larger threshold to group related elements
    
    for text in texts[1:]:  # Skip the first text which is entire document text
        content = text.description.strip()
        if not content:
            continue
            
        box = text.bounding_poly.vertices
        center_y = (box[0].y + box[2].y) / 2
        center_x = (box[0].x + box[2].x) / 2
        
        # Only consider elements in the data rows area
        if data_y_min <= center_y <= data_y_max:
            # Find or create appropriate row
            matched_row = None
            for row_y in rows.keys():
                if abs(center_y - row_y) < row_threshold:
                    matched_row = row_y
                    break
                    
            if matched_row is None:
                matched_row = center_y
                
            rows[matched_row].append((content, center_x))
    
    # Print the rows we found for debugging
    print(f"\nDEBUG: Found {len(rows)} potential data rows in the expected region")
    for row_idx, (row_y, elements) in enumerate(sorted(rows.items(), key=lambda x: x[0])):
        print(f"DEBUG: Row {row_idx+1} (y={row_y:.1f}): {[e[0] for e in elements]}")
    
    # Process rows into line items
    line_items = []
    column_threshold = 100  # Wider threshold for matching text to columns
    
    for row_y, elements in sorted(rows.items(), key=lambda x: x[0]):
        # Skip rows with too few elements (likely not actual line items)
        if len(elements) < 3:
            print(f"DEBUG: Skipping row at y={row_y:.1f} - only {len(elements)} elements")
            continue
        
        # Initialize item with empty values
        item = {
            "item_no": "",
            "qty_ordered": "",
            "qty_shipped": "",
            "unit": "",
            "description": "",
            "case_pack": "",
            "unit_price": "",
            "extension": ""
        }
        
        # Store elements for each column separately for better processing
        description_elements = []
        
        # Assign elements to columns based on X position
        for content, x_pos in elements:
            # Find the closest column
            best_match = None
            best_distance = float('inf')
            
            for col, col_x in column_positions.items():
                distance = abs(x_pos - col_x)
                if distance < best_distance and distance < column_threshold:
                    best_distance = distance
                    best_match = col
            
            if best_match:
                # Special handling for description column - collect all elements
                if best_match == "description":
                    description_elements.append(content)
                # For other columns, just set the value
                else:
                    item[best_match] = content
        
        # Combine all description elements into a single string
        if description_elements:
            item["description"] = " ".join(description_elements)
        
        # Clean up numeric fields
        for field in ["qty_ordered", "qty_shipped", "unit_price", "extension"]:
            if item[field]:
                # Remove any non-numeric characters except decimal points
                item[field] = re.sub(r'[^\d.]', '', item[field])
        
        # Check if we have meaningful data
        has_item_code = bool(item["item_no"])
        has_description = bool(item["description"])
        has_price = bool(item["unit_price"] or item["extension"])
        
        # Only add items that have at least an item code, description, or price
        if has_item_code or has_description or has_price:
            print(f"DEBUG: Found valid line item: {item}")
            line_items.append(item)
        else:
            print(f"DEBUG: Discarded row - no meaningful data")
    
    print(f"DEBUG: Successfully extracted {len(line_items)} line items")
    return line_items

def create_invoice_database():
    """Create SQLite database for invoice data with proper error handling."""
    print("DEBUG: Creating/checking invoice database...")
    
    try:
        # Connect to database (will create it if it doesn't exist)
        conn = sqlite3.connect('invoice_database.db')
        cursor = conn.cursor()
        
        # Check if tables already exist
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='invoices'")
        invoices_exists = cursor.fetchone() is not None
        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='invoice_items'")
        items_exists = cursor.fetchone() is not None
        
        # For invoices table
        if invoices_exists:
            print("DEBUG: 'invoices' table already exists, checking columns...")
            cursor.execute("PRAGMA table_info(invoices)")
            columns = [row[1] for row in cursor.fetchall()]
            print(f"DEBUG: Existing 'invoices' columns: {columns}")
            
            # If table exists but missing expected columns, drop and recreate
            expected_columns = ['invoice_id', 'customer_id', 'invoice_date', 'total_amount', 'shipping_amount', 'processed_date']
            if not all(col in columns for col in expected_columns):
                print("DEBUG: 'invoices' table missing required columns, recreating...")
                cursor.execute("DROP TABLE invoices")
                invoices_exists = False
        
        if not invoices_exists:
            print("DEBUG: Creating 'invoices' table...")
            cursor.execute('''
            CREATE TABLE invoices (
                invoice_id TEXT PRIMARY KEY,
                customer_id TEXT,
                invoice_date TEXT,
                total_amount REAL,
                shipping_amount REAL,
                processed_date TEXT
            )
            ''')
            print("DEBUG: 'invoices' table created successfully")
        
        # For invoice_items table
        if items_exists:
            print("DEBUG: 'invoice_items' table already exists, checking columns...")
            cursor.execute("PRAGMA table_info(invoice_items)")
            columns = [row[1] for row in cursor.fetchall()]
            print(f"DEBUG: Existing 'invoice_items' columns: {columns}")
            
            # If table exists but missing expected columns, drop and recreate
            expected_columns = ['id', 'invoice_id', 'item_no', 'qty_ordered', 'qty_shipped', 'unit', 'description', 'case_pack', 'unit_price', 'extension']
            if not all(col in columns for col in expected_columns):
                print("DEBUG: 'invoice_items' table missing required columns, recreating...")
                cursor.execute("DROP TABLE invoice_items")
                items_exists = False
        
        if not items_exists:
            print("DEBUG: Creating 'invoice_items' table...")
            cursor.execute('''
            CREATE TABLE invoice_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_id TEXT,
                item_no TEXT,
                qty_ordered INTEGER,
                qty_shipped INTEGER,
                unit TEXT,
                description TEXT,
                case_pack TEXT,
                unit_price REAL,
                extension REAL,
                FOREIGN KEY (invoice_id) REFERENCES invoices (invoice_id)
            )
            ''')
            print("DEBUG: 'invoice_items' table created successfully")
        
        conn.commit()
        print("DEBUG: Database setup complete")
        return True
        
    except sqlite3.Error as e:
        print(f"DEBUG: SQLite error during database creation: {e}")
        return False
        
    finally:
        if 'conn' in locals():
            conn.close()

def save_invoice_to_database(invoice_num, customer_id, date, line_items):
    """Save invoice and line items to SQLite database with better error handling."""
    print("\nDEBUG: Starting database save operation...")
    
    if not invoice_num or not customer_id or not date or not line_items:
        print("ERROR: Missing required invoice data, not saving to database.")
        return False
    
    # Create database if it doesn't exist or needs updating
    if not create_invoice_database():
        print("ERROR: Failed to create or verify database, cannot save data.")
        return False
    
    try:
        # Connect to database
        conn = sqlite3.connect('invoice_database.db')
        cursor = conn.cursor()
        
        # Calculate total amount from line items
        total_amount = 0
        for item in line_items:
            if item["extension"]:
                try:
                    total_amount += float(item["extension"])
                except ValueError:
                    print(f"DEBUG: Invalid extension value: '{item['extension']}', skipping in total")
        
        print(f"DEBUG: Calculated total amount: ${total_amount:.2f}")
        
        # Check if invoice already exists
        cursor.execute("SELECT invoice_id FROM invoices WHERE invoice_id = ?", (invoice_num,))
        if cursor.fetchone():
            print(f"INFO: Invoice {invoice_num} already exists in database, skipping.")
            conn.close()
            return False
        
        # Insert invoice record
        print(f"DEBUG: Adding invoice metadata to database...")
        cursor.execute(
            "INSERT INTO invoices (invoice_id, customer_id, invoice_date, total_amount, processed_date) VALUES (?, ?, ?, ?, ?)",
            (invoice_num, customer_id, date, total_amount, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        )
        
        # Insert line items
        print(f"DEBUG: Adding {len(line_items)} line items to database...")
        for i, item in enumerate(line_items, 1):
            # Print progress for large invoices
            if i % 5 == 0 or i == len(line_items):
                print(f"  DEBUG: Progress: {i}/{len(line_items)} items processed")
                
            try:
                # Convert values to appropriate types
                qty_ordered = int(float(item["qty_ordered"])) if item["qty_ordered"] else None
                qty_shipped = int(float(item["qty_shipped"])) if item["qty_shipped"] else None
                unit_price = float(item["unit_price"]) if item["unit_price"] else None
                extension = float(item["extension"]) if item["extension"] else None
                
                cursor.execute(
                    """INSERT INTO invoice_items 
                       (invoice_id, item_no, qty_ordered, qty_shipped, unit, description, case_pack, unit_price, extension)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        invoice_num,
                        item["item_no"],
                        qty_ordered,
                        qty_shipped,
                        item["unit"],
                        item["description"],
                        item["case_pack"],
                        unit_price,
                        extension
                    )
                )
                print(f"  DEBUG: Added line item {i}: {item['item_no']} - {item['description']}")
                
            except (ValueError, TypeError) as e:
                print(f"  WARNING: Error converting values for line item {i}: {e}")
                # Still try to insert with None values
                cursor.execute(
                    """INSERT INTO invoice_items 
                       (invoice_id, item_no, description, case_pack)
                       VALUES (?, ?, ?, ?)""",
                    (
                        invoice_num,
                        item["item_no"],
                        item["description"],
                        item["case_pack"]
                    )
                )
                print(f"  DEBUG: Added partial line item {i} with basic info only")
        
        conn.commit()
        print(f"SUCCESS: Invoice {invoice_num} with {len(line_items)} line items saved to database.")
        return True
        
    except sqlite3.Error as e:
        if 'conn' in locals():
            conn.rollback()
        print(f"ERROR: Database error during save operation: {e}")
        import traceback
        traceback.print_exc()
        return False
    except Exception as e:
        if 'conn' in locals():
            conn.rollback()
        print(f"ERROR: Unexpected error during save operation: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        if 'conn' in locals():
            conn.close()

def process_invoice_data_extraction(image_path, invoice_num, customer_id, date):
    """
    Process an invoice image to extract line items and save them to the database.
    This function should be called after the invoice metadata is extracted.
    
    Parameters:
    - image_path: Path to the invoice image
    - invoice_num: The invoice number (already extracted)
    - customer_id: The customer ID (already extracted)
    - date: The invoice date (already extracted)
    
    Returns:
    - True if data was successfully extracted and saved, False otherwise
    """
    try:
        print("\n" + "="*80)
        print(f"PROCESSING INVOICE DATA: #{invoice_num} | Customer: {customer_id} | Date: {date}")
        print("="*80)
        
        # Read the image
        image = cv2.imread(image_path)
        if image is None:
            print(f"Failed to read image: {image_path}")
            return False
        
        # Extract text annotations using Google Vision API
        texts = extract_text_and_positions(image)
        if not texts:
            print(f"No text extracted from image: {image_path}")
            return False
        
        # Print tabular representation of the data for debugging (optional)
        # print_tabular_data(texts)
        
        # Extract line items
        line_items = extract_invoice_items(texts)
        
        # Prepare data for display
        if line_items:
            # First, print the invoice metadata
            print("\n" + "-"*80)
            print("INVOICE METADATA:")
            print(f"Invoice Number: {invoice_num}")
            print(f"Customer ID:    {customer_id}")
            print(f"Invoice Date:   {date}")
            print("-"*80)
            
            # Create a DataFrame for better display
            df = pd.DataFrame(line_items)
            
            # Reorder columns for better readability
            columns_order = ["item_no", "description", "qty_ordered", "qty_shipped", 
                           "unit", "case_pack", "unit_price", "extension"]
            
            # Only include columns that exist
            display_columns = [col for col in columns_order if col in df.columns]
            df = df[display_columns]
            
            # Rename columns for display
            column_names = {
                "item_no": "Item #",
                "description": "Description",
                "qty_ordered": "Qty Ord",
                "qty_shipped": "Qty Ship",
                "unit": "Unit",
                "case_pack": "Case Pack",
                "unit_price": "Unit Price",
                "extension": "Total"
            }
            df = df.rename(columns={k: column_names[k] for k in display_columns})
            
            # Format numeric columns
            for col in ["Qty Ord", "Qty Ship"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            
            for col in ["Unit Price", "Total"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    df[col] = df[col].apply(lambda x: f"${x:.2f}")
            
            # Print the extracted items in a nicely formatted table
            print("\nEXTRACTED LINE ITEMS:")
            print("-"*80)
            print(df.to_string(index=False))
            print("-"*80)
            
            # Calculate and display the total
            try:
                total = sum(float(item['extension']) for item in line_items if item['extension'])
                shipping = 0  # You may extract this separately if available
                
                print(f"\nSUMMARY:")
                print(f"{'Items Total:':<15} ${total:.2f}")
                print(f"{'Shipping:':<15} ${shipping:.2f}")
                print(f"{'Grand Total:':<15} ${total + shipping:.2f}")
                print("-"*80)
            except (ValueError, TypeError):
                print("Could not calculate total - invalid numeric data")
        else:
            print("\nNo line items extracted from invoice")
            return False
        
        # Save to database
        print("\nSaving invoice data to database...")
        success = save_invoice_to_database(invoice_num, customer_id, date, line_items)
        
        if success:
            print(f"Successfully saved invoice #{invoice_num} to database")
        else:
            print(f"Failed to save invoice #{invoice_num} to database")
            
        print("="*80 + "\n")
        return success
        
    except Exception as e:
        print(f"Error processing invoice data: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def query_customer_purchases(customer_id=None, start_date=None, end_date=None):
    """
    Query purchases by customer with optional date range filtering.
    Returns a pandas DataFrame with the results.
    """
    conn = sqlite3.connect('invoice_database.db')
    
    query = """
    SELECT i.invoice_id, i.customer_id, i.invoice_date, i.total_amount,
           COUNT(it.id) AS item_count
    FROM invoices i
    LEFT JOIN invoice_items it ON i.invoice_id = it.invoice_id
    """
    
    conditions = []
    params = []
    
    if customer_id:
        conditions.append("i.customer_id = ?")
        params.append(customer_id)
    
    if start_date:
        conditions.append("i.invoice_date >= ?")
        params.append(start_date)
    
    if end_date:
        conditions.append("i.invoice_date <= ?")
        params.append(end_date)
    
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    query += " GROUP BY i.invoice_id ORDER BY i.invoice_date DESC"
    
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    
    return df

def query_product_sales(item_no=None, description=None, start_date=None, end_date=None):
    """
    Query sales data for specific products.
    You can search by item_no, description (partial match), or date range.
    Returns a pandas DataFrame with the results.
    """
    conn = sqlite3.connect('invoice_database.db')
    
    query = """
    SELECT it.item_no, it.description, it.case_pack, 
           SUM(it.qty_shipped) AS total_quantity,
           AVG(it.unit_price) AS avg_price,
           SUM(it.extension) AS total_sales,
           COUNT(DISTINCT i.invoice_id) AS order_count,
           COUNT(DISTINCT i.customer_id) AS customer_count
    FROM invoice_items it
    JOIN invoices i ON it.invoice_id = i.invoice_id
    """
    
    conditions = []
    params = []
    
    if item_no:
        conditions.append("it.item_no = ?")
        params.append(item_no)
    
    if description:
        conditions.append("it.description LIKE ?")
        params.append(f"%{description}%")
    
    if start_date:
        conditions.append("i.invoice_date >= ?")
        params.append(start_date)
    
    if end_date:
        conditions.append("i.invoice_date <= ?")
        params.append(end_date)
    
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    query += " GROUP BY it.item_no ORDER BY total_sales DESC"
    
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    
    return df

def query_sales_by_day_of_week():
    """
    Analyze sales patterns by day of week.
    Returns a pandas DataFrame with day of week and sales metrics.
    """
    conn = sqlite3.connect('invoice_database.db')
    
    # SQLite has a strftime function to extract day of week (0=Sunday, 6=Saturday)
    query = """
    SELECT 
        CASE 
            WHEN strftime('%w', i.invoice_date) = '0' THEN 'Sunday'
            WHEN strftime('%w', i.invoice_date) = '1' THEN 'Monday'
            WHEN strftime('%w', i.invoice_date) = '2' THEN 'Tuesday'
            WHEN strftime('%w', i.invoice_date) = '3' THEN 'Wednesday'
            WHEN strftime('%w', i.invoice_date) = '4' THEN 'Thursday'
            WHEN strftime('%w', i.invoice_date) = '5' THEN 'Friday'
            WHEN strftime('%w', i.invoice_date) = '6' THEN 'Saturday'
        END AS day_of_week,
        COUNT(DISTINCT i.invoice_id) AS invoice_count,
        SUM(i.total_amount) AS total_sales,
        AVG(i.total_amount) AS avg_invoice_amount,
        COUNT(DISTINCT i.customer_id) AS unique_customers
    FROM invoices i
    GROUP BY day_of_week
    ORDER BY CASE day_of_week
        WHEN 'Sunday' THEN 1
        WHEN 'Monday' THEN 2
        WHEN 'Tuesday' THEN 3
        WHEN 'Wednesday' THEN 4
        WHEN 'Thursday' THEN 5
        WHEN 'Friday' THEN 6
        WHEN 'Saturday' THEN 7
    END
    """
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    return df

def generate_customer_activity_report():
    """
    Generate a report of customer purchasing activity.
    Returns a pandas DataFrame sorted by total purchase amount.
    """
    conn = sqlite3.connect('invoice_database.db')
    
    query = """
    SELECT 
        i.customer_id,
        COUNT(DISTINCT i.invoice_id) AS invoice_count,
        MIN(i.invoice_date) AS first_order_date,
        MAX(i.invoice_date) AS last_order_date,
        SUM(i.total_amount) AS total_spent,
        AVG(i.total_amount) AS avg_order_value,
        COUNT(DISTINCT it.item_no) AS unique_products_ordered
    FROM invoices i
    LEFT JOIN invoice_items it ON i.invoice_id = it.invoice_id
    GROUP BY i.customer_id
    ORDER BY total_spent DESC
    """
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    return df

def test_invoice_extraction(image_path):
    """Test function for extracting data from a single invoice."""
    print(f"Testing invoice extraction on: {image_path}")
    
    # Read the image
    image = cv2.imread(image_path)
    texts = extract_text_and_positions(image)
    
    # Print tabular form
    print_tabular_data(texts)
    
    # Extract metadata
    invoice_num = None
    customer_id = None
    invoice_date = None
    
    for threshold in range(5, 501, 5):
        if not invoice_num:
            invoice_num = find_text_near_keyphrase(texts, "INVOICE NO", "below", threshold)
        if not customer_id:
            customer_id = find_text_near_keyphrase(texts, "ACCOUNT NO", "below", threshold)
        if not invoice_date:
            invoice_date = find_text_near_keyphrase(texts, "INVOICE DATE", "below", threshold)
        
        if invoice_num and customer_id and invoice_date:
            break
    
    print(f"\nExtracted Metadata:")
    print(f"Invoice Number: {invoice_num}")
    print(f"Customer ID: {customer_id}")
    print(f"Invoice Date: {invoice_date}")
    
    # Extract line items
    line_items = extract_invoice_items(texts)
    
    # Display as DataFrame
    if line_items:
        df = pd.DataFrame(line_items)
        print("\nExtracted Line Items:")
        print(df)
        
        # Save to CSV for review
        df.to_csv("test_extraction.csv", index=False)
        print("Line items saved to test_extraction.csv")
        
        # Calculate total
        try:
            total = sum(float(item['extension']) for item in line_items if item['extension'])
            print(f"\nTotal Invoice Amount: ${total:.2f}")
        except:
            print("Could not calculate total")
    else:
        print("No line items extracted")
    
    # If you have valid metadata, try saving to database
    if invoice_num and customer_id and invoice_date and line_items:
        save_invoice_to_database(invoice_num, customer_id, invoice_date, line_items)

# Add this at the end of your file, replacing or enhancing your existing main block

if __name__ == "__main__":
    print("\n" + "="*80)
    print("INVOICE PROCESSING AND DATA EXTRACTION SYSTEM")
    print("="*80)
    
    # Validate paths
    if not validate_paths():
        print("Exiting due to invalid paths.")
        exit(1)
    
    # Display menu
    print("\nSelect an option:")
    print("1. Process all invoices in folder")
    print("2. Test OCR on a single invoice")
    print("3. Generate reports from database")
    print("4. Exit")
    
    choice = input("\nEnter your choice (1-4): ")
    
    if choice == "1":
        # Process all invoices in the folder
        print("\nProcessing all invoices in folder:", invoice_folder)
        process_invoices(invoice_folder)
        check_inactive_customers()
        
    elif choice == "2":
        # Test OCR on a single invoice
        test_file = input("\nEnter path to invoice image file: ")
        if os.path.exists(test_file):
            test_ocr_extraction(test_file)
        else:
            print(f"ERROR: File not found: {test_file}")
            
    elif choice == "3":
        # Generate reports from database
        print("\n" + "-"*80)
        print("DATABASE REPORTS")
        print("-"*80)
        print("Select a report type:")
        print("1. Customer Purchase History")
        print("2. Product Sales Analysis")
        print("3. Sales by Day of Week")
        print("4. Customer Activity Report")
        print("5. Return to main menu")
        
        report_choice = input("\nEnter report choice (1-5): ")
        
        if report_choice == "1":
            customer_id = input("Enter customer ID (leave blank for all): ")
            if customer_id:
                df = query_customer_purchases(customer_id)
            else:
                df = query_customer_purchases()
                
            if not df.empty:
                print("\nCustomer Purchase History:")
                print(df)
                csv_path = f"customer_purchases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                df.to_csv(csv_path, index=False)
                print(f"\nReport saved to {csv_path}")
            else:
                print("No purchase history found")
                
        elif report_choice == "2":
            product = input("Enter product code or description (leave blank for all): ")
            if len(product) == 6 and product.isalnum():
                df = query_product_sales(item_no=product)
            elif product:
                df = query_product_sales(description=product)
            else:
                df = query_product_sales()
                
            if not df.empty:
                print("\nProduct Sales Report:")
                print(df)
                csv_path = f"product_sales_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                df.to_csv(csv_path, index=False)
                print(f"\nReport saved to {csv_path}")
            else:
                print("No product sales found")
                
        elif report_choice == "3":
            df = query_sales_by_day_of_week()
            if not df.empty:
                print("\nSales by Day of Week:")
                print(df)
                csv_path = f"sales_by_day_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                df.to_csv(csv_path, index=False)
                print(f"\nReport saved to {csv_path}")
            else:
                print("No sales data found")
                
        elif report_choice == "4":
            df = generate_customer_activity_report()
            if not df.empty:
                print("\nCustomer Activity Report:")
                print(df)
                csv_path = f"customer_activity_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                df.to_csv(csv_path, index=False)
                print(f"\nReport saved to {csv_path}")
            else:
                print("No customer activity data found")
                
    print("\nProgram execution completed.")
    input("Press Enter to exit...")