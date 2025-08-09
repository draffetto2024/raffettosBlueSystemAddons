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
import sqlite3
import logging
import tempfile
import time
import random
import string

# Simplified path definitions using relative paths
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
sender_email = "derekraffetto@gmail.com"
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
    """
    Extract text and positions from an image using Google Vision API.
    
    Args:
        image: Either a file path (string) or an image array (numpy.ndarray)
    
    Returns:
        Text annotations from Google Vision API
    """
    # Check if the input is a path or an image array
    if isinstance(image, str):
        # It's a file path
        image_data = cv2.imread(image)
        if image_data is None:
            raise Exception(f"Failed to read image from path: {image}")
    else:
        # It's already an image array
        image_data = image
        
    _, encoded_image = cv2.imencode('.jpg', image_data)
    content = encoded_image.tobytes()
    image_obj = types.Image(content=content)
    response = client.document_text_detection(image=image_obj)
    if response.error.message:
        raise Exception(f'{response.error.message}')
    
    return response.text_annotations

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

def align_invoice_image(image, texts):
    """
    Align (rotate) the invoice image to make text perfectly level.
    Uses 'INVOICE DATE' to detect the rotation angle needed.
    Returns the rotated image with level text.
    """
    print("DEBUG: Starting invoice alignment (rotation) using 'INVOICE DATE'...")
    
    # Find "INVOICE DATE" text - be more flexible in search
    invoice_date_text = None
    invoice_date_box = None
    
    # Search for INVOICE DATE in different ways
    for text in texts[1:]:  # Skip first element which is full text
        content = text.description.upper().strip()
        print(f"DEBUG: Checking text: '{content}'")
        
        # Look for exact match first
        if "INVOICE DATE" in content:
            invoice_date_text = text
            break
        # Look for "INVOICE" and "DATE" separately but nearby
        elif "INVOICE" in content:
            invoice_box = text.bounding_poly.vertices
            # Look for "DATE" nearby
            for other_text in texts[1:]:
                other_content = other_text.description.upper().strip()
                if "DATE" in other_content:
                    other_box = other_text.bounding_poly.vertices
                    # Check if they're on the same line (within 20 pixels vertically)
                    if abs(invoice_box[0].y - other_box[0].y) < 20:
                        # Use the combined line for angle calculation
                        invoice_date_text = text  # We'll use the INVOICE text for angle
                        invoice_date_box = other_box  # And DATE box for reference
                        print(f"DEBUG: Found 'INVOICE' and 'DATE' separately")
                        break
            if invoice_date_text:
                break
    
    if not invoice_date_text:
        print("WARNING: Could not find 'INVOICE DATE' for alignment, using image as-is")
        return image, 0.0
    
    # Calculate the angle of the text line
    box = invoice_date_text.bounding_poly.vertices
    
    # Get the angle of the text line using the bounding box
    # Use the top edge of the bounding box (from top-left to top-right)
    x1, y1 = box[0].x, box[0].y  # Top-left
    x2, y2 = box[1].x, box[1].y  # Top-right
    
    # Calculate angle in degrees
    # atan2 gives us the angle of the line from horizontal
    angle_radians = np.arctan2(y2 - y1, x2 - x1)
    angle_degrees = np.degrees(angle_radians)
    
    print(f"DEBUG: Text line from ({x1}, {y1}) to ({x2}, {y2})")
    print(f"DEBUG: Delta X: {x2-x1}, Delta Y: {y2-y1}")
    print(f"DEBUG: Calculated text angle: {angle_degrees:.2f} degrees")
    
    if y2 > y1:
        print(f"DEBUG: Text slopes DOWNWARD from left to right (needs counter-clockwise correction)")
    elif y2 < y1:
        print(f"DEBUG: Text slopes UPWARD from left to right (needs clockwise correction)")
    else:
        print(f"DEBUG: Text is already horizontal")
    
    # To make the text horizontal, we need to rotate by the negative of the detected angle
    rotation_angle = -angle_degrees
    
    # HACK: For some reason the rotation is still going the wrong way, so flip it
    rotation_angle = rotation_angle * -1  # This makes it go the right direction
    
    print(f"DEBUG: Applying rotation of {rotation_angle:.2f} degrees to level the text")
    print(f"DEBUG: (Negative angle = counter-clockwise, Positive angle = clockwise)")
    
    # Apply rotation around the center of the image
    height, width = image.shape[:2]
    center = (width // 2, height // 2)
    
    # Create rotation matrix - THIS IS WHERE THE ACTUAL ROTATION HAPPENS
    rotation_matrix = cv2.getRotationMatrix2D(center, rotation_angle, 1.0)
    
    # Apply rotation - THIS IS THE LINE THAT DOES THE ROTATION
    rotated_image = cv2.warpAffine(image, rotation_matrix, (width, height), 
                                   borderMode=cv2.BORDER_CONSTANT, borderValue=(255, 255, 255))
    
    # Save debug images showing before and after rotation
    debug_dir = os.getcwd()
    
    # Save original image with angle visualization
    angle_debug = image.copy()
    # Draw the detected text line
    cv2.line(angle_debug, (x1, y1), (x2, y2), (0, 0, 255), 3)  # Red line showing detected angle
    # Draw a horizontal reference line
    cv2.line(angle_debug, (x1, y1), (x1 + (x2-x1), y1), (0, 255, 0), 2)  # Green line showing horizontal
    cv2.putText(angle_debug, f"Detected: {angle_degrees:.1f}°", (x1, y1-10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
    cv2.putText(angle_debug, f"Correction: {rotation_angle:.1f}°", (x1, y1-35), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 0, 0), 2)
    
    angle_debug_path = os.path.join(debug_dir, "rotation_debug_angle_detection.jpg")
    cv2.imwrite(angle_debug_path, angle_debug)
    print(f"DEBUG: Saved angle detection debug to {angle_debug_path}")
    
    # Save original image
    orig_debug_path = os.path.join(debug_dir, "rotation_debug_original.jpg")
    cv2.imwrite(orig_debug_path, image)
    print(f"DEBUG: Saved original image to {orig_debug_path}")
    
    # Save rotated image
    rotated_debug_path = os.path.join(debug_dir, "rotation_debug_rotated.jpg")
    cv2.imwrite(rotated_debug_path, rotated_image)
    print(f"DEBUG: Saved rotated image to {rotated_debug_path}")
    
    return rotated_image, rotation_angle

def extract_line_items_precise(image, texts, debug=True):
    """
    Extract line items using precise pixel coordinates after alignment.
    Uses the pre-existing OCR results instead of making new API calls.
    """
    print("="*60)
    print("DEBUG: STARTING PRECISE LINE ITEM EXTRACTION")
    print("="*60)
    
    # Check if image is valid
    if image is None:
        print("ERROR: Image is None!")
        return []
    
    if not hasattr(image, 'shape'):
        print("ERROR: Image doesn't have shape attribute!")
        return []
    
    print(f"DEBUG: Image is valid, dimensions: {image.shape[1]}x{image.shape[0]}")
    print(f"DEBUG: Using existing OCR results with {len(texts)} text elements")

    img_height, img_width = image.shape[:2]
    
    # Define precise pixel boundaries for each column (after alignment)
    ITEM_CODE_X_START = 0
    ITEM_CODE_X_END = 140 
    
    QTY_ORDERED_X_START = 145
    QTY_ORDERED_X_END = 200
    
    QTY_SHIPPED_X_START = 200
    QTY_SHIPPED_X_END = 280
    
    EXTENSION_X_START = 1090
    EXTENSION_X_END = img_width
    
    FIRST_ROW_Y = 445
    ROW_HEIGHT = 45 
    MAX_ROWS = 20    
    
    print(f"\nDEBUG: COORDINATES BEING USED:")
    print(f"  ITEM_CODE:    {ITEM_CODE_X_START} to {ITEM_CODE_X_END} (GREEN)")
    print(f"  QTY_ORDERED:  {QTY_ORDERED_X_START} to {QTY_ORDERED_X_END} (RED)")
    print(f"  QTY_SHIPPED:  {QTY_SHIPPED_X_START} to {QTY_SHIPPED_X_END} (BLUE)")
    print(f"  EXTENSION:    {EXTENSION_X_START} to {EXTENSION_X_END} (YELLOW)")
    print(f"  FIRST_ROW_Y:  {FIRST_ROW_Y} (MAGENTA)")
    print(f"  ROW_HEIGHT:   {ROW_HEIGHT}")
    
    line_items = []
    
    if debug:
        print(f"\nDEBUG: CREATING DEBUG VISUALIZATION...")
        try:
            debug_img = image.copy()
            print(f"DEBUG: Successfully copied image for debug visualization")
        except Exception as e:
            print(f"ERROR: Failed to copy image: {e}")
            return []
        
        # Check if coordinates are within image bounds
        print(f"DEBUG: Image bounds check - Width: {img_width}, Height: {img_height}")
        
        if EXTENSION_X_END > img_width:
            print(f"WARNING: EXTENSION column ({EXTENSION_X_START}-{EXTENSION_X_END}) extends beyond image width ({img_width})")
        
        print(f"DEBUG: Drawing column boundaries...")
        
        try:
            # Draw ITEM_CODE column boundaries (GREEN)
            cv2.line(debug_img, (ITEM_CODE_X_START, 0), (ITEM_CODE_X_START, img_height), (0, 255, 0), 5)
            cv2.line(debug_img, (ITEM_CODE_X_END, 0), (ITEM_CODE_X_END, img_height), (0, 255, 0), 5)
            cv2.putText(debug_img, "ITEM_CODE", (ITEM_CODE_X_START + 5, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 3)
            print(f"DEBUG: ✓ Drew ITEM_CODE boundaries (GREEN)")
            
            # Draw QTY_ORDERED column boundaries (RED)
            cv2.line(debug_img, (QTY_ORDERED_X_START, 0), (QTY_ORDERED_X_START, img_height), (0, 0, 255), 5)
            cv2.line(debug_img, (QTY_ORDERED_X_END, 0), (QTY_ORDERED_X_END, img_height), (0, 0, 255), 5)
            cv2.putText(debug_img, "QTY_ORD", (QTY_ORDERED_X_START + 5, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 3)
            print(f"DEBUG: ✓ Drew QTY_ORDERED boundaries (RED)")
            
            # Draw QTY_SHIPPED column boundaries (BLUE)
            cv2.line(debug_img, (QTY_SHIPPED_X_START, 0), (QTY_SHIPPED_X_START, img_height), (255, 0, 0), 5)
            cv2.line(debug_img, (QTY_SHIPPED_X_END, 0), (QTY_SHIPPED_X_END, img_height), (255, 0, 0), 5)
            cv2.putText(debug_img, "QTY_SHIP", (QTY_SHIPPED_X_START + 5, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 3)
            print(f"DEBUG: ✓ Drew QTY_SHIPPED boundaries (BLUE)")
            
            # Draw EXTENSION column boundaries (YELLOW) - only if within bounds
            if EXTENSION_X_END <= img_width:
                cv2.line(debug_img, (EXTENSION_X_START, 0), (EXTENSION_X_START, img_height), (0, 255, 255), 5)
                cv2.line(debug_img, (EXTENSION_X_END, 0), (EXTENSION_X_END, img_height), (0, 255, 255), 5)
                cv2.putText(debug_img, "EXTENSION", (EXTENSION_X_START + 5, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 255), 3)
                print(f"DEBUG: ✓ Drew EXTENSION boundaries (YELLOW)")
            else:
                print(f"DEBUG: ⚠️ Skipped EXTENSION boundaries (outside image bounds)")
            
            # Draw starting row line (MAGENTA)
            cv2.line(debug_img, (0, FIRST_ROW_Y), (img_width, FIRST_ROW_Y), (255, 0, 255), 5)
            cv2.putText(debug_img, f"FIRST_ROW_Y={FIRST_ROW_Y}", (10, FIRST_ROW_Y - 10), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 255), 3)
            print(f"DEBUG: ✓ Drew FIRST_ROW line (MAGENTA)")
            
        except Exception as e:
            print(f"ERROR: Failed to draw lines: {e}")
            import traceback
            traceback.print_exc()
        
        print(f"DEBUG: Drawing row boundaries...")
    
    # Extract each row using existing OCR results
    for row_index in range(min(MAX_ROWS, 15)):  # Increased limit for debugging
        row_y_start = FIRST_ROW_Y + (row_index * ROW_HEIGHT)
        row_y_end = row_y_start + ROW_HEIGHT
        
        if row_y_end >= image.shape[0]:
            print(f"DEBUG: Row {row_index + 1} extends beyond image bounds, stopping")
            break
        
        print(f"\nDEBUG: Processing Row {row_index + 1} (Y: {row_y_start} to {row_y_end})")
        
        if debug:
            try:
                # Draw row boundaries (GRAY)
                cv2.line(debug_img, (0, row_y_start), (img_width, row_y_start), (128, 128, 128), 3)
                cv2.line(debug_img, (0, row_y_end), (img_width, row_y_end), (128, 128, 128), 2)
                cv2.putText(debug_img, f"R{row_index + 1}", (5, row_y_start + 25), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 0, 255), 2)
                print(f"DEBUG: ✓ Drew row {row_index + 1} boundaries")
            except Exception as e:
                print(f"ERROR: Failed to draw row boundaries: {e}")
        
        # Find text within each column for this row using existing OCR results
        item_code = find_text_in_region(texts, ITEM_CODE_X_START, ITEM_CODE_X_END, row_y_start, row_y_end)
        qty_ordered = find_text_in_region(texts, QTY_ORDERED_X_START, QTY_ORDERED_X_END, row_y_start, row_y_end)
        qty_shipped = find_text_in_region(texts, QTY_SHIPPED_X_START, QTY_SHIPPED_X_END, row_y_start, row_y_end)
        extension = find_text_in_region(texts, EXTENSION_X_START, EXTENSION_X_END, row_y_start, row_y_end)
        
        print(f"  OCR found: Item='{item_code}', QtyOrd='{qty_ordered}', QtyShip='{qty_shipped}', Ext='{extension}'")
        
        # Only process if we found a valid 6-character item code
        if not item_code or not re.match(r'^[A-Za-z0-9]{6}$', item_code):
            print(f"  -> Skipping row {row_index + 1}: Invalid item code")
            continue
        
        # Create line item
        line_item = {
            "item_no": item_code,
            "qty_ordered": clean_quantity(qty_ordered),
            "qty_shipped": clean_quantity(qty_shipped),  # Now using actual qty_shipped!
            "extension": clean_price(extension)
        }
        
        line_items.append(line_item)
        print(f"  -> ✓ Added line item: {line_item}")
    
    if debug:
        print(f"\nDEBUG: ATTEMPTING TO SAVE DEBUG IMAGE...")
        
        # Try multiple save locations
        save_locations = [
            os.path.join(os.getcwd(), "debug_precise_extraction.jpg"),
            "./debug_precise_extraction.jpg",
            "debug_precise_extraction.jpg",
            os.path.expanduser("~/debug_precise_extraction.jpg")
        ]
        
        saved_successfully = False
        for i, debug_path in enumerate(save_locations):
            try:
                print(f"DEBUG: Attempt {i+1}: Trying to save to {debug_path}")
                success = cv2.imwrite(debug_path, debug_img)
                print(f"DEBUG: cv2.imwrite returned: {success}")
                
                if success and os.path.exists(debug_path):
                    file_size = os.path.getsize(debug_path)
                    print(f"DEBUG: ✓ SUCCESS! Saved debug image to {debug_path}")
                    print(f"DEBUG: File size: {file_size} bytes")
                    saved_successfully = True
                    break
                else:
                    print(f"DEBUG: ✗ Failed to save to {debug_path}")
                    
            except Exception as e:
                print(f"DEBUG: ✗ Exception saving to {debug_path}: {e}")
        
        if not saved_successfully:
            print(f"DEBUG: ✗ FAILED TO SAVE DEBUG IMAGE TO ANY LOCATION!")
    
    print(f"\nDEBUG: EXTRACTION COMPLETE")
    print(f"DEBUG: Extracted {len(line_items)} line items")
    print("="*60)
    
    return line_items

def find_text_in_region(texts, x_start, x_end, y_start, y_end):
    """
    Find text within a specific rectangular region using existing OCR results.
    Now combines adjacent text elements to handle split product codes.
    """
    found_texts = []
    
    # Skip the first element which is the full document text
    for text in texts[1:]:
        if not hasattr(text, 'bounding_poly') or not text.bounding_poly.vertices:
            continue
            
        box = text.bounding_poly.vertices
        
        # Get text center point
        text_x = (box[0].x + box[2].x) / 2
        text_y = (box[0].y + box[2].y) / 2
        
        # Check if text center is within our region
        if (x_start <= text_x <= x_end and y_start <= text_y <= y_end):
            found_texts.append({
                'text': text.description.strip(),
                'x': text_x,
                'y': text_y,
                'box': box
            })
    
    if not found_texts:
        return ""
    
    # Sort by x position (left to right)
    found_texts.sort(key=lambda x: x['x'])
    
    # Try different combinations
    candidates = []
    
    # Single longest text
    if found_texts:
        longest_single = max(found_texts, key=lambda x: len(x['text']))
        candidates.append(longest_single['text'])
    
    # Try combining adjacent texts (up to 3 elements)
    for i in range(len(found_texts)):
        # Combine 2 adjacent
        if i < len(found_texts) - 1:
            combined = found_texts[i]['text'] + found_texts[i+1]['text']
            candidates.append(combined)
        
        # Combine 3 adjacent
        if i < len(found_texts) - 2:
            combined = found_texts[i]['text'] + found_texts[i+1]['text'] + found_texts[i+2]['text']
            candidates.append(combined)
    
    # Try all texts concatenated
    if len(found_texts) > 1:
        all_combined = ''.join([t['text'] for t in found_texts])
        candidates.append(all_combined)
    
    # Check which candidate looks most like a product code (6 alphanumeric)
    for candidate in candidates:
        if re.match(r'^[A-Za-z0-9]{6}$', candidate):
            return candidate
    
    # If no perfect 6-char match, return the longest candidate
    if candidates:
        return max(candidates, key=len)
    
    return ""

def extract_text_from_roi(roi):
    """
    Extract text from a region of interest using Google Vision API.
    This is used for the precise pixel-based extraction.
    """
    if roi.size == 0:
        print(f"  DEBUG: ROI is empty (size=0)")
        return ""
    
    try:
        print(f"  DEBUG: Processing ROI of size {roi.shape[1]}x{roi.shape[0]}")
        
        # Apply preprocessing to improve OCR accuracy
        _, binary = cv2.threshold(roi, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # Resize for better OCR (scale up small regions)
        height, width = binary.shape
        if height < 20 or width < 20:
            scale_factor = max(2, 40 // min(height, width))
            binary = cv2.resize(binary, None, fx=scale_factor, fy=scale_factor, interpolation=cv2.INTER_CUBIC)
            print(f"  DEBUG: Scaled ROI by {scale_factor}x to {binary.shape[1]}x{binary.shape[0]}")
        
        # Check if ROI has any content
        content_pixels = np.sum(binary)
        print(f"  DEBUG: ROI has {content_pixels} content pixels")
        
        if content_pixels < 50:  # Very little content
            print(f"  DEBUG: Too few content pixels, returning empty")
            return ""
        
        # Use Google Vision API for this ROI
        _, encoded_image = cv2.imencode('.jpg', binary)
        content = encoded_image.tobytes()
        image_obj = types.Image(content=content)
        response = client.document_text_detection(image=image_obj)
        
        if response.text_annotations:
            text = response.text_annotations[0].description.strip()
            print(f"  DEBUG: OCR returned: '{text}'")
            return text
        else:
            print(f"  DEBUG: OCR returned no text")
            return ""
            
    except Exception as e:
        print(f"  DEBUG: Error extracting text from ROI: {e}")
        return ""

def clean_quantity(text):
    """Clean and validate quantity text."""
    if not text:
        return ""
    
    # Remove non-numeric characters except decimal point
    cleaned = re.sub(r'[^\d.]', '', text)
    
    if not cleaned or cleaned == '.':
        return ""
    
    try:
        value = float(cleaned)
        # Return as integer if it's a whole number
        return str(int(value)) if value.is_integer() else str(value)
    except ValueError:
        return ""

def clean_price(text):
    """Clean and validate price text."""
    if not text:
        return ""
    
    # Remove currency symbols and non-numeric characters except decimal point
    cleaned = re.sub(r'[^\d.]', '', text)
    
    if not cleaned or cleaned == '.':
        return ""
    
    try:
        value = float(cleaned)
        return str(value)
    except ValueError:
        return ""

def copy_image_to_sorted_folder(image_path, six_digit_number, customer_id, date):
    year, month, day = extract_year(date), extract_month(date), extract_day(date)
    destination_folder = os.path.join(destination_base_folder, year, month, day)
    os.makedirs(destination_folder, exist_ok=True)
    
    base_name = f"{customer_id}_{six_digit_number}"
    destination_path = os.path.join(destination_folder, f"{base_name}.pdf")
    counter = 1
    
    while os.path.exists(destination_path):
        destination_path = os.path.join(destination_folder, f"{base_name}_{counter}.pdf")
        counter += 1
    
    try:
        if image_path.lower().endswith('.pdf'):
            shutil.copy2(image_path, destination_path)
        else:
            # Convert to optimized PDF
            img = cv2.imread(image_path)
            if img is None:
                print(f"ERROR: Cannot read image: {image_path}")
                return
                
            if len(img.shape) > 2 and img.shape[2] > 1:
                img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            else:
                img_gray = img
                
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_file:
                temp_image_path = temp_file.name
                cv2.imwrite(temp_image_path, img_gray, [cv2.IMWRITE_JPEG_QUALITY, 40])
            
            doc = fitz.open()
            page = doc.new_page()
            page.insert_image(fitz.Rect(0, 0, page.rect.width, page.rect.height), filename=temp_image_path)
            doc.save(destination_path, deflate=True, garbage=4, clean=True)
            doc.close()
            os.remove(temp_image_path)
        
        print(f"Copied and saved as optimized PDF: {destination_path}")
        customer_tracker.update_customer_activity(customer_id, date)
    except Exception as e:
        print(f"DEBUG: Error copying/converting file: {e}")
        
    send_email_with_attachment(destination_path, customer_id, six_digit_number, date)

def copy_image_to_unsorted_folder(image_path):
    today = datetime.now().strftime("%Y-%m-%d")
    destination_folder = os.path.join(unsorted_base_folder, today)
    os.makedirs(destination_folder, exist_ok=True)
    
    base_name = os.path.splitext(os.path.basename(image_path))[0]
    destination_path = os.path.join(destination_folder, f"{base_name}.pdf")
    
    print(f"DEBUG: Attempting to copy unsorted invoice from {image_path} to {destination_path}")
    
    try:
        if image_path.lower().endswith('.pdf'):
            shutil.copy2(image_path, destination_path)
        else:
            img = cv2.imread(image_path)
            if img is None:
                print(f"ERROR: Cannot read image: {image_path}")
                return
                
            if len(img.shape) > 2 and img.shape[2] > 1:
                img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            else:
                img_gray = img
                
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_file:
                temp_image_path = temp_file.name
                cv2.imwrite(temp_image_path, img_gray, [cv2.IMWRITE_JPEG_QUALITY, 40])
            
            doc = fitz.open()
            page = doc.new_page()
            page.insert_image(fitz.Rect(0, 0, page.rect.width, page.rect.height), filename=temp_image_path)
            doc.save(destination_path, deflate=True, garbage=4, clean=True)
            doc.close()
            os.remove(temp_image_path)
        
        print(f"Copied unsorted invoice to optimized PDF: {destination_path}")
    except Exception as e:
        print(f"DEBUG: Error copying/converting unsorted file: {e}")

def send_email_with_attachment(file_path, customer_id, invoice_number, date):
    receiver_email = customer_emails.get(customer_id, None)
    if receiver_email is None:
        print(f"No email found for customer {customer_id}. Skipping email.")
        return

    subject = f"RAFFETTO'S FRESH PASTA INVOICE Invoice {invoice_number}"
    body = f"""Please find the attached invoice, {invoice_number} shipped, {date}. Thank you.
    
    Derek Raffetto
    Raffetto's Fresh Pasta"""

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

def convert_pdf_to_images(pdf_path):
    doc = fitz.open(pdf_path)
    image_paths = []
    try:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                pix.save(tmp.name)
                image_paths.append(tmp.name)
            print(f"Converted page {i+1} to image: {tmp.name}")
    finally:
        doc.close()
    return image_paths

def process_invoice_image(image_path, manual_mode=False):
    """Enhanced process_invoice_image function with precise pixel-based extraction."""
    log_folder = "./Logs"
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, f"ocr_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    logging.basicConfig(filename=log_file, level=logging.DEBUG, format='%(asctime)s - %(message)s')
    
    logging.info(f"DEBUG: Processing image: {image_path}")
    print(f"DEBUG: Starting to process image: {image_path}")
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
        print(f"DEBUG: Loading image for processing...")
        # Load and preprocess the image
        if image_path.lower().endswith('.pdf'):
            print(f"DEBUG: Processing PDF file...")
            # For PDF files, extract the first page as an image
            doc = fitz.open(image_path)
            if doc.page_count > 0:
                print(f"DEBUG: Extracting page 1 from PDF...")
                pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_file:
                    temp_image_path = temp_file.name
                pix.save(temp_image_path)
                image = cv2.imread(temp_image_path)
                doc.close()
                print(f"DEBUG: PDF page extracted to temporary image: {temp_image_path}")
                cleanup_temp = lambda: os.path.exists(temp_image_path) and os.remove(temp_image_path)
            else:
                print(f"Error: PDF has no pages: {image_path}")
                return False
        else:
            print(f"DEBUG: Processing image file directly...")
            image = cv2.imread(image_path)
            cleanup_temp = lambda: None
            
        if image is None:
            print(f"Error reading image: {image_path}")
            return False
            
        print(f"DEBUG: Image loaded successfully, dimensions: {image.shape[1]}x{image.shape[0]}")
        
        # Extract text using Google Vision API
        print(f"DEBUG: Extracting text using Google Vision API...")
        texts = extract_text_and_positions(image)
        print(f"DEBUG: Found {len(texts)} text elements")

        # LOG ALL OCR RESULTS FOR DEBUGGING
        logging.info("="*80)
        logging.info("COMPLETE OCR RESULTS FROM GOOGLE VISION API")
        logging.info("="*80)
        logging.info(f"Total text elements found: {len(texts)}")
        logging.info("-"*80)

        for i, text in enumerate(texts):
            if hasattr(text, 'bounding_poly') and text.bounding_poly.vertices:
                box = text.bounding_poly.vertices
                content = text.description.strip()
                logging.info(f"Element {i:2d}: '{content}'")
                logging.info(f"   Position: ({box[0].x:3d}, {box[0].y:3d}) to ({box[2].x:3d}, {box[2].y:3d})")
                logging.info(f"   Size: {box[2].x - box[0].x}x{box[2].y - box[0].y} pixels")
                logging.info("-"*40)
            else:
                logging.info(f"Element {i:2d}: '{text.description.strip() if hasattr(text, 'description') else 'NO DESCRIPTION'}'")
                logging.info("   Position: NO BOUNDING BOX")
                logging.info("-"*40)

        logging.info("="*80)
        
        # Align the image using INVOICE DATE as reference
        print(f"DEBUG: Aligning (rotating) image using INVOICE DATE...")
        aligned_image, rotation_angle = align_invoice_image(image, texts)
        print(f"DEBUG: Image rotated by {rotation_angle:.2f} degrees")
        
        # Extract metadata using the existing method (works well)
        print(f"DEBUG: Extracting invoice metadata...")
        invoice_num = None
        customer_num = None
        date_num = None
        
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
                print(f"DEBUG: Found all metadata - Invoice: {invoice_num}, Customer: {customer_num}, Date: {date_num}")
                break
                
        if (invoice_num and is_six_alphanumeric(invoice_num) and
            customer_num and is_six_alphanumeric(customer_num) and
            date_num and is_date_format(date_num)):
            logging.info(f"DEBUG: Automatic detection successful")
            
            # Use precise pixel-based extraction for line items
            print(f"DEBUG: Starting line item extraction...")
            line_items = extract_line_items_precise(aligned_image, texts, debug=True)
            print(f"DEBUG: Extracted {len(line_items)} line items")
            
            # Process the extracted data
            if line_items:
                print(f"DEBUG: Processing extracted line items...")
                process_invoice_data_extraction_precise(image_path, invoice_num, customer_num, date_num, line_items)
            
            print(f"DEBUG: Copying invoice to sorted folder...")
            copy_image_to_sorted_folder(image_path, invoice_num, customer_num, date_num)
            processed_successfully = True
        else:
            print(f"DEBUG: Automatic detection failed - Invoice: {invoice_num}, Customer: {customer_num}, Date: {date_num}")
            logging.info(f"DEBUG: Automatic detection failed")
        
        cleanup_temp()
    
    if processed_successfully:
        try:
            os.remove(image_path)
            print(f"Deleted original file: {image_path}")
        except Exception as e:
            print(f"Error deleting original file {image_path}: {e}")
    
    return processed_successfully

def process_invoice_data_extraction_precise(image_path, invoice_num, customer_id, date, line_items):
    """
    Process extracted line items and save to database.
    Simplified version that focuses on the three key data points.
    """
    try:
        print("\n" + "="*80)
        print(f"PROCESSING INVOICE DATA: #{invoice_num} | Customer: {customer_id} | Date: {date}")
        print("="*80)
        
        if line_items:
            print("\n" + "-"*80)
            print("INVOICE METADATA:")
            print(f"Invoice Number: {invoice_num}")
            print(f"Customer ID:    {customer_id}")
            print(f"Invoice Date:   {date}")
            print("-"*80)
            
            # Create a simplified DataFrame focusing on key data
            df_data = []
            for item in line_items:
                df_data.append({
                    "Item Code": item["item_no"],
                    "Qty Ordered": item["qty_ordered"],
                    "Qty Shipped": item["qty_shipped"],
                    "Total": f"${float(item['extension']):.2f}" if item["extension"] else ""
                })
            
            df = pd.DataFrame(df_data)
            
            print("\nEXTRACTED LINE ITEMS:")
            print("-"*80)
            print(df.to_string(index=False))
            print("-"*80)
            
            # Calculate totals
            try:
                subtotal = sum(float(item['extension']) for item in line_items if item['extension'])
                print(f"\nSUMMARY:")
                print(f"{'Items Total:':<15} ${subtotal:.2f}")
                print("-"*80)
            except (ValueError, TypeError) as e:
                print(f"Could not calculate total: {e}")
        else:
            print("\nNo line items extracted from invoice")
            return False
        
        # Save to database
        print("\nSaving invoice data to database...")
        success = save_invoice_to_database_simplified(invoice_num, customer_id, date, line_items)
        
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

def create_simplified_invoice_database():
    """Create simplified SQLite database focusing on key data."""
    print("DEBUG: Creating/checking simplified invoice database...")
    
    try:
        conn = sqlite3.connect('invoice_database.db')
        cursor = conn.cursor()
        
        # Check if tables exist and recreate with simplified schema
        cursor.execute("DROP TABLE IF EXISTS invoice_items")
        cursor.execute("DROP TABLE IF EXISTS invoices")
        
        # Create invoices table
        cursor.execute('''
        CREATE TABLE invoices (
            invoice_id TEXT PRIMARY KEY,
            customer_id TEXT,
            invoice_date TEXT,
            total_amount REAL,
            processed_date TEXT
        )
        ''')
        
        # Create simplified invoice_items table
        cursor.execute('''
        CREATE TABLE invoice_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id TEXT,
            item_no TEXT,
            qty_ordered INTEGER,
            qty_shipped INTEGER,
            extension REAL,
            FOREIGN KEY (invoice_id) REFERENCES invoices (invoice_id)
        )
        ''')
        
        conn.commit()
        print("DEBUG: Simplified database setup complete")
        return True
        
    except sqlite3.Error as e:
        print(f"DEBUG: SQLite error during database creation: {e}")
        return False
        
    finally:
        if 'conn' in locals():
            conn.close()

def save_invoice_to_database_simplified(invoice_num, customer_id, date, line_items):
    """Save invoice and line items to simplified database."""
    print("\nDEBUG: Starting simplified database save operation...")
    
    if not invoice_num or not customer_id or not date or not line_items:
        print("ERROR: Missing required invoice data, not saving to database.")
        return False
    
    if not create_simplified_invoice_database():
        print("ERROR: Failed to create or verify database, cannot save data.")
        return False
    
    try:
        conn = sqlite3.connect('invoice_database.db')
        cursor = conn.cursor()
        
        # Calculate total amount
        total_amount = sum(float(item["extension"]) for item in line_items if item["extension"])
        print(f"DEBUG: Calculated total amount: ${total_amount:.2f}")
        
        # Check if invoice already exists
        cursor.execute("SELECT invoice_id FROM invoices WHERE invoice_id = ?", (invoice_num,))
        if cursor.fetchone():
            print(f"INFO: Invoice {invoice_num} already exists in database, skipping.")
            conn.close()
            return False
        
        # Insert invoice record
        cursor.execute(
            "INSERT INTO invoices (invoice_id, customer_id, invoice_date, total_amount, processed_date) VALUES (?, ?, ?, ?, ?)",
            (invoice_num, customer_id, date, total_amount, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        )
        
        # Insert line items
        for item in line_items:
            try:
                qty_ordered = int(float(item["qty_ordered"])) if item["qty_ordered"] else None
                qty_shipped = int(float(item["qty_shipped"])) if item["qty_shipped"] else None
                extension = float(item["extension"]) if item["extension"] else None
                
                cursor.execute(
                    "INSERT INTO invoice_items (invoice_id, item_no, qty_ordered, qty_shipped, extension) VALUES (?, ?, ?, ?, ?)",
                    (invoice_num, item["item_no"], qty_ordered, qty_shipped, extension)
                )
                
            except (ValueError, TypeError) as e:
                print(f"WARNING: Error converting values for line item: {e}")
                continue
        
        conn.commit()
        print(f"SUCCESS: Invoice {invoice_num} with {len(line_items)} line items saved to database.")
        return True
        
    except sqlite3.Error as e:
        if 'conn' in locals():
            conn.rollback()
        print(f"ERROR: Database error during save operation: {e}")
        return False
    except Exception as e:
        if 'conn' in locals():
            conn.rollback()
        print(f"ERROR: Unexpected error during save operation: {e}")
        return False
    finally:
        if 'conn' in locals():
            conn.close()

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
            time.sleep(0.5)

def preprocess_all_images_in_folder(folder_path):
    """Batch preprocess all image files in a folder."""
    if not os.path.exists(folder_path):
        print(f"ERROR: Folder does not exist: {folder_path}")
        return (0, 0)
    
    success_count = 0
    failure_count = 0
    
    print(f"\nBatch preprocessing all images in: {folder_path}")
    
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            file_path = os.path.join(folder_path, filename)
            
            print(f"Processing: {filename}")
            
            pdf_path = permanently_preprocess_invoice_image(file_path)
            
            if pdf_path:
                success_count += 1
                print(f"Successfully preprocessed and converted to PDF: {os.path.basename(pdf_path)}")
            else:
                failure_count += 1
                print(f"Failed to preprocess: {filename}")
    
    print(f"\nPreprocessing completed: {success_count} successes, {failure_count} failures")
    return (success_count, failure_count)

def permanently_preprocess_invoice_image(image_path):
    """
    Convert image file to PDF with basic preprocessing.
    """
    try:
        print(f"Converting image to PDF: {image_path}")
        
        # Read the image
        img = cv2.imread(image_path)
        if img is None:
            print(f"ERROR: Cannot read image: {image_path}")
            return None
        
        # Convert to grayscale for better compression
        if len(img.shape) > 2 and img.shape[2] > 1:
            img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        else:
            img_gray = img
        
        # Create PDF path
        pdf_path = os.path.splitext(image_path)[0] + ".pdf"
        
        # Create a PDF using PyMuPDF (fitz)
        doc = fitz.open()
        
        # Create a temporary file for the image
        with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_file:
            temp_image_path = temp_file.name
            cv2.imwrite(temp_image_path, img_gray, [cv2.IMWRITE_JPEG_QUALITY, 80])
        
        # Create a new page and insert the image
        page = doc.new_page()
        page.insert_image(fitz.Rect(0, 0, page.rect.width, page.rect.height), filename=temp_image_path)
        
        # Save the PDF
        doc.save(pdf_path, deflate=True, garbage=4, clean=True)
        doc.close()
        
        # Clean up
        os.remove(temp_image_path)
        os.remove(image_path)  # Remove original image
        
        print(f"Successfully converted to PDF: {pdf_path}")
        return pdf_path
        
    except Exception as e:
        print(f"Error converting image {image_path} to PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def process_invoices(invoice_folder):
    """Enhanced process_invoices function with precise pixel-based extraction."""
    print(f"DEBUG: Processing invoices in folder: {invoice_folder}")
    if not os.path.exists(invoice_folder):
        print(f"ERROR: Invoice folder does not exist: {invoice_folder}")
        return

    failed_invoices = []

    print("\nStarting preprocessing of invoices...")
    for filename in os.listdir(invoice_folder):
        file_path = os.path.join(invoice_folder, filename)
        
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            print(f"Preprocessing image file and converting to PDF: {filename}")
            pdf_path = permanently_preprocess_invoice_image(file_path)
            
            if pdf_path:
                print(f"Successfully converted to PDF: {os.path.basename(pdf_path)}")
        
        elif filename.lower().endswith('.pdf'):
            print(f"PDF file will be processed directly: {filename}")

    print("\nStarting automatic processing of invoices...")
    for filename in os.listdir(invoice_folder):
        if not filename.lower().endswith('.pdf'):
            continue
            
        file_path = os.path.join(invoice_folder, filename)
        print(f"Processing PDF file: {filename}")
        
        # try:
        doc = fitz.open(file_path)
        all_pages_processed = True
        
        for page_num in range(len(doc)):
            pix = doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2))
            temp_image_path = save_pixmap_with_retry(pix)
            
            # try:
            if not process_invoice_image(temp_image_path):
                failed_invoices.append(temp_image_path)
                all_pages_processed = False
            else:
                if os.path.exists(temp_image_path):
                    os.remove(temp_image_path)
            # except Exception as e:
            # print(f"Error processing PDF page {page_num + 1}: {str(e)}")
            all_pages_processed = False
        
        doc.close()
        
        if all_pages_processed:
            os.remove(file_path)
            print(f"Deleted original PDF: {file_path}")
                
        # except Exception as e:
        #     print(f"Error processing PDF {filename}: {str(e)}")

    if failed_invoices:
        print(f"\nFound {len(failed_invoices)} invoices that need manual processing.")
        print("Starting manual processing...\n")
        
        for failed_invoice in failed_invoices:
            if not os.path.exists(failed_invoice):
                continue
                
            try:
                manual_result = process_invoice_image(failed_invoice, manual_mode=True)
                
                if failed_invoice.startswith(tempfile.gettempdir()) and os.path.exists(failed_invoice):
                    os.remove(failed_invoice)
                    print(f"Removed temporary file: {failed_invoice}")
                    
            except Exception as e:
                print(f"Error during manual processing of {os.path.basename(failed_invoice)}: {str(e)}")
    
    else:
        print("\nAll invoices were processed automatically. No manual input needed.")

    print("\nInvoice processing completed.")

def add_customer_email(excel_file, customer_id, email):
    df = pd.read_excel(excel_file, header=None)
    new_entry = pd.DataFrame({0: [customer_id], 1: [email]})
    df = pd.concat([df, new_entry], ignore_index=True)
    df.to_excel(excel_file, index=False, header=False)
    print(f"Added {customer_id} with email {email} to the list.")

def update_customer_email(excel_file, customer_id, new_email):
    df = pd.read_excel(excel_file, header=None)
    df.loc[df[0] == customer_id, 1] = new_email
    df.to_excel(excel_file, index=False, header=False)
    print(f"Updated {customer_id}'s email to {new_email}.")

def remove_customer_email(excel_file, customer_id):
    df = pd.read_excel(excel_file, header=None)
    df = df[df[0] != customer_id]
    df.to_excel(excel_file, index=False, header=False)
    print(f"Removed {customer_id} from the list.")

def query_customer_purchases(customer_id=None, start_date=None, end_date=None):
    """Query purchases by customer with optional date range filtering."""
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

def query_product_sales(item_no=None, start_date=None, end_date=None):
    """Query sales data for specific products."""
    conn = sqlite3.connect('invoice_database.db')
    
    query = """
    SELECT it.item_no, 
           SUM(it.qty_shipped) AS total_quantity,
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

def generate_customer_activity_report():
    """Generate a report of customer purchasing activity."""
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
    """Test function for extracting data from a single invoice with precise pixel extraction."""
    print(f"Testing precise invoice extraction on: {image_path}")
    
    # If it's an image file, test the whitespace removal first
    if image_path.lower().endswith(('.png', '.jpg', '.jpeg')):
        print(f"\nTesting whitespace removal on image file...")
        test_preprocessed = permanently_preprocess_invoice_image(image_path)
        if test_preprocessed:
            print(f"Whitespace removal completed, now testing on: {test_preprocessed}")
            image_path = test_preprocessed  # Use the processed PDF
        else:
            print(f"Whitespace removal failed, using original image")
    
    # Read the image
    if image_path.lower().endswith('.pdf'):
        print(f"Loading PDF and extracting first page...")
        doc = fitz.open(image_path)
        if doc.page_count > 0:
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_file:
                temp_image_path = temp_file.name
            pix.save(temp_image_path)
            image = cv2.imread(temp_image_path)
            doc.close()
            cleanup_temp = lambda: os.path.exists(temp_image_path) and os.remove(temp_image_path)
        else:
            print(f"Error: PDF has no pages")
            return
    else:
        image = cv2.imread(image_path)
        cleanup_temp = lambda: None
    
    if image is None:
        print(f"Error: Could not read image from {image_path}")
        return
    
    print(f"Image dimensions: {image.shape[1]}x{image.shape[0]}")
    
    # Extract text using Google Vision API
    texts = extract_text_and_positions(image)
    
    # Print all detected text for debugging
    print(f"\nDEBUG: All detected text elements:")
    print("=" * 80)
    for i, text in enumerate(texts[1:], 1):  # Skip first element (full text)
        box = text.bounding_poly.vertices
        content = text.description.strip()
        print(f"{i:2d}. '{content}' at ({box[0].x:3d}, {box[0].y:3d}) - ({box[2].x:3d}, {box[2].y:3d})")
    print("=" * 80)
    
    # Align the image
    aligned_image, rotation_angle = align_invoice_image(image, texts)
    
    # Extract metadata using existing method
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
    print(f"Rotation Applied: {rotation_angle:.2f} degrees")
    
    # Extract line items using precise pixel method
    print(f"\nStarting line item extraction...")
    line_items = extract_line_items_precise(aligned_image, texts, debug=True)
    
    # Display results in invoice format
    if line_items:
        print(f"\nFINAL EXTRACTED LINE ITEMS:")
        print("=" * 80)
        print(f"{'Item Code':<10} | {'Qty Ord':<7} | {'Qty Ship':<8} | {'Extension':<10}")
        print("-" * 80)
        
        total = 0
        for item in line_items:
            ext_display = f"${item['extension']}" if item['extension'] else ""
            if item['extension']:
                try:
                    total += float(item['extension'])
                except:
                    pass
            
            print(f"{item['item_no']:<10} | {item['qty_ordered']:<7} | {item['qty_shipped']:<8} | {ext_display:<10}")
        
        print("-" * 80)
        print(f"{'TOTAL:':<35} ${total:.2f}")
        print("=" * 80)
        
        # Save to CSV for review
        df_data = []
        for item in line_items:
            df_data.append({
                "Item Code": item["item_no"],
                "Qty Ordered": item["qty_ordered"],
                "Qty Shipped": item["qty_shipped"],
                "Extension": f"${float(item['extension']):.2f}" if item["extension"] else ""
            })
        
        df = pd.DataFrame(df_data)
        df.to_csv("test_precise_extraction.csv", index=False)
        print("Line items saved to test_precise_extraction.csv")
        
    else:
        print("No line items extracted")
    
    # Clean up temp files
    cleanup_temp()
    
    # If you have valid metadata, try saving to database
    if invoice_num and customer_id and invoice_date and line_items:
        print(f"\nAttempting to save to database...")
        save_invoice_to_database_simplified(invoice_num, customer_id, invoice_date, line_items)

if __name__ == "__main__":
    print("\n" + "="*80)
    print("ENHANCED INVOICE PROCESSING SYSTEM WITH PRECISE PIXEL EXTRACTION")
    print("="*80)
    
    # Validate paths
    if not validate_paths():
        print("Exiting due to invalid paths.")
        exit(1)
    
    while True:
        # Display menu
        print("\nSelect an option:")
        print("1. Process all invoices in folder")
        print("2. Generate reports from database")
        print("3. Exit program")
        
        choice = input("\nEnter your choice (1-3): ")
        
        if choice == "1":
            # Process all invoices in the folder
            print("\nProcessing all invoices in folder:", invoice_folder)
            process_invoices(invoice_folder)
            check_inactive_customers()
            print("\nProcessing completed. Returning to main menu...")
                
        elif choice == "2":
            # Generate reports from database
            print("\n" + "-"*80)
            print("DATABASE REPORTS")
            print("-"*80)
            print("Select a report type:")
            print("1. Customer Purchase History")
            print("2. Product Sales Analysis")
            print("3. Customer Activity Report")
            print("4. Return to main menu")
            
            report_choice = input("\nEnter report choice (1-4): ")
            
            if report_choice == "1":
                customer_id = input("Enter customer ID (leave blank for all): ")
                if customer_id:
                    df = query_customer_purchases(customer_id)
                else:
                    df = query_customer_purchases()
                    
                if not df.empty:
                    print("\nCustomer Purchase History:")
                    print(df.to_string(index=False))
                    csv_path = f"customer_purchases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                    df.to_csv(csv_path, index=False)
                    print(f"\nReport saved to {csv_path}")
                else:
                    print("No purchase history found")
                    
            elif report_choice == "2":
                product = input("Enter product code (leave blank for all): ")
                if len(product) == 6 and product.isalnum():
                    df = query_product_sales(item_no=product)
                else:
                    df = query_product_sales()
                    
                if not df.empty:
                    print("\nProduct Sales Report:")
                    print(df.to_string(index=False))
                    csv_path = f"product_sales_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                    df.to_csv(csv_path, index=False)
                    print(f"\nReport saved to {csv_path}")
                else:
                    print("No product sales found")
                    
            elif report_choice == "3":
                df = generate_customer_activity_report()
                if not df.empty:
                    print("\nCustomer Activity Report:")
                    print(df.to_string(index=False))
                    csv_path = f"customer_activity_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                    df.to_csv(csv_path, index=False)
                    print(f"\nReport saved to {csv_path}")
                else:
                    print("No customer activity data found")
            
            elif report_choice == "4":
                print("Returning to main menu...")
                continue
                
        elif choice == "3":
            print("\nExiting program...")
            break
            
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")
    
    print("Program execution completed.")
