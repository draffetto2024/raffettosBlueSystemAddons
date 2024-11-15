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
    
    # Now search for text near the keyphrase
    for adjacent_text in texts:
        if adjacent_text in keyphrase_texts:
            continue  # Skip the keyphrase itself
        adjacent_box = adjacent_text.bounding_poly.vertices
        
        if is_in_position(keyphrase_box, adjacent_box, position, threshold):
            # Check if the adjacent text contains alphanumeric characters
            if re.search(r'[a-zA-Z0-9]', adjacent_text.description):
                logging.info(f"\nMATCH FOUND: '{adjacent_text.description}'")
                logging.info(f"Match bounding box:")
                logging.info(f"  Top-left: ({adjacent_box[0].x}, {adjacent_box[0].y})")
                logging.info(f"  Bottom-right: ({adjacent_box[2].x}, {adjacent_box[2].y})")
                return adjacent_text.description
    
    logging.info(f"No alphanumeric matching text found near '{keyphrase}'")
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

def is_six_alphanumeric(s):
    # Strip any leading or trailing whitespace
    s = s.strip()
    
    # Debug output to check what is being passed
    print(f"DEBUG: Checking if '{s}' is six alphanumeric characters")

    # Check if it matches exactly 6 alphanumeric characters
    return bool(re.match(r'^[A-Za-z0-9]{6}$', s))


def is_date_format(s):
    pattern = re.compile(r'^(0[1-9]|1[0-2])/(0[1-9]|[12][0-9]|3[01])/([0-9]{2})$')
    return bool(pattern.match(s))

def extract_year(date_str):
    return date_str.split('/')[2]

def extract_month(date_str):
    return date_str.split('/')[0]

def extract_day(date_str):
    return date_str.split('/')[1]

def copy_image_to_sorted_folder(image_path, six_digit_number, customer_id, date):
    year, month, day = extract_year(date), extract_month(date), extract_day(date)
    destination_folder = os.path.join(destination_base_folder, year, month, day)
    os.makedirs(destination_folder, exist_ok=True)
    destination_path = os.path.join(destination_folder, f"{customer_id}_{six_digit_number}.jpg")
    print(f"DEBUG: Attempting to copy from {image_path} to {destination_path}")
    try:
        shutil.copy2(image_path, destination_path)
        print(f"Copied and renamed image to: {destination_path}")
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

def process_invoice_image(image_path):
    logging.info(f"DEBUG: Processing image: {image_path}")
    image = cv2.imread(image_path)
    texts = extract_text_and_positions(image)
    
    # Log the full texts output
    logging.info("Full OCR Results:")
    for i, text in enumerate(texts):
        logging.info(f"Text {i}: '{text.description}'")
        box = text.bounding_poly.vertices
        logging.info(f"  Bounding box: Top-left ({box[0].x}, {box[0].y}), Bottom-right ({box[2].x}, {box[2].y})")
    logging.info("End of OCR Results")

    invoice_num = find_text_near_keyphrase(texts, "INVOICE NO", "below", 50)
    logging.info(f"DEBUG: Text found below INVOICE NO: {invoice_num}")

    # Use expanding threshold for ACCOUNT NO
    customer_num = None
    for threshold in range(5, 101, 5):  # Start from 5, increment by 5, up to 100
        customer_num = find_text_near_keyphrase(texts, "ACCOUNT NO", "below", threshold)
        if customer_num:
            logging.info(f"DEBUG: Text found below ACCOUNT NO with threshold {threshold}: {customer_num}")
            break
    if not customer_num:
        logging.info("DEBUG: No alphanumeric text found below ACCOUNT NO even with increased threshold")

    date_num = find_text_near_keyphrase(texts, "INVOICE DATE", "below", 50)
    logging.info(f"DEBUG: Text found below INVOICE DATE: {date_num}")

    logging.info(f"DEBUG: Found invoice_num: {invoice_num}, customer_num: {customer_num}, date_num: {date_num}")

    if (invoice_num and is_six_alphanumeric(invoice_num) and
        customer_num and is_six_alphanumeric(customer_num) and
        date_num and is_date_format(date_num)):
        logging.info(f"DEBUG: All conditions met, copying image to sorted folder")
        copy_image_to_sorted_folder(image_path, invoice_num, customer_num, date_num)
        return True
    else:
        logging.info(f"DEBUG: Conditions not met for sorting, copying to unsorted folder")
        copy_image_to_unsorted_folder(image_path)
        return False

def process_invoice(invoice_path):
    print(f"DEBUG: Processing invoice: {invoice_path}")
    if invoice_path.lower().endswith('.pdf'):
        image_paths = convert_pdf_to_images(invoice_path)
        print(f"DEBUG: Converted PDF to images: {image_paths}")

        processed_invoices = 0
        for image_path in image_paths:
            if process_invoice_image(image_path):
                processed_invoices += 1
            os.remove(image_path)  # Remove temporary image file
            print(f"DEBUG: Removed temporary image: {image_path}")

        print(f"DEBUG: Processed {processed_invoices} invoices from PDF: {invoice_path}")
    else:
        process_invoice_image(invoice_path)

    # Do not delete the original invoice after processing
    # os.remove(invoice_path)
    # print(f"DEBUG: Deleted original invoice: {invoice_path}")

    print("="*40)  # Separator for readability

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
    print(f"DEBUG: Processing invoices in folder: {invoice_folder}")
    if not os.path.exists(invoice_folder):
        print(f"ERROR: Invoice folder does not exist: {invoice_folder}")
        return

    for filename in os.listdir(invoice_folder):
        print(f"DEBUG: Found file: {filename}")
        file_path = os.path.join(invoice_folder, filename)

        if filename.lower().endswith('.pdf'):
            print(f"DEBUG: Processing PDF file: {filename}")
            try:
                process_pdf_invoice(file_path)
                # Delete the original PDF file after successful processing
                os.remove(file_path)
                print(f"DEBUG: Deleted original PDF: {file_path}")
            except Exception as e:
                print(f"ERROR: Failed to process PDF {filename}: {str(e)}")
                print(f"DEBUG: Original PDF not deleted due to processing error: {file_path}")
        elif filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            print(f"DEBUG: Processing image file: {filename}")
            try:
                process_invoice_image(file_path)
                # Optionally, you can delete the original image file here as well
                os.remove(file_path)
                print(f"DEBUG: Deleted original image: {file_path}")
            except Exception as e:
                print(f"ERROR: Failed to process image {filename}: {str(e)}")
        else:
            print(f"DEBUG: Skipped non-invoice file: {filename}")

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
    process_invoices(invoice_folder)
    print("Program execution completed.")

    input("Press Enter to exit...")
    
    print("DEBUG: Finished invoice processing")
