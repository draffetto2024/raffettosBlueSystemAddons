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

# Directly set the path to the service account key file
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\Derek\Downloads\caramel-compass-429017-h3-c2d4e157e809.json"

# Initialize the Vision API client
client = vision.ImageAnnotatorClient()

# Path to the folders
invoice_folder = r'C:\Users\Derek\Documents\Invoices\InvoicePictures'
destination_base_folder = r'C:\Users\Derek\Documents\Invoices\SortedInvoices'
unsorted_base_folder = r'C:\Users\Derek\Documents\Invoices\UnsortedInvoices'

# Email details
sender_email = "gingoso2@gmail.com"
app_password = "soiz avjw bdtu hmtn"

def load_customer_emails(excel_file):
    df = pd.read_excel(excel_file, header=None)  # No header row
    customer_emails = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))  # First column as keys, second as values
    return customer_emails

customer_emails = load_customer_emails('customer_emails.xlsx')

def rotate_image(image_path):
    image = cv2.imread(image_path)
    rotated_image = image

    # Display the rotated image
    cv2.imshow("Rotated Image", rotated_image)
    cv2.waitKey(0)  # Wait for a key press to close the image window
    cv2.destroyAllWindows()

    return rotated_image

def extract_text_and_positions(image):
    _, encoded_image = cv2.imencode('.jpg', image)
    content = encoded_image.tobytes()
    image = types.Image(content=content)
    response = client.document_text_detection(image=image)
    if response.error.message:
        raise Exception(f'{response.error.message}')
    return response.text_annotations

def find_text_near_keyphrase(texts, keyphrase, position, threshold):
    for text in texts:
        if keyphrase.lower() in text.description.lower():
            keyphrase_box = text.bounding_poly.vertices
            for adjacent_text in texts:
                adjacent_box = adjacent_text.bounding_poly.vertices
                if is_in_position(keyphrase_box, adjacent_box, position, threshold):
                    return adjacent_text.description
    return None

def is_in_position(box1, box2, position, threshold):
    box1_top = min(v.y for v in box1)
    box1_bottom = max(v.y for v in box1)
    box1_left = min(v.x for v in box1)
    box1_right = max(v.x for v in box1)

    box2_top = min(v.y for v in box2)
    box2_bottom = max(v.y for v in box2)
    box2_left = min(v.x for v in box2)
    box2_right = max(v.x for v in box2)

    if position == 'below':
        return (box2_top > box1_bottom) and (abs(box2_left - box1_left) < threshold or abs(box2_right - box1_right) < threshold)
    elif position == 'above':
        return (box2_bottom < box1_top) and (abs(box2_left - box1_left) < threshold or abs(box2_right - box1_right) < threshold)
    elif position == 'left':
        return (box2_right < box1_left) and (abs(box2_top - box1_top) < threshold or abs(box2_bottom - box1_bottom) < threshold)
    elif position == 'right':
        return (box2_left > box1_right) and (abs(box2_top - box1_top) < threshold or abs(box2_bottom - box1_bottom) < threshold)
    return False

def is_six_alphanumeric(s):
    return len(s) == 6 and s.isalnum()

def is_mixed_string(s):
    has_digit = any(char.isdigit() for char in s)
    has_alpha = any(char.isalpha() for char in s)
    return has_digit and has_alpha

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

def convert_pdf_to_images(pdf_path):
    doc = fitz.open(pdf_path)
    image_paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap()
        image_path = f"{pdf_path}_page_{i+1}.png"
        pix.save(image_path)
        image_paths.append(image_path)
    doc.close()
    return image_paths

def process_invoice(invoice_path):
    print(f"DEBUG: Processing invoice: {invoice_path}")
    if invoice_path.lower().endswith('.pdf'):
        image_paths = convert_pdf_to_images(invoice_path)
        print(f"DEBUG: Converted PDF to images: {image_paths}")
    else:
        image_paths = [invoice_path]

    for image_path in image_paths:
        print(f"DEBUG: Processing image: {image_path}")
        rotated_image = rotate_image(image_path)
        texts = extract_text_and_positions(rotated_image)

        invoice_num = find_text_near_keyphrase(texts, "Invoice", "below", 10)
        customer_num = find_text_near_keyphrase(texts, "Account", "below", 100)
        date_num = find_text_near_keyphrase(texts, "Date", "below", 10)
        
        print(f"DEBUG: Found invoice_num: {invoice_num}, customer_num: {customer_num}, date_num: {date_num}")
        
        if (invoice_num and is_six_alphanumeric(invoice_num) and
            customer_num and is_six_alphanumeric(customer_num) and   #cust id an invoice are 6 digit strings
            date_num and is_date_format(date_num)):
            print(f"DEBUG: All conditions met, copying image to sorted folder")
            copy_image_to_sorted_folder(image_path, invoice_num, customer_num, date_num)
        else:
            print(f"DEBUG: Conditions not met for sorting, copying to unsorted folder")
            copy_image_to_unsorted_folder(image_path)
        
        if invoice_path.lower().endswith('.pdf'):
            os.remove(image_path)  # Remove temporary image file
            print(f"DEBUG: Removed temporary image: {image_path}")

    print("="*40)  # Separator for readability

def process_invoices(invoice_folder):
    print(f"DEBUG: Processing invoices in folder: {invoice_folder}")
    for filename in os.listdir(invoice_folder):
        if filename.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
            invoice_path = os.path.join(invoice_folder, filename)
            print(f"DEBUG: Found invoice: {invoice_path}")
            process_invoice(invoice_path)

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
    print("DEBUG: Finished invoice processing")