import os
import re
import cv2
import shutil
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.cloud import vision
from google.cloud.vision_v1 import types
from pdf2image import convert_from_path

# Directly set the path to the service account key file
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\Derek\Downloads\caramel-compass-429017-h3-c2d4e157e809.json"

# Initialize the Vision API client
client = vision.ImageAnnotatorClient()

# Path to the folder containing images
image_folder = r'C:\Users\Derek\Documents\Invoices\InvoicePictures'
destination_base_folder = r'C:\Users\Derek\Documents\Invoices\SortedInvoices'

# Email details
sender_email = "gingoso2@gmail.com"
app_password = "soiz avjw bdtu hmtn"

# Load customer emails from Excel file
def load_customer_emails(excel_file):
    df = pd.read_excel(excel_file, header=None)  # No header row
    customer_emails = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))  # First column as keys, second as values
    return customer_emails

# Test the function
customer_emails = load_customer_emails('customer_emails.xlsx')

# Function to rotate image 90 degrees counterclockwise
def rotate_image(image_path):
    image = cv2.imread(image_path)
    rotated_image = image

    # Display the rotated image
    cv2.imshow("Rotated Image", rotated_image)
    cv2.waitKey(0)  # Wait for a key press to close the image window
    cv2.destroyAllWindows()

    return rotated_image

# Function to extract text and bounding boxes using document_text_detection
def extract_text_and_positions(rotated_image):
    _, encoded_image = cv2.imencode('.jpg', rotated_image)
    content = encoded_image.tobytes()

    image = types.Image(content=content)

    # Perform document text detection on the image
    response = client.document_text_detection(image=image)
    texts = response.text_annotations

    if response.error.message:
        raise Exception(f'{response.error.message}')

    return texts

# Function to find text near a key phrase based on position
def find_text_near_keyphrase(texts, keyphrase, position):
    for text in texts:
        if keyphrase.lower() in text.description.lower():
            keyphrase_box = text.bounding_poly.vertices
            for adjacent_text in texts:
                adjacent_box = adjacent_text.bounding_poly.vertices
                if is_in_position(keyphrase_box, adjacent_box, position):
                    return adjacent_text.description
    return None

def is_in_position(box1, box2, position, threshold=50):
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
        return (box2_bottom < box1_top) and (abs(box2_left - box1_left) < threshold or abs(box2.right - box1.right) < threshold)
    elif position == 'left':
        return (box2_right < box1.left) and (abs(box2_top - box1.top) < threshold or abs(box2_bottom - box1.bottom) < threshold)
    elif position == 'right':
        return (box2_left > box1.right) and (abs(box2.top - box1.top) < threshold or abs(box2.bottom - box1.bottom) < threshold)
    return False

# Function to check if a string is a 6-digit number
def is_six_digit_number(s):
    return s.isdigit() and len(s) == 6

def is_mixed_string(s):
    has_digit = any(char.isdigit() for char in s)
    has_alpha = any(char.isalpha() for char in s)
    return has_digit and has_alpha

def is_date_format(s):
    # Define the regex pattern for mm/dd/yy
    pattern = re.compile(r'^(0[1-9]|1[0-2])/(0[1-9]|[12][0-9]|3[01])/([0-9]{2})$')
    # Check if the string matches the pattern
    return bool(pattern.match(s))

# mm/dd/yy format
def extract_year(date_str):
    return date_str.split('/')[2]

def extract_month(date_str):
    return date_str.split('/')[0]

def extract_day(date_str):
    return date_str.split('/')[1]

# Function to copy the image to the sorted folder structure
def copy_image_to_sorted_folder(image_path, six_digit_number, customer_id, date):
    year = extract_year(date)
    month = extract_month(date)
    day = extract_day(date)
    
    # Create folder structure
    destination_folder = os.path.join(destination_base_folder, year, month, day)
    os.makedirs(destination_folder, exist_ok=True)
    
    # Create destination path
    destination_path = os.path.join(destination_folder, f"{customer_id}_{six_digit_number}.jpg")
    
    # Copy and rename the image
    shutil.copy2(image_path, destination_path)
    print(f"Copied and renamed image to: {destination_path}")

    # Send the email with attachment
    send_email_with_attachment(destination_path, customer_id, six_digit_number, date)

# Function to send an email with the sorted image as an attachment
def send_email_with_attachment(file_path, customer_id, invoice_number, date):
    receiver_email = customer_emails.get(customer_id, None)
    if receiver_email is None:
        print(f"No email found for customer {customer_id}. Skipping email.")
        return

    # Email content
    subject = f"Invoice {invoice_number} for Customer {customer_id} dated {date}"
    body = f"Attached is the sorted invoice {invoice_number} for customer {customer_id} dated {date}."

    # Set up the MIME
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject

    # Attach the body with the msg instance
    message.attach(MIMEText(body, "plain"))

    # Open the file in binary mode
    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(file_path)}",
    )

    # Attach the part to the message
    message.attach(part)

    # Create a secure SSL context and send the email
    try:
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(sender_email, app_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print(f"Email sent successfully to {receiver_email}!")
    except Exception as e:
        print(f"Failed to send email: {e}")
    finally:
        server.quit()

# Function to convert PDF to images
def convert_pdf_to_images(pdf_path):
    images = convert_from_path(pdf_path)
    image_paths = []
    for i, image in enumerate(images):
        image_path = os.path.join(os.path.dirname(pdf_path), f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page_{i + 1}.png")
        image.save(image_path, "PNG")
        image_paths.append(image_path)
    return image_paths

def convert_all_pdfs(image_folder):
    pdf_image_paths = set()
    for filename in os.listdir(image_folder):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(image_folder, filename)
            image_paths = convert_pdf_to_images(pdf_path)
            pdf_image_paths.update(image_paths)
    return pdf_image_paths

# Main processing function
def process_images(image_folder):
    # Convert all PDFs to images first
    pdf_image_paths = convert_all_pdfs(image_folder)

    # Get all images (including converted PDFs, existing PNGs, and JPGs)
    all_image_paths = set(os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.lower().endswith(('.jpg', '.png', '.jpeg')))
    all_image_paths.update(pdf_image_paths)

    for image_path in all_image_paths:
        # Rotate the image
        rotated_image = rotate_image(image_path)
        
        # Extract texts and their positions
        texts = extract_text_and_positions(rotated_image)
        
        final_invoice_num = None
        final_customer_num = None
        final_date_num = None

        invoice_num = find_text_near_keyphrase(texts, "Invoice", "below")
        if invoice_num and is_six_digit_number(invoice_num):
            final_invoice_num = invoice_num

        customer_num = find_text_near_keyphrase(texts, "Account", "below")
        if customer_num and is_mixed_string(customer_num):
            final_customer_num = customer_num

        date_num = find_text_near_keyphrase(texts, "Date", "below")
        if date_num and is_date_format(date_num):
            final_date_num = date_num

        if final_date_num and final_customer_num and final_invoice_num:
            copy_image_to_sorted_folder(image_path, final_invoice_num, final_customer_num, final_date_num)

        print("="*40)  # Separator for readability

# Run the main processing function
process_images(image_folder)

# Functions to update the customer email list
def add_customer_email(excel_file, customer_id, email):
    df = pd.read_excel(excel_file, header=None)  # No header row
    new_entry = pd.DataFrame({0: [customer_id], 1: [email]})
    df = df.append(new_entry, ignore_index=True)
    df.to_excel(excel_file, index=False)
    print(f"Added {customer_id} with email {email} to the list.")

def update_customer_email(excel_file, customer_id, new_email):
    df = pd.read_excel(excel_file, header=None)  # No header row
    df.loc[df[0] == customer_id, 1] = new_email
    df.to_excel(excel_file, index=False)
    print(f"Updated {customer_id}'s email to {new_email}.")

def remove_customer_email(excel_file, customer_id):
    df = pd.read_excel(excel_file, header=None)  # No header row
    df = df[df[0] != customer_id]
    df.to_excel(excel_file, index=False)
    print(f"Removed {customer_id} from the list.")
