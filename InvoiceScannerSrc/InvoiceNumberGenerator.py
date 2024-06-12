import os
import pytesseract
from PIL import Image
import re
import datetime

import cv2
import numpy as np

# Path to the folder containing images
image_folder = r'C:\Users\Derek\Documents\InvoicePictures'

# Path to pytesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR/tesseract'

# Loop through all the images in the folder
for filename in os.listdir(image_folder):
    # Check if the file is an image
    if filename.lower().endswith(('.jpg', '.png', '.jpeg')):
        # Load the image
        image_path = os.path.join(image_folder, filename)
        image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

        # Binarize the image
        _, binary_image = cv2.threshold(image, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

        # Remove noise
        binary_image = cv2.medianBlur(binary_image, 3)

        # Detect and correct skew
        coords = np.column_stack(np.where(binary_image > 0))
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle

        (h, w) = binary_image.shape[:2]
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, angle, 1.0)
        deskewed_image = cv2.warpAffine(binary_image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

        # Perform OCR on the preprocessed image
        text = pytesseract.image_to_string(deskewed_image)
        
        print(f"Text extracted from {filename}:")
        print(text)
        print("="*40)  # Separator for readability


# # Loop through all the images in the folder
# for filename in os.listdir(image_folder):
#     # Check if the file is an image
#     if filename.endswith('.jpg') or filename.endswith('.png') or filename.endswith('.jpeg'):
#         # Load the image
#         img = Image.open(os.path.join(image_folder, filename))
#         # Get the width and height of the image
#         width, height = img.size
#         # Define the region of interest (ROI) where the invoice number is located (bottom right corner)
#         roi = (int(0.6 * width), int(0.8 * height), width, height)
#         # Crop the image to the ROI
#         cropped_img = img.crop(roi)
#         # Perform OCR on the cropped image
#         ocr_result = pytesseract.image_to_string(cropped_img)
#         print(ocr_result)
#         # Extract the invoice number and date using regular expressions
#         invoice_number = None
#         invoice_date = None
#         if 'invoice date' in ocr_result.lower():
#             invoice_number_match = re.search(r'\d{6}', ocr_result)
#             if invoice_number_match:
#                 invoice_number = invoice_number_match.group(0)
#             invoice_date_match = re.search(r'\d{1}/\d{2}/\d{2}', ocr_result) #d{2} controls how many month digits it looks for
#             if invoice_date_match:
#                 invoice_date_str = invoice_date_match.group(0)
#                 invoice_date = datetime.datetime.strptime(invoice_date_str, '%m/%d/%y')
#         if invoice_number:
#             # Rename the original image with the invoice number and date (if found)
#             if invoice_date:
#                 new_filename = f'{invoice_number}_{invoice_date.strftime("%m-%d-%y")}.jpg'
#             else:
#                 new_filename = f'{invoice_number}.jpg'
#             os.rename(os.path.join(image_folder, filename), os.path.join(image_folder, new_filename))

