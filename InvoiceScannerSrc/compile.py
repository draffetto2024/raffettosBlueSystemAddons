import PyInstaller.__main__
import os
import shutil

# Define the paths
script_path = "InvoiceImageToText.py"
json_path = r".\caramel-compass-429017-h3-c2d4e157e809.json"
output_dir = "dist"

# Create the output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Compile the script
PyInstaller.__main__.run([
    script_path,
    "--onefile",
    "--add-data", f"{json_path};.",
    "--distpath", output_dir, 
    "--name", "InvoiceProcessor"
])

# Copy the JSON file to the dist folder
shutil.copy2(json_path, os.path.join(output_dir, "service_account_key.json"))

print("Compilation complete. Executable and JSON file are in the 'dist' folder.")
