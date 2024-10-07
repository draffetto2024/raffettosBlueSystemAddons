import os
import sys
import shutil
from PyInstaller.__main__ import run

# Define the name of your main script
main_script = "EmailToOrders.py"

# Define the names of your output executables
output_name_windowed = "EmailOrderProcessor"
output_name_console = "EmailOrderProcessor_Console"

# Common PyInstaller options
common_options = [
    '--add-data=customer_product_codes:customer_product_codes',
    '--add-data=orders.db:.',
    '--hidden-import=pandas',
    '--hidden-import=nltk',
    '--hidden-import=sqlite3',
    '--hidden-import=tkcalendar',
    '--hidden-import=pyautogui',
    '--hidden-import=babel.numbers',
    '--hidden-import=babel.dates',
    '--collect-submodules=babel',
    '--collect-data=babel',
]

# PyInstaller command for windowed version
pyinstaller_command_windowed = [
    '--name=%s' % output_name_windowed,
    '--onefile',
    '--windowed',
] + common_options + [main_script]

# PyInstaller command for console version
pyinstaller_command_console = [
    '--name=%s' % output_name_console,
    '--onefile',
] + common_options + [main_script]

# Run PyInstaller for both versions
if __name__ == '__main__':
    print("Building windowed version...")
    run(pyinstaller_command_windowed)
   
    print("Building console version...")
    run(pyinstaller_command_console)

    # After PyInstaller finishes, copy necessary files to the dist folder
    dist_dir = 'dist'
   
    # Copy orders.db
    shutil.copy('orders.db', dist_dir)
   
    # Copy customer_product_codes folder
    customer_codes_dest = os.path.join(dist_dir, 'customer_product_codes')
    if os.path.exists(customer_codes_dest):
        shutil.rmtree(customer_codes_dest)
    shutil.copytree('customer_product_codes', customer_codes_dest)

    print(f"Build complete. Executables and necessary files can be found in {dist_dir}")
    print(f"Run {output_name_console}.exe to see console output and error messages.")
    print(f"Run {output_name_windowed}.exe for the windowed application.")