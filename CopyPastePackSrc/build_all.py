import os
import sys
import shutil
from PyInstaller.__main__ import run

# Define the names of your main scripts
main_scripts = ["CopyPastePacker.py", "ConfigSetup.py", "DataDashboard.py"]

# Common PyInstaller options
common_options = [
    '--add-data=UPCCodes.xlsx:.',
    '--add-data=CopyPastePack.db:.',
    '--hidden-import=pandas',
    '--hidden-import=sqlite3',
    '--hidden-import=tkinter',
    '--hidden-import=re',
]

def build_executable(script_name, is_windowed):
    output_name = os.path.splitext(script_name)[0]
    if is_windowed:
        output_name += "_Windowed"
    else:
        output_name += "_Console"

    pyinstaller_command = [
        '--name=%s' % output_name,
        '--onefile',
    ]

    if is_windowed:
        pyinstaller_command.append('--windowed')

    pyinstaller_command.extend(common_options)
    pyinstaller_command.append(script_name)

    print(f"Building {'windowed' if is_windowed else 'console'} version of {script_name}...")
    run(pyinstaller_command)

# Run PyInstaller for all scripts
if __name__ == '__main__':
    for script in main_scripts:
        build_executable(script, is_windowed=True)
        build_executable(script, is_windowed=False)

    # After PyInstaller finishes, copy necessary files to the dist folder
    dist_dir = 'dist'
   
    # Copy CopyPastePack.db
    shutil.copy('CopyPastePack.db', dist_dir)
   
    # Copy UPCCodes.xlsx
    shutil.copy('UPCCodes.xlsx', dist_dir)

    print(f"Build complete. Executables and necessary files can be found in {dist_dir}")
    print("The following executables have been created:")
    for script in main_scripts:
        base_name = os.path.splitext(script)[0]
        print(f"  {base_name}_Windowed.exe (Windowed version)")
        print(f"  {base_name}_Console.exe (Console version for debugging)")