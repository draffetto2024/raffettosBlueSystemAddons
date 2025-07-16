import os
import shutil
from PyInstaller.__main__ import run

# Only build this script
script_name = "CopyPastePacker.py"

# Custom output folder
dist_dir = "dist_copypastepackerONLY"

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
    output_name += "_Windowed" if is_windowed else "_Console"

    pyinstaller_command = [
        f'--name={output_name}',
        f'--distpath={dist_dir}',
        '--onefile',
        '--windowed' if is_windowed else '',
        *common_options,
        script_name
    ]

    pyinstaller_command = [opt for opt in pyinstaller_command if opt]

    print(f"Building {'windowed' if is_windowed else 'console'} version of {script_name}...")
    run(pyinstaller_command)

if __name__ == '__main__':
    build_executable(script_name, is_windowed=True)
    build_executable(script_name, is_windowed=False)

    # Copy additional required files
    shutil.copy('CopyPastePack.db', dist_dir)
    shutil.copy('UPCCodes.xlsx', dist_dir)

    print(f"\nBuild complete. Executables and necessary files are in the '{dist_dir}' folder.")
    print(f"  CopyPastePacker_Windowed.exe (Windowed version)")
    print(f"  CopyPastePacker_Console.exe (Console version for debugging)")
