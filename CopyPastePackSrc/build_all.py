import os
import subprocess
import sys

def build_executable(script_name, data_files):
    # Determine the appropriate separator based on the operating system
    separator = ';' if sys.platform.startswith('win') else ':'
    
    # Create the --add-data arguments
    add_data_args = [f'--add-data={file}{separator}.' for file in data_files]
    
    cmd = [
        "pyinstaller",
        "--onefile",
        "--name", os.path.splitext(script_name)[0],
    ] + add_data_args + [script_name]
    
    subprocess.run(cmd, check=True)

# List of scripts and their associated data files
scripts_and_data = [
    ("CopyPastePacker.py", ["UPCCodes.xlsx", "input.txt", "CopyPastePack.db"]),
    ("ConfigSetup.py", ["UPCCodes.xlsx", "input.txt", "CopyPastePack.db"]),
    ("DataDashboard.py", ["UPCCodes.xlsx", "input.txt", "CopyPastePack.db"])
]

# Build each executable
for script, data_files in scripts_and_data:
    print(f"Building {script}...")
    build_executable(script, data_files)

print("All executables built successfully!")

# Create distribution folder
dist_folder = "DistributionPackage"
os.makedirs(dist_folder, exist_ok=True)

# Move executables and data files to distribution folder
for script, data_files in scripts_and_data:
    executable_name = os.path.splitext(script)[0] + (".exe" if sys.platform.startswith('win') else "")
    src_path = os.path.join("dist", executable_name)
    dst_path = os.path.join(dist_folder, executable_name)
    if os.path.exists(src_path):
        os.rename(src_path, dst_path)
    for data_file in data_files:
        if os.path.exists(data_file):
            os.rename(data_file, os.path.join(dist_folder, data_file))

print(f"Distribution package created in {dist_folder}")