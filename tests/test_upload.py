import subprocess
import sys
import os

# Path to your Python interpreter (same as in master_script)
PYTHON_PATH = r"C:\Users\AliAmani\Miniconda3\envs\bids\python.exe"

# Folder to upload
folder_to_upload = "2024-11-08\\03_TXSMartBuy_COMPLETED"

# Verify folder exists
if not os.path.exists(folder_to_upload):
    print(f"❌ Folder not found: {folder_to_upload}")
    sys.exit(1)

print(f"Found folder to upload: {folder_to_upload}")
print(f"Current working directory: {os.getcwd()}")

try:
    # Run upload_bids.py with the specified folder without capturing output
    cmd = [PYTHON_PATH, "upload_bids.py", folder_to_upload]
    print(f"Running command: {' '.join(cmd)}")

    process = subprocess.run(
        cmd,
        check=False,  # Don't raise exception on non-zero return code
        shell=True,  # Run in shell to see output in real-time
    )

    if process.returncode == 0:
        print("✅ Upload completed successfully!")
    else:
        print(f"❌ Upload failed with return code: {process.returncode}")

except Exception as e:
    print(f"❌ Unexpected error: {str(e)}")
