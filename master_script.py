import os
import subprocess
import threading
import time
import queue
from datetime import datetime, timedelta
import sys
import io
import signal
import msvcrt
import pandas as pd
import shutil
import winsound
import argparse
from enum import Enum
import ctypes
import atexit
import glob
from pathlib import Path
import logging
from utils.excel_processor import ExcelProcessor
import traceback
from rich.console import Console
from rich.progress import (
    Progress,
    SpinnerColumn,
    TextColumn,
    BarColumn,
    TimeElapsedColumn,
)
from rich.table import Table
from rich.panel import Panel
from rich.live import Live
from rich.logging import RichHandler
from typing import Tuple

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# Add this class definition before it's used
class ScriptStatus(Enum):
    PENDING = "Pending"
    RUNNING = "Running"
    SUCCESS = "Done"
    ERROR = "Done"

# Environment Configuration
PYTHON_PATH = r"C:\Users\AliAmani\Miniconda3\envs\bids\python.exe"  # Path to Python executable
CONDA_PATH = r"C:\Users\AliAmani\Miniconda3"  # Path to Miniconda installation
CONDA_ENV = "bids"  # Conda environment name
CONDA_ENV_PATH = r"C:\Users\AliAmani\Miniconda3\envs\bids"  # Path to Conda environment

# Script Configuration
MAX_CONCURRENT_SCRIPTS = 4  # Number of scripts to run simultaneously
UPLOAD_SCRIPT = "upload_bids.py"  # Script for uploading data

# Script Order and Lists
SCRIPT_ORDER = [
    "scrapers/01_BuySpeed_01.py",
    "scrapers/01_BuySpeed_02.py",
    "scrapers/02_NYC.py",
    "scrapers/03_TXSMartBuy.py",
    "scrapers/05_NYSCR.py",
    "scrapers/06_MyFloridaMarketPlace.py",
    "scrapers/07_StateOfGeorgia.py",
    "scrapers/08_SFCityPartner.py",
    "scrapers/09_CGIEVA.py",
    "scrapers/10_BidBuysIllinoise.py",
    "scrapers/11_PlanetBids_Hartford.py",
    "scrapers/12_Bonfire_FairfaxCounty_1.py",
    "scrapers/12_Bonfire_FairfaxCounty_2.py",
    "scrapers/12_Bonfire_FairfaxCounty_3.py",
    "scrapers/12_Bonfire_FairfaxCounty_4.py",
    "scrapers/13_eMaryland_eMMA.py",
    "scrapers/14_NorthCarolina_VendorPortal_eVP.py",
    "scrapers/15_State_of_Conneticut_BidBoard.py",
    "scrapers/16_CalProcure.py",
    "scrapers/17_BidNet.py",
    "scrapers/18_Ionwave.py",
    "scrapers/19_Pennsylvania_eMarketplace.py",
    "scrapers/20_County_of_San_Diego.py",
]

scripts = SCRIPT_ORDER.copy()  # Create a copy of script order for the scripts list

# Global Variables
yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
terminate_flag = threading.Event()
script_queue = queue.Queue()
print_lock = threading.Lock()
running_processes = {}  # Dictionary to keep track of running processes
auto_started_scripts = set()  # Track scripts that have been started
active_scripts = set()  # Track currently running scripts
active_scripts_lock = threading.Lock()
script_semaphore = threading.Semaphore(MAX_CONCURRENT_SCRIPTS)
script_progress = None
main_log_buffer = []  # Store main terminal logs
processed_excel_folders = set()  # Track which folders have been processed
processed_scripts = set()  # Track which scripts have been fully completed

# Initialize script queue at startup
for script in scripts:
    script_queue.put(script)

# Initialize script statuses
script_statuses = {script: ScriptStatus.PENDING for script in scripts}

# Add this near the top of the file, after other global variables
UPLOAD_SCRIPT = (
    "upload_bids.py"  # This can be easily changed to a different script name
)

# List of scripts to run
scripts = [
    "scrapers/01_BuySpeed_01.py",
    "scrapers/01_BuySpeed_02.py",
    "scrapers/02_NYC.py",
    "scrapers/03_TXSMartBuy.py",
    "scrapers/05_NYSCR.py",
    "scrapers/06_MyFloridaMarketPlace.py",
    "scrapers/07_StateOfGeorgia.py",
    "scrapers/08_SFCityPartner.py",
    "scrapers/09_CGIEVA.py",
    "scrapers/10_BidBuysIllinoise.py",
    "scrapers/11_PlanetBids_Hartford.py",
    "scrapers/12_Bonfire_FairfaxCounty_1.py",
    "scrapers/12_Bonfire_FairfaxCounty_2.py",
    "scrapers/12_Bonfire_FairfaxCounty_3.py",
    "scrapers/12_Bonfire_FairfaxCounty_4.py",
    "scrapers/13_eMaryland_eMMA.py",
    "scrapers/14_NorthCarolina_VendorPortal_eVP.py",
    "scrapers/15_State_of_Conneticut_BidBoard.py",
    "scrapers/16_CalProcure.py",
    "scrapers/17_BidNet.py",
    "scrapers/18_Ionwave.py",
    "scrapers/19_Pennsylvania_eMarketplace.py",
    "scrapers/20_County_of_San_Diego.py",
]

# Queue to manage scripts
script_queue = queue.Queue()
for script in scripts:
    script_queue.put(script)

# Lock for thread-safe printing
print_lock = threading.Lock()

# Global flag to signal script termination
terminate_flag = threading.Event()

# Dictionary to keep track of running processes
running_processes = {}

# Semaphore to limit concurrent scripts
max_concurrent_scripts = 4
script_semaphore = threading.Semaphore(max_concurrent_scripts)

# Dictionary to track script statuses
script_statuses = {script: ScriptStatus.PENDING for script in scripts}

# Add near the top with other global variables
auto_started_scripts = set()  # Track scripts that have been started

# Move helper functions to the top, after imports and before global variables
def all_scripts_completed() -> bool:
    """Check if all scripts have completed (either success or error) or been terminated"""
    for script in scripts:
        # Skip scripts that were never started if we're terminating
        if terminate_flag.is_set() and script not in auto_started_scripts:
            continue
            
        # Check if script has been fully processed
        if script not in processed_scripts:
            return False
                
    return True

def check_and_start_processing():
    """Check if all scripts are done and start Excel processing and uploads"""
    try:
        # Double check we should start processing
        if not all_scripts_completed():
            return
            
        print("\n🔄 Starting batch processing...")
        
        # Get yesterday's date folder
        yesterday_folder = os.path.join(os.getcwd(), yesterday)
        
        if not os.path.exists(yesterday_folder):
            print(f"❌ Yesterday's folder not found: {yesterday_folder}")
            return
            
        # Find all COMPLETED folders that haven't been processed yet
        completed_folders = []
        for root, dirs, files in os.walk(yesterday_folder):
            for dir_name in dirs:
                folder_path = os.path.join(root, dir_name)
                if dir_name.endswith('_COMPLETED') and folder_path not in processed_excel_folders:
                    completed_folders.append(folder_path)
        
        if not completed_folders:
            print("No new COMPLETED folders to process")
            return
            
        print(f"Found {len(completed_folders)} new COMPLETED folders")
        
        # Process all Excel files first
        print("\n📊 Processing all Excel files...")
        for folder in completed_folders:
            if folder in processed_excel_folders:
                print(f"Skipping already processed folder: {folder}")
                continue
                
            if process_excel_files(folder):
                print(f"✅ Successfully processed Excel files in {folder}")
                processed_excel_folders.add(folder)  # Mark as processed
            else:
                print(f"❌ Failed to process Excel files in {folder}")
        
        # Then do uploads for all processed folders
        print("\n📤 Starting uploads for all processed folders...")
        for folder in processed_excel_folders:  # Use processed_excel_folders instead of completed_folders
            try:
                # Upload directly without re-processing
                success, message = upload_data(folder)
                if success:
                    print(f"✅ Successfully uploaded {folder}")
                else:
                    print(f"❌ Failed to upload {folder}: {message}")
            except Exception as e:
                print(f"❌ Error uploading {folder}: {str(e)}")
        
        print("\n🎉 All processing complete!")
        
    except Exception as e:
        print(f"❌ Error in batch processing: {str(e)}")

def start_next_script():
    """Start the next pending script that hasn't been run yet"""
    try:
        # Get list of pending scripts that haven't been started or completed
        pending_scripts = [
            script for script, status in script_statuses.items()
            if status == ScriptStatus.PENDING 
            and script not in auto_started_scripts
            and script not in processed_scripts
        ]
        
        # Count currently running scripts
        running_count = len([
            script for script, status in script_statuses.items()
            if status == ScriptStatus.RUNNING
        ])
        
        # Start next script if we have pending scripts and room to run more
        if pending_scripts and running_count < max_concurrent_scripts and not terminate_flag.is_set():
            next_script = pending_scripts[0]
            script_statuses[next_script] = ScriptStatus.RUNNING  # Mark as running before starting
            auto_started_scripts.add(next_script)  # Track that we've started this script
            thread = threading.Thread(target=run_script, args=(next_script,))
            thread.daemon = True
            thread.start()
            return True
            
        # If no more pending scripts to start, check if we should start processing
        elif not pending_scripts and all_scripts_completed():
            print("\n✅ All scripts have been run. Starting batch processing...")
            check_and_start_processing()
            
        return False
    except Exception as e:
        print(f"Error starting next script: {e}")
        return False

# Function to close CMD window
def close_cmd_window(hwnd):
    try:
        ctypes.windll.user32.PostMessageA(hwnd, 0x0112, 0xF060, 0)
    except Exception as e:
        print(f"Error closing CMD window: {e}")


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Run multiple scraper scripts with specified days parameter"
    )
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Number of days to look back for bids (default: 2)",
    )
    return parser.parse_args()


def remove_resources(script_name):
    """
    Remove Excel file and bid folders related to the completed script.
    """
    excel_file = f"{script_name.split('.')[0].split('/')[-1]}.xlsx"
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        solicitation_numbers = df["Solicitation Number"].tolist()

        # Remove bid folders
        for folder in os.listdir():
            if os.path.isdir(folder) and folder in solicitation_numbers:
                shutil.rmtree(folder)
                print(f"Removed folder: {folder}")

        # Remove Excel file
        os.remove(excel_file)
        print(f"Removed Excel file: {excel_file}")
    else:
        print(f"Excel file not found: {excel_file}")


def process_excel_files(completed_folder_path: str) -> bool:
    """Process Excel files with API data before upload"""
    try:
        print(f"\nProcessing Excel files in {completed_folder_path}")
        
        # Initialize processor
        processor = ExcelProcessor()
        
        # Fetch API data first
        print("\n📥 Fetching API data...")
        processor.api_categories = processor.fetch_api_data("category")
        processor.api_notice_types = processor.fetch_api_data("notice")
        processor.api_agencies = processor.fetch_api_data("agency")
        processor.api_states = processor.fetch_api_data("state", {"country_id": 10})
        
        if not all([processor.api_categories, processor.api_notice_types, 
                   processor.api_agencies, processor.api_states]):
            print("❌ Failed to fetch API data")
            return False
            
        print("✅ API data fetched successfully")

        # Find all Excel files
        excel_files = glob.glob(os.path.join(completed_folder_path, "*.xlsx"))
        if not excel_files:
            print("❌ No Excel files found to process")
            return False

        print(f"📊 Found {len(excel_files)} Excel files to process")
        
        # Process each Excel file
        success = True
        for file_idx, excel_file in enumerate(excel_files, 1):
            try:
                print(f"\nProcessing file {file_idx}/{len(excel_files)}: {os.path.basename(excel_file)}")
                
                # Load Excel file
                df = pd.read_excel(excel_file)
                total_rows = len(df)
                print(f"📊 Total rows in file: {total_rows}")

                # Handle different title column names
                title_column = 'Solicitation Title' if 'Solicitation Title' in df.columns else 'Title'
                if title_column not in df.columns:
                    print("❌ No Title or Solicitation Title column found")
                    continue

                # Add API columns next to related columns
                if 'Category' in df.columns:
                    category_pos = df.columns.get_loc('Category') + 1
                    df.insert(category_pos, 'API_Category', None)
                    df.insert(category_pos + 1, 'API_Category_ID', None)
                else:
                    df['API_Category'] = None
                    df['API_Category_ID'] = None

                if 'Notice Type' in df.columns:
                    notice_pos = df.columns.get_loc('Notice Type') + 1
                    df.insert(notice_pos, 'API_Notice_Type', None)
                else:
                    df['API_Notice_Type'] = None

                if 'Agency' in df.columns:
                    agency_pos = df.columns.get_loc('Agency') + 1
                    df.insert(agency_pos, 'API_Agency', None)
                else:
                    df['API_Agency'] = None

                if 'State' in df.columns:
                    state_pos = df.columns.get_loc('State') + 1
                    df.insert(state_pos, 'API_State', None)
                else:
                    df['API_State'] = None

                print("\nProcessing rows:")
                # Process each row
                for index, row in df.iterrows():
                    try:
                        progress = int(((index + 1) / total_rows) * 100)
                        print(f"\rProcessing file {file_idx}/{len(excel_files)} - {os.path.basename(excel_file)} - Row {index + 1}/{total_rows} ({progress}%)", end='', flush=True)

                        # Get fields for matching
                        title = str(row.get(title_column, ''))
                        description = str(row.get('Description', ''))
                        original_category = str(row.get('Category', ''))
                        agency_name = str(row.get('Agency', ''))
                        bid_url = str(row.get('Bid Detail Page URL', ''))

                        print(f"\n\nProcessing Row {index + 1}:")
                        print(f"Original Category: {original_category}")
                        print(f"Original Agency: {agency_name}")
                        print(f"Original Notice Type from title: {title}")
                        print(f"Original State info from URL: {bid_url}")

                        # Match category using improved matching
                        category_match = None
                        category_id = None
                        
                        # Split original category if multiple categories are present
                        categories = [cat.strip() for cat in original_category.split(';') if cat.strip()]
                        
                        for cat in categories:
                            # Try exact match first
                            for api_cat in processor.api_categories:
                                if cat.lower() in api_cat['name'].lower() or api_cat['name'].lower() in cat.lower():
                                    category_match = api_cat['name']
                                    category_id = api_cat['id']
                                    break
                            
                            # If no exact match, try fuzzy matching
                            if not category_match:
                                # Use title and description for context
                                combined_text = f"{cat} {title} {description}"
                                for api_cat in processor.api_categories:
                                    if any(word.lower() in combined_text.lower() for word in api_cat['name'].split()):
                                        category_match = api_cat['name']
                                        category_id = api_cat['id']
                                        break

                            if category_match:
                                break

                        if category_match:
                            df.loc[index, 'API_Category'] = category_match
                            df.loc[index, 'API_Category_ID'] = category_id
                            print(f"✓ Matched Category: {category_match} (ID: {category_id})")
                        else:
                            print("✗ No category match found")

                        # Match notice type with improved matching
                        notice_match = None
                        combined_text = f"{title} {description}"
                        
                        # Define common notice type keywords
                        notice_types = {
                            'Request For Proposal': ['rfp', 'request for proposal', 'proposal'],
                            'Invitation For Bid': ['ifb', 'invitation for bid', 'invitation to bid'],
                            'Request For Quote': ['rfq', 'request for quote', 'quotation']
                        }
                        
                        for api_notice in processor.api_notice_types:
                            notice_name = api_notice['name']
                            keywords = notice_types.get(notice_name, [notice_name.lower()])
                            
                            if any(keyword in combined_text.lower() for keyword in keywords):
                                notice_match = notice_name
                                break

                        if notice_match:
                            df.loc[index, 'API_Notice_Type'] = notice_match
                            print(f"✓ Matched Notice Type: {notice_match}")
                        else:
                            print("✗ No notice type match found")

                        # Match agency with improved matching
                        agency_match = None
                        
                        # Clean agency name
                        clean_agency = agency_name.strip().lower()
                        
                        for api_agency in processor.api_agencies:
                            api_agency_name = api_agency['name'].lower()
                            
                            # Try exact match first
                            if clean_agency == api_agency_name:
                                agency_match = api_agency['name']
                                break
                                
                            # Try partial match
                            if clean_agency in api_agency_name or api_agency_name in clean_agency:
                                agency_match = api_agency['name']
                                break
                                
                            # Try matching from URL
                            if bid_url and any(word in bid_url.lower() for word in api_agency_name.split()):
                                agency_match = api_agency['name']
                                break

                        if agency_match:
                            df.loc[index, 'API_Agency'] = agency_match
                            print(f"✓ Matched Agency: {agency_match}")
                        else:
                            print("✗ No agency match found")

                        # Match state with improved matching
                        state_match = None
                        
                        # Extract state from URL or agency name
                        for api_state in processor.api_states:
                            state_name = api_state['name'].lower()
                            state_code = api_state['code'].lower()
                            
                            # Check URL
                            if bid_url and (state_name in bid_url.lower() or state_code in bid_url.lower()):
                                state_match = api_state['name']
                                break
                                
                            # Check agency name
                            if agency_name and (state_name in agency_name.lower() or state_code in agency_name.lower()):
                                state_match = api_state['name']
                                break
                                
                            # Check description
                            if description and (state_name in description.lower() or state_code in description.lower()):
                                state_match = api_state['name']
                                break

                        if state_match:
                            df.loc[index, 'API_State'] = state_match
                            print(f"✓ Matched State: {state_match}")
                        else:
                            print("✗ No state match found")

                        print("-" * 50)  # Separator between rows

                    except Exception as e:
                        print(f"\n❌ Error processing row {index + 1}: {str(e)}")
                        continue

                print("\n")  # New line after progress bar
                
                # Verify data was assigned before saving
                print("\nVerifying data assignment:")
                for col in ['API_Category', 'API_Category_ID', 'API_Notice_Type', 'API_Agency', 'API_State']:
                    filled_count = df[col].notna().sum()
                    print(f"{col}: {filled_count}/{total_rows} rows filled")
                    if filled_count > 0:
                        print("Sample values:")
                        sample_values = df[df[col].notna()][col].head()
                        for i, val in enumerate(sample_values, 1):
                            print(f"  {i}. {val}")

                # Save processed file
                try:
                    df.to_excel(excel_file, index=False, engine='openpyxl')
                    print(f"\n✅ Saved processed file: {os.path.basename(excel_file)}")
                    
                    # Print final DataFrame info
                    print("\nFinal DataFrame columns and non-null counts:")
                    print(df.info())
                    
                except Exception as e:
                    print(f"❌ Error saving file: {str(e)}")
                    success = False
                    continue

                # Verify save was successful
                try:
                    verification_df = pd.read_excel(excel_file)
                    api_columns_present = all(col in verification_df.columns for col in ['API_Category', 'API_Category_ID', 'API_Notice_Type', 'API_Agency', 'API_State'])
                    data_present = any(verification_df[col].notna().any() for col in ['API_Category', 'API_Category_ID', 'API_Notice_Type', 'API_Agency', 'API_State'])
                    
                    if api_columns_present and data_present:
                        print("✅ Verified API columns and data were saved successfully")
                    else:
                        print("⚠️ Warning: Some API columns or data may not have been saved properly")
                        success = False
                except Exception as e:
                    print(f"❌ Error verifying save: {str(e)}")
                    success = False

            except Exception as e:
                print(f"\n❌ Error processing file {os.path.basename(excel_file)}: {str(e)}")
                success = False

        return success

    except Exception as e:
        print(f"\n❌ Error during Excel processing: {str(e)}")
        return False


def upload_data(folder_path: str) -> Tuple[bool, str]:
    """Upload data from a folder without re-processing Excel files"""
    try:
        # Check if folder exists
        if not os.path.exists(folder_path):
            return False, f"Folder not found: {folder_path}"
            
        # Upload the folder contents directly
        upload_cmd = [PYTHON_PATH, UPLOAD_SCRIPT, folder_path]
        result = subprocess.run(upload_cmd, check=True)
        
        if result.returncode == 0:
            return True, "Upload successful"
        else:
            return False, f"Upload failed with return code {result.returncode}"
            
    except Exception as e:
        return False, f"Upload error: {str(e)}"


def remove_empty_folders(folder_path):
    """Remove empty folders within the given folder path"""
    print(f"\nChecking for empty folders in {folder_path}...")
    removed = False

    for root, dirs, files in os.walk(folder_path, topdown=False):
        for dir_name in dirs:
            dir_path = os.path.join(root, dir_name)
            try:
                # Check if directory is empty
                if not os.listdir(dir_path):
                    shutil.rmtree(dir_path)
                    print(f"️ Removed empty folder: {dir_path}")
                    removed = True
            except Exception as e:
                print(f"⚠️ Error removing folder {dir_path}: {e}")

    if not removed:
        print("No empty folders found.")
    return removed


# Add after other global variables
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("scraper_log.txt"), logging.StreamHandler()],
)

console = Console()


# Add this new class for managing script progress
class ScriptProgress:
    def __init__(self):
        self.progress = Progress(
            SpinnerColumn(),
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TimeElapsedColumn(),
        )
        self.tasks = {}

    def add_script(self, script_name):
        """Add a new script to track"""
        task_id = self.progress.add_task(
            f"[cyan]{os.path.basename(script_name)}", total=100
        )
        self.tasks[script_name] = task_id

    def update_script(self, script_name, message, advance=None):
        """Update script progress"""
        if script_name in self.tasks:
            task_id = self.tasks[script_name]
            self.progress.update(
                task_id,
                description=f"[cyan]{os.path.basename(script_name)}: [white]{message}",
                advance=advance,
            )


# Add near the top with other global variables
class ProcessingStats:
    def __init__(self):
        self.total_scripts = len(scripts)
        self.completed_scripts = 0
        self.excel_processed = 0
        self.uploads_completed = 0
        
    def update_progress(self):
        """Calculate and return overall progress"""
        script_progress = (self.completed_scripts / self.total_scripts) * 60  # 60% weight
        excel_progress = (self.excel_processed / max(1, len(glob.glob(os.path.join(yesterday, "*_COMPLETED"))))) * 20  # 20% weight
        upload_progress = (self.uploads_completed / max(1, len(glob.glob(os.path.join(yesterday, "*_COMPLETED"))))) * 20  # 20% weight
        return int(script_progress + excel_progress + upload_progress)

    def log_progress(self):
        """Log current progress"""
        progress = self.update_progress()
        completed = self.completed_scripts
        total = self.total_scripts
        print(f"\n=== Processing Progress: {progress}% ===")
        print(f"Scripts: {completed}/{total} ({int((completed/total)*100)}%)")
        print(f"Excel Processing: {self.excel_processed} folders")
        print(f"Uploads Completed: {self.uploads_completed} folders")
        print("=" * 40)

# Initialize stats tracker
processing_stats = ProcessingStats()

# Modify the run_script function to use rich for output
def run_script(script_name):
    if terminate_flag.is_set():
        return

    # Skip if script has already been run
    if script_name in processed_scripts:
        print(f"Skipping {script_name} - already completed")
        return
        
    if script_name in auto_started_scripts and script_statuses[script_name] != ScriptStatus.RUNNING:
        print(f"Skipping {script_name} - already started")
        return

    with script_semaphore:
        start_time = datetime.now()
        script_progress.add_script(script_name)

        # Setup logging
        log_file = setup_script_logging(script_name, "IN_PROGRESS")

        with print_lock:
            script_statuses[script_name] = ScriptStatus.RUNNING
            console.print(f"\n[bold cyan]Starting {script_name} at {start_time}[/bold cyan]")

        try:
            window_title = f"Scraper - {os.path.basename(script_name)}"

            # Add environment variables to potentially help with webdriver startup
            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1"
            env["PATH"] = env["PATH"] + os.pathsep + os.path.dirname(PYTHON_PATH)

            # Enhanced PowerShell command with better output formatting
            powershell_command = (
                # Setup console encoding
                "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
                "[Console]::InputEncoding = [System.Text.Encoding]::UTF8; "
                "$OutputEncoding = [System.Text.Encoding]::UTF8; "
                f'$host.ui.RawUI.WindowTitle = "{window_title}"; '
                # Setup transcript logging
                f'Start-Transcript -Path "{log_file}" -Append; '
                # Initial status display with colors
                'Write-Host "`n=== Scraper Status ===" -ForegroundColor Cyan; '
                'Write-Host "Script: '
                + os.path.basename(script_name)
                + '" -ForegroundColor Yellow; '
                'Write-Host "Status: Running" -ForegroundColor Green; '
                'Write-Host "Start Time: '
                + start_time.strftime("%H:%M:%S")
                + '" -ForegroundColor Yellow; '
                'Write-Host "Log File: ' + log_file + '" -ForegroundColor Cyan; '
                'Write-Host "==============================`n"; '
                # Enhanced output formatting function
                "function Format-ScraperOutput { "
                "    param($Message) "
                '    $timestamp = Get-Date -Format "HH:mm:ss"; '
                '    if ($Message -match "^ERROR|FAILED|CRITICAL") { '
                '        Write-Host "[$timestamp] " -NoNewline; '
                "        Write-Host $Message -ForegroundColor Red; "
                "    } "
                '    elseif ($Message -match "^WARNING") { '
                '        Write-Host "[$timestamp] " -NoNewline; '
                "        Write-Host $Message -ForegroundColor Yellow; "
                "    } "
                '    elseif ($Message -match "^SUCCESS|COMPLETED|✓") { '
                '        Write-Host "[$timestamp] " -NoNewline; '
                "        Write-Host $Message -ForegroundColor Green; "
                "    } "
                '    elseif ($Message -match "^INFO|PROGRESS") { '
                '        Write-Host "[$timestamp] " -NoNewline; '
                "        Write-Host $Message -ForegroundColor Cyan; "
                "    } "
                "    else { "
                '        Write-Host "[$timestamp] $Message"; '
                "    } "
                "}; "
                # Run the script with enhanced output
                f'& "{PYTHON_PATH}" "{script_name}" --days {args.days} 2>&1 | '
                "ForEach-Object { "
                "    Format-ScraperOutput $_; "
                f'    $host.ui.RawUI.WindowTitle = "{window_title} - Running..."; '
                "}; "
                # Handle script completion and update log file name
                "$exitCode = $LASTEXITCODE; "
                "Stop-Transcript; "
                "if ($exitCode -eq 0) { "
                f'    Rename-Item -Path "{log_file}" -NewName "{os.path.splitext(log_file)[0].replace("IN_PROGRESS", "COMPLETED")}.log" -Force; '
                "} else { "
                f'    Rename-Item -Path "{log_file}" -NewName "{os.path.splitext(log_file)[0].replace("IN_PROGRESS", "FAILED")}.log" -Force; '
                "} "
                # Final status display
                'Write-Host "`n=== Final Status ===" -ForegroundColor Cyan; '
                'Write-Host "End Time: " -NoNewline; '
                'Write-Host $(Get-Date -Format "HH:mm:ss") -ForegroundColor Yellow; '
                'Write-Host "Duration: " -NoNewline; '
                'Write-Host $([math]::Round(((Get-Date) - (Get-Date "'
                + start_time.strftime("%H:%M:%S")
                + '")).TotalMinutes, 2)) -NoNewline -ForegroundColor Yellow; '
                'Write-Host " minutes"; '
                "if ($exitCode -eq 0) { "
                '    Write-Host "Status: " -NoNewline; '
                '    Write-Host "Completed Successfully" -ForegroundColor Green; '
                f'    $host.ui.RawUI.WindowTitle = "{window_title} - Completed ✓"; '
                "} else { "
                '    Write-Host "Status: " -NoNewline; '
                '    Write-Host "Failed" -ForegroundColor Red; '
                f'    $host.ui.RawUI.WindowTitle = "{window_title} - Failed ✗"; '
                '    Write-Host "Exit Code: $exitCode" -ForegroundColor Red; '
                "} "
                'Write-Host "==============================" -ForegroundColor Cyan; '
                "exit $exitCode; "
            )

            # Start PowerShell with a visible window
            process = subprocess.Popen(
                ["powershell", "-NoExit", "-Command", powershell_command],
                creationflags=subprocess.CREATE_NEW_CONSOLE,
                env=env,
            )
            running_processes[script_name] = process

            # Wait for process to complete or detect if window was closed
            while True:
                try:
                    return_code = process.wait(timeout=1)  # Check every second
                    break
                except subprocess.TimeoutExpired:
                    # Check if process is still running
                    if process.poll() is not None:  # Process ended
                        return_code = process.returncode
                        break
                    continue

            end_time = datetime.now()
            execution_time = end_time - start_time

            with print_lock:
                if return_code == 0 or process.poll() is not None:  # Success or window closed
                    print(f"\n{'='*50}")
                    print(f"Script {script_name} completed successfully")
                    print(f"{'='*50}")
                    script_statuses[script_name] = ScriptStatus.SUCCESS
                    processed_scripts.add(script_name)  # Mark as fully completed
                    start_next_script()
                else:
                    print(f"\nScript {script_name} failed with return code {return_code}")
                    script_statuses[script_name] = ScriptStatus.ERROR
                    processed_scripts.add(script_name)  # Mark as completed even if failed
                    start_next_script()

                print_status_report()

        except Exception as e:
            with print_lock:
                console.print(f"[bold red]Error running {script_name}: {str(e)}[/bold red]")
                script_statuses[script_name] = ScriptStatus.ERROR
                processed_scripts.add(script_name)  # Mark as completed on error
                print_status_report()
                start_next_script()

        finally:
            if script_name in running_processes:
                del running_processes[script_name]
            
            # Update completion stats
            if script_statuses[script_name] in [ScriptStatus.RUNNING, ScriptStatus.PENDING]:
                script_statuses[script_name] = ScriptStatus.SUCCESS
                processing_stats.completed_scripts += 1
                processing_stats.log_progress()
            
            # Check if all scripts are either completed or terminated
            remaining_scripts = [
                script for script, status in script_statuses.items()
                if status not in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]
                and script not in auto_started_scripts
            ]
            
            if not remaining_scripts:  # No more scripts to run
                print("\n✅ All scripts have completed or been terminated")
                print("Starting batch processing...")
                check_and_start_processing()
            elif not terminate_flag.is_set():  # Still have scripts to run
                start_next_script()


def close_script_window(window_title):
    """Close the CMD window for a completed script"""
    try:

        def enum_windows_callback(hwnd, _):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                buff = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
                # Make the window title match more flexible
                if window_title.lower() in buff.value.lower():
                    close_cmd_window(hwnd)
                    return False  # Stop enumeration after finding the window
            return True

        # Give the window a moment to finish any final output
        time.sleep(1)

        ctypes.windll.user32.EnumWindows(
            ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)(
                enum_windows_callback
            ),
            0,
        )
    except Exception as e:
        print(f"Error closing window {window_title}: {e}")


def terminate_scripts():
    """Stop all scripts and start processing"""
    try:
        print("Stopping all scripts...")
        
        # Set terminate flag to prevent new scripts
        terminate_flag.set()
        
        # Stop all running processes
        for script_name, process in list(running_processes.items()):
            print(f"Stopping script: {script_name}")
            try:
                process.terminate()
                process.wait(timeout=5)
            except Exception as e:
                print(f"Error stopping {script_name}: {e}")
            
            if script_name in running_processes:
                del running_processes[script_name]
                
        print("All scripts stopped")
        
        # Start batch processing after all scripts are stopped
        print("\n🔄 Starting batch processing of all completed folders...")
        check_and_start_processing()
        
    except Exception as e:
        print(f"Error stopping scripts: {str(e)}")


def ctrl_d_handler():
    print("\nCtrl+D detected. Terminating all scripts...")
    terminate_scripts()


def print_final_status():
    print("\nFinal Script Execution Status:")
    print("-" * 50)
    for script, status in script_statuses.items():
        script_name = os.path.basename(script)
        # Show both SUCCESS and ERROR as "Done"
        display_status = status.value
        if status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]:
            display_status = "Done"

        status_color = (
            "\033[92m"
            if status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]
            else "\033[93m" if status == ScriptStatus.RUNNING else "\033[97m"
        )  # White for pending
        print(f"{script_name}: {status_color}{display_status}\033[0m")
    print("-" * 50)


def cleanup_windows():
    """Clean up any remaining CMD windows on script exit"""
    try:

        def enum_windows_callback(hwnd, _):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                buff = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
                # Close any window with "Scraper -" in the title
                if "Scraper -" in buff.value:
                    close_cmd_window(hwnd)
            return True

        ctypes.windll.user32.EnumWindows(
            ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)(
                enum_windows_callback
            ),
            0,
        )
    except Exception as e:
        print(f"Error during cleanup: {e}")


def signal_handler(signum, frame):
    print("\nCtrl+C detected. Cleaning up...")
    terminate_scripts()
    sys.exit(0)


def print_status_report():
    """Print current status of all scripts"""
    print("\nCurrent Script Status:")
    print("-" * 50)
    for script, status in script_statuses.items():
        script_name = os.path.basename(script)
        # Modify status display - show both SUCCESS and ERROR as "Done"
        display_status = status.value
        if status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]:
            display_status = "Done"

        status_color = (
            "\033[92m"
            if status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]
            else "\033[93m" if status == ScriptStatus.RUNNING else "\033[97m"
        )  # White for pending
        print(f"{script_name:<35}: {status_color}{display_status}\033[0m")
    print("-" * 50)
    sys.stdout.flush()


# Add this new function to periodically update the status display
def update_status_display():
    while not terminate_flag.is_set():
        with print_lock:
            print("\033[2J\033[H")  # Clear screen and move cursor to top
            print("🤖 Bid Scraper Status Dashboard")
            print("=" * 60)

            # Show running scripts with progress
            print("\n Currently Running:")
            running = [
                s
                for s, status in script_statuses.items()
                if status == ScriptStatus.RUNNING
            ]
            if running:
                for script in running:
                    print(f"▶️  {os.path.basename(script)}")
            else:
                print("None")

            # Show pending scripts
            print("\n⏳ Pending:")
            pending = [
                s
                for s, status in script_statuses.items()
                if status == ScriptStatus.PENDING
            ]
            if pending:
                for script in pending:
                    print(f"⌛ {os.path.basename(script)}")
            else:
                print("None")

            # Show completed scripts
            print("\n✅ Completed:")
            completed = [
                s
                for s, status in script_statuses.items()
                if status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]
            ]
            if completed:
                for script in completed:
                    status = script_statuses[script]
                    icon = "✅" if status == ScriptStatus.SUCCESS else "❌"
                    print(f"{icon} {os.path.basename(script)}")
            else:
                print("None")

            print("\n" + "=" * 60)
            sys.stdout.flush()

        time.sleep(5)  # Check every 5 seconds


def upload_pending_folders():
    """Upload any existing COMPLETED folders at startup"""
    print("\nChecking for pending COMPLETED folders...")

    # Look for date folders (YYYY-MM-DD format)
    date_folders = glob.glob("[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]")

    for date_folder in date_folders:
        # Look for *_COMPLETED folders in each date folder
        completed_folders = glob.glob(os.path.join(date_folder, "*_COMPLETED"))

        for completed_folder in completed_folders:
            print(f"\nFound pending folder: {completed_folder}")
            try:
                # Remove empty folders before upload
                remove_empty_folders(completed_folder)

                # Run upload_bids.py with the specified folder
                cmd = [PYTHON_PATH, UPLOAD_SCRIPT, completed_folder]
                print(f"Running command: {' '.join(cmd)}")

                process = subprocess.run(cmd, check=False, shell=True)

                if process.returncode == 0:
                    print(f"✅ Successfully uploaded {completed_folder}")
                    # Remove the folder after successful upload
                    shutil.rmtree(completed_folder)
                    print(f"🗑️ Removed {completed_folder} after upload")
                else:
                    print(
                        f"❌ Upload failed for {completed_folder} with return code: {process.returncode}"
                    )

            except Exception as e:
                print(
                    f"❌ Unexpected error during upload of {completed_folder}: {str(e)}"
                )


def setup_script_logging(script_name: str, status: str) -> str:
    """Setup logging for individual scripts with color coding"""
    # Create logs directory in script folder if it doesn't exist
    script_dir = os.path.dirname(script_name)
    logs_dir = os.path.join(script_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)

    # Create log filename based on script name and status
    script_base = os.path.splitext(os.path.basename(script_name))[0]
    log_file = os.path.join(logs_dir, f"{script_base}_{status}.log")

    return log_file


def display_final_summary():
    """Display a rich table with final execution summary"""
    table = Table(title="Script Execution Summary")
    table.add_column("Script Name", style="cyan")
    table.add_column("Status", style="green")
    table.add_column("Duration", justify="right")
    table.add_column("Log File", style="blue")

    for script_name, status in script_statuses.items():
        script_base = os.path.basename(script_name)
        log_status = "COMPLETED" if status == ScriptStatus.SUCCESS else "FAILED"
        log_file = os.path.join(
            os.path.dirname(script_name),
            "logs",
            f"{os.path.splitext(script_base)[0]}_{log_status}.log",
        )

        status_style = "green" if status == ScriptStatus.SUCCESS else "red"
        duration = "N/A"  # You could track actual duration if needed

        table.add_row(
            script_base,
            f"[{status_style}]{status.value}[/{status_style}]",
            duration,
            os.path.basename(log_file) if os.path.exists(log_file) else "N/A",
        )

    console.print("\n")
    console.print(table)


# Add this near the top with other global variables
yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

# Add script_progress as a global variable
script_progress = None


def main():
    try:
        # Register cleanup function to run on exit
        atexit.register(cleanup_windows)

        # Set up signal handler for Ctrl+C
        signal.signal(signal.SIGINT, signal_handler)

        # Get command line arguments
        global args, script_progress
        args = parse_arguments()

        # Change to the script's directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)
        print(f"Changed working directory to: {script_dir}")

        # Check if Python is available
        if not os.path.exists(PYTHON_PATH):
            print(f"Error: Python interpreter not found at {PYTHON_PATH}")
            print(
                "Please update the PYTHON_PATH variable in the script with the correct path to your Python interpreter."
            )
            return

        # Initialize script progress
        script_progress = ScriptProgress()

        with Live(script_progress.progress, refresh_per_second=10):
            # Upload any pending COMPLETED folders before starting new scrapes
            upload_pending_folders()

            # Start only the status display thread
            status_thread = threading.Thread(target=update_status_display)
            status_thread.daemon = True
            status_thread.start()

            # Process scripts in batches of 4
            while not script_queue.empty() and not terminate_flag.is_set():
                # Start batch of 4 scripts
                batch_threads = []
                for _ in range(min(4, script_queue.qsize())):
                    if not terminate_flag.is_set():
                        script = script_queue.get()
                        thread = threading.Thread(target=run_script, args=(script,))
                        thread.start()
                        batch_threads.append(thread)

                # Wait for all scripts in batch to complete
                for thread in batch_threads:
                    thread.join()

                if terminate_flag.is_set():
                    break

            # Wait for any remaining scripts to finish
            time.sleep(5)

            # Only start processing after ALL scripts are done or terminated
            if terminate_flag.is_set():
                console.print("[yellow]Script execution terminated by user.[/yellow]")
            else:
                console.print("[green]All scripts completed normally[/green]")
                winsound.Beep(2000, 1000)

            # Now start the batch processing
            if all_scripts_completed() or terminate_flag.is_set():
                print("\n🔄 Starting batch processing of all completed folders...")
                check_and_start_processing()

            # Display final summary
            display_final_summary()

    except Exception as e:
        console.print(f"[bold red]Error in main execution: {str(e)}[/bold red]")
        logging.error(f"Error in main execution: {str(e)}")
        logging.error(traceback.format_exc())
    finally:
        cleanup_windows()


if __name__ == "__main__":
    main()