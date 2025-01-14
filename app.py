from flask import Flask, render_template, jsonify, Response, request
from flask_socketio import SocketIO, emit
import os
import sys
import subprocess
from datetime import datetime, timedelta
import json
from pathlib import Path
import threading
import queue
import logging
import pandas as pd
import shutil
import winsound
import argparse
from enum import Enum
import ctypes
import atexit
import glob
from utils.excel_processor import ExcelProcessor
from utils.category_matcher import CategoryMatcher  # Add category matching
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
from rich.text import Text
from rich.console import Group
from dataclasses import dataclass
from typing import Optional, Dict
import time
import signal
import boto3
from werkzeug.utils import secure_filename
import keyboard
import psutil
from utils.utils import play_notification_sound

# Environment Configuration
PYTHON_PATH = r"C:\Users\AliAmani\Miniconda3\envs\bids\python.exe"  # Path to Python executable
CONDA_PATH = r"C:\Users\AliAmani\Miniconda3"  # Path to Miniconda installation
CONDA_ENV = "bids"  # Conda environment name
CONDA_ENV_PATH = r"C:\Users\AliAmani\Miniconda3\envs\bids"  # Path to Conda environment

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Add this near other global variables at the top of the file
running_processes = {}  # Dictionary to keep track of running processes
auto_started_scripts = set()  # Track scripts that have been automatically started

# Global variables
yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

app = Flask(__name__, static_folder="static", template_folder="templates")
app.config["SECRET_KEY"] = "your-secret-key"
app.config["UPLOAD_FOLDER"] = os.path.join(os.getcwd(), yesterday)
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading', logger=True, engineio_logger=True)

# Global variables from master script
UPLOAD_SCRIPT = "upload_bids.py"

# Add script order for automatic chaining
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
    "scrapers/12_Bonfire_FairfaxCounty.py",
    "scrapers/13_eMaryland_eMMA.py",
    "scrapers/14_NorthCarolina_VendorPortal_eVP.py",
    "scrapers/15_State_of_Conneticut_BidBoard.py",
    "scrapers/16_CalProcure.py",
    "scrapers/17_BidNet.py",
    "scrapers/18_Ionwave.py",
    "scrapers/19_Pennsylvania_eMarketplace.py",
    "scrapers/20_County_of_San_Diego.py",
]

def get_next_script(current_script):
    """Get the next script to run based on the script order"""
    try:
        current_index = SCRIPT_ORDER.index(current_script)
        if current_index < len(SCRIPT_ORDER) - 1:
            return SCRIPT_ORDER[current_index + 1]
    except ValueError:
        pass
    return None

class ScriptStatus(Enum):
    PENDING = "Pending"
    RUNNING = "Running"
    SUCCESS = "Done"
    ERROR = "Error"


class ScriptInfo:
    """Class to track script execution information"""

    def __init__(self, script_name):
        self.script_name = script_name
        self.status = ScriptStatus.PENDING
        self.start_time = None
        self.end_time = None
        self.progress = 0
        self.log_file = None
        self.process = None  # Add process tracking
        self.output_buffer = []  # Add output buffer
        self.excel_status = 'Pending'  # Initialize Excel status as Pending
        self.excel_progress = 0   # Add Excel processing progress

    def to_dict(self):
        """Convert to dictionary for JSON serialization"""
        return {
            "script_name": self.script_name,
            "status": self.status.value,
            "progress": self.progress,
            "runtime": (
                str(datetime.now() - self.start_time) if self.start_time else None
            ),
            "log_file": self.log_file,
            "output": "\n".join(self.output_buffer[-100:]),  # Keep last 100 lines
            "excel_status": self.excel_status,  # Include Excel status
            "excel_progress": self.excel_progress  # Include Excel progress
        }


# Initialize script-related variables
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
    "scrapers/12_Bonfire_FairfaxCounty.py",
    "scrapers/13_eMaryland_eMMA.py",
    "scrapers/14_NorthCarolina_VendorPortal_eVP.py",
    "scrapers/15_State_of_Conneticut_BidBoard.py",
    "scrapers/16_CalProcure.py",
    "scrapers/17_BidNet.py",
    "scrapers/18_Ionwave.py",
    "scrapers/19_Pennsylvania_eMarketplace.py",
    "scrapers/20_County_of_San_Diego.py",
]

# Initialize script management
script_queue = queue.Queue()
script_semaphore = threading.Semaphore(8)  # Limit to 8 concurrent scripts
terminate_flag = threading.Event()

# Initialize script infos at startup
script_infos = {}
for script in scripts:
    script_infos[script] = ScriptInfo(script)
    script_queue.put(script)  # Add to queue initially

script_progress = None


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
        self.live = None

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
            # Emit progress update via Socket.IO
            socketio.emit(
                "progress_update",
                {
                    "script": script_name,
                    "progress": self.progress.tasks[task_id].completed,
                    "message": message,
                },
            )

    def start_live_display(self):
        """Start live progress display"""
        self.live = Live(self.progress, refresh_per_second=10)
        self.live.start()

    def stop_live_display(self):
        """Stop live progress display"""
        if self.live:
            self.live.stop()


# Add this near other global variables
active_scripts = set()  # Track currently running scripts

# Add this near other global variables
console = Console()


class ConsoleDisplay:
    def __init__(self):
        self.progress_table = Table(show_header=True, header_style="bold magenta")
        self.progress_table.add_column("Script Name")
        self.progress_table.add_column("Status")
        self.progress_table.add_column("Progress")
        self.progress_table.add_column("Runtime")

        self.summary_table = Table(show_header=True, header_style="bold cyan")
        self.summary_table.add_column("Category", style="cyan")
        self.summary_table.add_column("Count", justify="right")

        self.live = Live(
            Panel(Group(self.progress_table, Text(""), self.summary_table)),  # Spacer
            refresh_per_second=4,
        )

    def update(self):
        """Update the display with current script status"""
        # Clear tables
        self.progress_table.rows.clear()
        self.summary_table.rows.clear()

        # Count statistics
        pending = 0
        running = 0
        completed = 0
        failed = 0

        # Update progress table
        for script_name, info in script_infos.items():
            script_base = os.path.basename(script_name)
            status = info.status.value
            progress = f"{info.progress}%"
            runtime = (
                str(datetime.now() - info.start_time) if info.start_time else "N/A"
            )

            # Update counts
            if status == ScriptStatus.PENDING.value:
                pending += 1
                status_style = "white"
            elif status == ScriptStatus.RUNNING.value:
                running += 1
                status_style = "yellow"
            elif status == ScriptStatus.SUCCESS.value:
                completed += 1
                status_style = "green"
            else:  # ERROR
                failed += 1
                status_style = "red"

            self.progress_table.add_row(
                script_base,
                f"[{status_style}]{status}[/{status_style}]",
                progress,
                runtime,
            )

        # Update summary table
        self.summary_table.add_row("▶️ Running", f"[yellow]{running}[/yellow]")
        self.summary_table.add_row("⏳ Pending", f"[white]{pending}[/white]")
        self.summary_table.add_row("✅ Completed", f"[green]{completed}[/green]")
        self.summary_table.add_row("❌ Failed", f"[red]{failed}[/red]")


# Initialize the console display
console_display = ConsoleDisplay()


# Add this function to set up logging for individual scripts
def setup_script_logging(script_name: str, status: str) -> str:
    """Setup logging for individual scripts with color coding"""
    # Get script base name
    script_base = os.path.splitext(os.path.basename(script_name))[0]

    # Create yesterday's date folder
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    yesterday_folder = os.path.join(os.getcwd(), yesterday)
    os.makedirs(yesterday_folder, exist_ok=True)

    # Create log file path directly in yesterday's folder
    log_file = os.path.join(yesterday_folder, f"{script_base}.log")

    return log_file


# Add this near the top with other global variables
main_log_buffer = []  # Store main terminal logs


def log_to_ui(message):
    """Log message to both terminal and UI"""
    logger.info(message)
    main_log_buffer.append(message)
    socketio.emit("main_log", {"message": message})


def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description="Run web scraping dashboard")
    parser.add_argument(
        "--days", type=int, default=2, help="Number of days to scrape (default: 2)"
    )
    return parser.parse_args()


# Add this before the Flask app initialization
args = parse_arguments()


def process_excel_files(completed_folder_path: str) -> bool:
    """Process Excel files with API data before upload"""
    try:
        # Get script name from folder path - fix path handling
        script_base = os.path.basename(completed_folder_path)
        if script_base.endswith('_COMPLETED'):
            script_name = f"scrapers/{script_base[:-10]}.py"
        else:
            parent_dir = os.path.basename(os.path.dirname(completed_folder_path))
            script_name = f"scrapers/{parent_dir}.py"
            
        if script_name not in script_infos:
            log_to_ui(f"❌ Cannot find script info for {script_name}")
            return False

        log_to_ui("🔄 Starting Excel file processing...")
        
        processor = ExcelProcessor()
        
        # Fetch API data first
        log_to_ui("📥 Fetching API data...")
        processor.api_categories = processor.fetch_api_data("category")
        processor.api_notice_types = processor.fetch_api_data("notice")
        processor.api_agencies = processor.fetch_api_data("agency")
        processor.api_states = processor.fetch_api_data("state", {"country_id": 10})
        
        if not all([processor.api_categories, processor.api_notice_types, 
                   processor.api_agencies, processor.api_states]):
            log_to_ui("❌ Failed to fetch API data")
            return False
            
        log_to_ui("✅ API data fetched successfully")

        # Verify Excel files exist
        excel_files = glob.glob(os.path.join(completed_folder_path, "*.xlsx"))
        if not excel_files:
            log_to_ui("❌ No Excel files found to process")
            return False

        log_to_ui(f"📊 Found {len(excel_files)} Excel files to process")
        total_files = len(excel_files)
        
        # Process each file
        success = True
        for file_idx, excel_file in enumerate(excel_files, 1):
            try:
                file_name = os.path.basename(excel_file)
                log_to_ui(f"📊 Processing {file_name}")

                # Load Excel file
                df = pd.read_excel(excel_file)
                total_rows = len(df)
                log_to_ui(f"📊 Total rows to process: {total_rows}")

                # Initialize API columns if they don't exist
                api_columns = {
                    'API_Category': None,
                    'API_Category_ID': None,
                    'API_Notice_Type': None,
                    'API_Agency': None,
                    'API_State': None
                }
                
                for col in api_columns:
                    if col not in df.columns:
                        # Find appropriate position to insert column
                        if col.startswith('API_Category'):
                            pos = df.columns.get_loc('Category') + 1 if 'Category' in df.columns else len(df.columns)
                        elif col == 'API_Notice_Type':
                            pos = df.columns.get_loc('Notice Type') + 1 if 'Notice Type' in df.columns else len(df.columns)
                        elif col == 'API_Agency':
                            pos = df.columns.get_loc('Agency') + 1 if 'Agency' in df.columns else len(df.columns)
                        elif col == 'API_State':
                            pos = df.columns.get_loc('State') + 1 if 'State' in df.columns else len(df.columns)
                        else:
                            pos = len(df.columns)
                        df.insert(pos, col, None)
                        log_to_ui(f"Added column: {col}")

                # Process each row
                for index, row in df.iterrows():
                    try:
                        # Calculate progress
                        progress = int(((index + 1) / total_rows) * 100)
                        script_infos[script_name].excel_progress = progress
                        script_infos[script_name].progress = progress
                        
                        socketio.emit('script_update', {
                            'script': script_name,
                            'status': script_infos[script_name].status.value,
                            'excel_status': 'Running',
                            'excel_progress': progress,
                            'progress': progress,
                            'message': f'Processing row {index + 1} of {total_rows}'
                        }, namespace='/')

                        # Get fields for matching
                        title = str(row.get('Title', ''))
                        description = str(row.get('Description', ''))
                        original_category = str(row.get('Category', ''))
                        agency_name = str(row.get('Agency', ''))
                        bid_url = str(row.get('Bid Detail Page URL', ''))
                        
                        # 1. Match category
                        category_match, confidence = match_category(title, description, original_category, processor.api_categories)
                        if category_match:
                            df.at[index, 'API_Category'] = category_match.get('category_name')
                            df.at[index, 'API_Category_ID'] = category_match.get('category_id')
                            log_to_ui(f"✅ Category matched: {category_match.get('category_name')} (confidence: {confidence:.2f})")

                        # 2. Match notice type
                        notice_type, _ = processor.determine_notice_type(
                            f"{title} {description}", processor.api_notice_types
                        )
                        df.at[index, 'API_Notice_Type'] = notice_type
                        if notice_type:
                            log_to_ui(f"✅ Notice type matched: {notice_type}")

                        # 3. Match agency
                        agency_name_match, _ = processor.find_best_agency_match(
                            agency_name, bid_url, processor.api_agencies
                        )
                        df.at[index, 'API_Agency'] = agency_name_match
                        if agency_name_match:
                            log_to_ui(f"✅ Agency matched: {agency_name_match}")

                        # 4. Match state
                        state_name, _ = processor.find_state_match(
                            description, agency_name, bid_url, processor.api_states
                        )
                        df.at[index, 'API_State'] = state_name
                        if state_name:
                            log_to_ui(f"✅ State matched: {state_name}")

                    except Exception as e:
                        logger.error(f"Error processing row {index + 1}: {str(e)}")
                        log_to_ui(f"❌ Error processing row {index + 1}: {str(e)}")
                        continue

                # Save back to the same file
                df.to_excel(excel_file, index=False)
                log_to_ui(f"✅ Updated file: {file_name}")

            except Exception as e:
                logger.error(f"Error processing file {file_name}: {str(e)}")
                log_to_ui(f"❌ Error processing file {file_name}: {str(e)}")
                success = False

        return success

    except Exception as e:
        logger.error(f"Error in Excel processing: {str(e)}")
        log_to_ui(f"❌ Error in Excel processing: {str(e)}")
        return False

def match_category(title: str, description: str, category: str, api_categories: list) -> tuple:
    """
    Match a bid to a category using the similarity method
    Returns: (matched_category, confidence_score)
    """
    try:
        # Initialize category matcher
        matcher = CategoryMatcher(api_categories)
        
        # Use similarity method to find match
        match, confidence = matcher.match_by_similarity(title, description, category)
        
        if match:
            return match, confidence
        return None, 0.0
        
    except Exception as e:
        logger.error(f"Error in category matching: {str(e)}")
        return None, 0.0

def run_script(script_name):
    """Run a script and manage its lifecycle"""
    try:
        if script_name in active_scripts:
            log_to_ui(f"Script {script_name} is already running")
            return
            
        active_scripts.add(script_name)
        script_infos[script_name].start_time = datetime.now()
        script_infos[script_name].status = ScriptStatus.RUNNING
        script_infos[script_name].progress = 0
        script_infos[script_name].excel_status = 'Pending'
        script_infos[script_name].excel_progress = 0

        # Set up logging
        log_file = setup_script_logging(script_name, "Running")
        script_infos[script_name].log_file = log_file

        # Create completed folder path
        completed_folder = os.path.join(
            os.path.dirname(script_name),
            f"{os.path.splitext(os.path.basename(script_name))[0]}_COMPLETED",
        )

        # Prepare PowerShell command with output redirection
        powershell_command = f"conda activate bids; & python {script_name} | Tee-Object -FilePath '{log_file}' -Append"
        
        # Set up environment variables for the process
        env = os.environ.copy()
        env["MOZ_HEADLESS"] = "0"  # For Firefox
        env["HEADLESS"] = "False"  # Generic headless flag
        env["PYTHONUNBUFFERED"] = "1"  # Ensure Python output is unbuffered
        env["PYTHONIOENCODING"] = "utf-8"  # Ensure proper encoding

        # Start PowerShell process with visible window
        script_process = subprocess.Popen(
            ["powershell", "-NoExit", "-Command", powershell_command],
            creationflags=subprocess.CREATE_NEW_CONSOLE | subprocess.CREATE_NEW_PROCESS_GROUP,
            env=env,
        )

        # Store process for monitoring
        script_infos[script_name].process = script_process
        running_processes[script_name] = script_process

        # Wait for process to complete
        return_code = script_process.wait()

        # Check for COMPLETED folder or return code
        if os.path.exists(completed_folder) or return_code == 0:
            script_infos[script_name].status = ScriptStatus.SUCCESS
            script_infos[script_name].progress = 100
            log_to_ui(f"Script {script_name} completed successfully")
            
            # Get next script that hasn't been run yet
            next_script = None
            current_index = SCRIPT_ORDER.index(script_name)
            for script in SCRIPT_ORDER[current_index + 1:]:
                if script_infos[script].status == ScriptStatus.PENDING:
                    next_script = script
                    break
            
            if next_script:
                log_to_ui(f"Starting next unrun script: {next_script}")
                thread = threading.Thread(target=run_script, args=(next_script,))
                thread.daemon = True
                thread.start()
            else:
                log_to_ui("No more unrun scripts to start")
                # Check if all scripts are done
                if all(script_infos[s].status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR] for s in SCRIPT_ORDER):
                    log_to_ui("All scripts have completed. Starting Excel processing...")
                    process_all_excel_files()
        else:
            script_infos[script_name].status = ScriptStatus.ERROR
            log_to_ui(f"Script {script_name} failed")

    except Exception as e:
        logger.error(f"Error running script {script_name}: {str(e)}")
        log_to_ui(f"Error running script {script_name}: {str(e)}")
        script_infos[script_name].status = ScriptStatus.ERROR

    finally:
        # Clean up process
        if script_name in active_scripts:
            active_scripts.remove(script_name)
        if script_name in running_processes:
            del running_processes[script_name]
        script_infos[script_name].end_time = datetime.now()
        
        # Kill the process if it's still running
        if script_process and script_process.poll() is None:
            terminate_process(script_process, script_name)

def process_all_excel_files():
    """Process all Excel files after all scripts are complete"""
    try:
        log_to_ui("Starting batch Excel processing...")
        
        # Get all completed folders
        completed_folders = []
        for script in SCRIPT_ORDER:
            completed_folder = os.path.join(
                os.path.dirname(script),
                f"{os.path.splitext(os.path.basename(script))[0]}_COMPLETED",
            )
            if os.path.exists(completed_folder):
                completed_folders.append(completed_folder)
        
        if not completed_folders:
            log_to_ui("No completed folders found for Excel processing")
            return
            
        log_to_ui(f"Found {len(completed_folders)} completed folders to process")
        
        # First process all Excel files
        excel_success = True
        for folder in completed_folders:
            script_name = f"scrapers/{os.path.basename(folder)[:-10]}.py"  # Remove _COMPLETED
            log_to_ui(f"Processing Excel files in {folder}")
            
            # Update script info
            if script_name in script_infos:
                script_infos[script_name].excel_status = 'Running'
                script_infos[script_name].excel_progress = 0
                socketio.emit('script_update', {
                    'script': script_name,
                    'status': script_infos[script_name].status.value,
                    'excel_status': 'Running',
                    'excel_progress': 0
                }, namespace='/')
                
                # Process Excel files
                success = process_excel_files(folder)
                excel_success = excel_success and success
                
                # Update status
                script_infos[script_name].excel_status = 'Done'
                script_infos[script_name].excel_progress = 100
                socketio.emit('script_update', {
                    'script': script_name,
                    'status': script_infos[script_name].status.value,
                    'excel_status': 'Done',
                    'excel_progress': 100
                }, namespace='/')
                
                if not success:
                    log_to_ui(f"❌ Failed to process Excel files in {folder}")
        
        # After ALL Excel files are processed, start upload for each folder
        if excel_success:
            log_to_ui("Excel processing complete. Starting upload process...")
            for folder in completed_folders:
                upload_success, upload_message = upload_data(folder)
                if not upload_success:
                    log_to_ui(f"❌ Upload failed for {folder}: {upload_message}")
                else:
                    log_to_ui(f"✅ Upload completed successfully for {folder}")
        else:
            log_to_ui("❌ Excel processing had errors. Upload process skipped.")
                
    except Exception as e:
        logger.error(f"Error in batch Excel processing: {str(e)}")
        log_to_ui(f"❌ Error in batch Excel processing: {str(e)}")

def start_initial_batch():
    try:
        # Start up to 8 scripts initially
        for _ in range(8):
            # Get next script that hasn't been started
            remaining_scripts = [s for s in SCRIPT_ORDER if s not in auto_started_scripts]
            if not remaining_scripts or terminate_flag.is_set():
                break
                
            next_script = remaining_scripts[0]
            thread = threading.Thread(target=run_script, args=(next_script,))
            thread.daemon = True
            thread.start()
            auto_started_scripts.add(next_script)
            log_to_ui(f"Started initial script: {next_script}")
            time.sleep(1)  # Small delay between starts

        log_to_ui("Initial batch of scripts started")
    except Exception as e:
        log_to_ui(f"Error starting initial scripts: {e}")

def should_start_excel_processing():
    """Check if we should start Excel processing"""
    try:
        # Check if all scripts have been either run or are in error state
        all_scripts_done = all(script in auto_started_scripts for script in SCRIPT_ORDER)
        all_scripts_finished = all(
            script_infos[script].status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]
            for script in SCRIPT_ORDER
        )
        return all_scripts_done and all_scripts_finished
    except Exception as e:
        logger.error(f"Error checking if should start Excel processing: {str(e)}")
        return False

def start_next_script():
    """Start the next script that hasn't been run yet"""
    try:
        # Get remaining scripts that haven't been auto-started
        remaining_scripts = [s for s in SCRIPT_ORDER if s not in auto_started_scripts]
        
        if not remaining_scripts:
            log_to_ui("No more scripts to run - all scripts have been started once")
            # Check if all scripts have finished (success or error)
            if all(script_infos[s].status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR] for s in SCRIPT_ORDER):
                log_to_ui("All scripts have completed. Starting Excel processing and category matching...")
                process_all_excel_files()  # This will handle Excel processing, category matching, and then upload
            return

        # Get the next script
        next_script = remaining_scripts[0]
        
        # Skip if script is already running
        if next_script in active_scripts:
            log_to_ui(f"Script {next_script} is already running")
            return
            
        # Start the script
        thread = threading.Thread(target=run_script, args=(next_script,))
        thread.daemon = True
        thread.start()
        auto_started_scripts.add(next_script)  # Mark as auto-started
        log_to_ui(f"Started next script: {next_script}")
            
    except Exception as e:
        log_to_ui(f"Error starting next script: {str(e)}")

def terminate_process(process, script_name=None):
    """Enhanced process termination"""
    try:
        if process and process.poll() is None:  # If process is still running
            log_to_ui(f"Terminating process {process.pid}")
            kill_process_tree(process.pid)
            process.wait(timeout=5)
            
            if script_name:
                # Clean up process tracking
                if script_name in active_scripts:
                    active_scripts.remove(script_name)
                if script_name in running_processes:
                    del running_processes[script_name]
                
                # Mark as error and track as run
                script_infos[script_name].status = ScriptStatus.ERROR
                auto_started_scripts.add(script_name)
                
                # Start next unrun script if not stopping all
                if not terminate_flag.is_set():
                    # Get remaining scripts that haven't been run
                    remaining_scripts = [s for s in SCRIPT_ORDER if s not in auto_started_scripts]
                    
                    if not remaining_scripts:
                        log_to_ui("No more unrun scripts. Starting Excel processing...")
                        process_all_excel_files()
                    else:
                        next_script = remaining_scripts[0]
                        log_to_ui(f"Starting next unrun script: {next_script}")
                        thread = threading.Thread(target=run_script, args=(next_script,))
                        thread.daemon = True
                        thread.start()
                        auto_started_scripts.add(next_script)

    except Exception as e:
        log_to_ui(f"Error terminating process: {str(e)}")
        if script_name:
            # Clean up even if error occurs
            if script_name in active_scripts:
                active_scripts.remove(script_name)
            if script_name in running_processes:
                del running_processes[script_name]
            # Mark as error and track as run
            script_infos[script_name].status = ScriptStatus.ERROR
            auto_started_scripts.add(script_name)
            # Try to start next script
            if not terminate_flag.is_set():
                start_next_script()

def terminate_scripts():
    """Stop all scripts and exit the application"""
    try:
        log_to_ui("Stopping all scripts...")

        # Set terminate flag to prevent new scripts
        terminate_flag.set()

        # Stop all running processes
        for script_name, process in list(running_processes.items()):
            log_to_ui(f"Stopping script: {script_name}")
            terminate_process(process, script_name)
            if script_name in running_processes:
                del running_processes[script_name]
            if script_name in active_scripts:
                active_scripts.remove(script_name)

        # Check if we should start Excel processing
        if should_start_excel_processing():
            log_to_ui("All scripts have been processed. Starting Excel processing...")
            process_all_excel_files()

        log_to_ui("All scripts stopped")
        play_notification_sound()  # Play sound when all scripts are stopped

    except Exception as e:
        log_to_ui(f"Error stopping scripts: {str(e)}")
        play_notification_sound()  # Play sound for error in stopping scripts


def kill_process_tree(pid):
    """Kill a process and all its children"""
    try:
        parent = psutil.Process(pid)
        children = parent.children(recursive=True)

        # First try to terminate children gracefully
        for child in children:
            try:
                child.terminate()
            except psutil.NoSuchProcess:
                pass

        # Give them some time to terminate
        gone, alive = psutil.wait_procs(children, timeout=3)

        # If still alive, force kill
        for p in alive:
            try:
                p.kill()
            except psutil.NoSuchProcess:
                pass

        # Finally terminate/kill parent
        if parent.is_running():
            parent.terminate()
            parent.wait(3)
            if parent.is_running():
                parent.kill()

    except psutil.NoSuchProcess:
        pass
    except Exception as e:
        log_to_ui(f"Error killing process tree: {e}")


def setup_keyboard_handler():
    def on_ctrl_q(e):
        if e.name == "q" and keyboard.is_pressed("ctrl"):
            log_to_ui("Ctrl+Q detected - stopping all scripts...")
            terminate_scripts()

    keyboard.on_press(on_ctrl_q)


# Add this near the top with other global variables
FLASK_RELOADER_ACTIVE = False  # Track if we've already started scripts


def create_app():
    app = Flask(__name__, static_folder="static", template_folder="templates")
    app.config["SECRET_KEY"] = "your-secret-key"
    app.config["UPLOAD_FOLDER"] = os.path.join(os.getcwd(), yesterday)

    # Initialize script queue at startup
    for script in scripts:
        script_queue.put(script)

    # Start initial batch of scripts
    def start_initial_batch():
        try:
            # Start up to 8 scripts initially
            for _ in range(8):
                # Get next script that hasn't been started
                remaining_scripts = [s for s in SCRIPT_ORDER if s not in auto_started_scripts]
                if not remaining_scripts or terminate_flag.is_set():
                    break
                    
                next_script = remaining_scripts[0]
                thread = threading.Thread(target=run_script, args=(next_script,))
                thread.daemon = True
                thread.start()
                auto_started_scripts.add(next_script)
                log_to_ui(f"Started initial script: {next_script}")
                time.sleep(1)  # Small delay between starts

            log_to_ui("Initial batch of scripts started")
        except Exception as e:
            log_to_ui(f"Error starting initial scripts: {e}")

    # Start scripts after a short delay
    threading.Timer(2.0, start_initial_batch).start()

    # Register routes
    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/api/scripts")
    def get_scripts():
        """Get list of all scripts and their current status"""
        try:
            scripts_info = []
            for script_name, info in script_infos.items():
                # Get log content if available
                log_content = ""
                if info.log_file and os.path.exists(info.log_file):
                    try:
                        with open(info.log_file, "r", encoding="utf-8") as f:
                            log_content = f.read()
                    except Exception as e:
                        pass

                scripts_info.append(
                    {
                        "name": script_name,
                        "status": info.status.value,
                        "progress": info.progress,
                        "excel_status": info.excel_status,
                        "excel_progress": info.excel_progress,
                        "log": log_content,
                        "runtime": (
                            str(datetime.now() - info.start_time)
                            if info.start_time
                            else None
                        ),
                    }
                )

            return jsonify(scripts_info)
        except Exception as e:
            log_to_ui(f"Error getting scripts: {str(e)}")
            return jsonify([])

    @app.route("/api/master/status")
    def get_master_status():
        """Get master script status and output"""
        try:
            running_count = len(
                [s for s in script_infos.values() if s.status == ScriptStatus.RUNNING]
            )
            pending_count = len(
                [s for s in script_infos.values() if s.status == ScriptStatus.PENDING]
            )
            completed_count = len(
                [
                    s
                    for s in script_infos.values()
                    if s.status in [ScriptStatus.SUCCESS, ScriptStatus.ERROR]
                ]
            )

            return jsonify(
                {
                    "running": running_count,
                    "pending": pending_count,
                    "completed": completed_count,
                    "total": len(scripts),
                }
            )
        except Exception as e:
            log_to_ui(f"Error getting master status: {str(e)}")
            return jsonify(
                {"running": 0, "pending": 0, "completed": 0, "total": len(scripts)}
            )

    @app.route("/api/table")
    def get_table():
        """Get the current status table"""
        try:
            table = Table(title="Script Execution Summary")
            table.add_column("Script Name", style="cyan")
            table.add_column("Status", style="green")
            table.add_column("Progress", justify="right")
            table.add_column("Runtime", style="blue")

            for script_name, info in script_infos.items():
                script_base = os.path.basename(script_name)
                status = info.status.value
                progress = info.progress
                runtime = (
                    str(datetime.now() - info.start_time) if info.start_time else "N/A"
                )

                status_style = (
                    "green"
                    if status == "Done"
                    else "red" if status == "Error" else "yellow"
                )

                table.add_row(
                    script_base,
                    f"[{status_style}]{status}[/{status_style}]",
                    f"{progress}%",
                    runtime,
                )

            # Capture the table output
            console = Console(record=True)
            console.print(table)
            return jsonify({"table": console.export_html()})
        except Exception as e:
            log_to_ui(f"Error generating table: {str(e)}")
            return jsonify({"table": ""})

    @app.route("/api/main_log")
    def get_main_log():
        """Get the main terminal log"""
        return jsonify({"log": "\n".join(main_log_buffer)})

    def match_category(title: str, description: str, category: str, api_categories: list) -> tuple:
        """
        Match a bid to a category using the similarity method
        Returns: (matched_category, confidence_score)
        """
        try:
            # Initialize category matcher
            matcher = CategoryMatcher(api_categories)
            
            # Use similarity method to find match
            match, confidence = matcher.match_by_similarity(title, description, category)
            
            if match:
                return match, confidence
            return None, 0.0
            
        except Exception as e:
            logger.error(f"Error in category matching: {str(e)}")
            return None, 0.0

    @app.route('/api/match_category', methods=['POST'])
    def api_match_category():
        """API endpoint for category matching"""
        try:
            data = request.get_json()
            title = data.get('title', '')
            description = data.get('description', '')
            category = data.get('category', '')
            api_categories = data.get('api_categories', [])
            
            if not title or not api_categories:
                return jsonify({
                    'error': 'Missing required fields',
                    'status': 'error'
                }), 400
                
            match, confidence = match_category(title, description, category, api_categories)
            
            if match:
                return jsonify({
                    'status': 'success',
                    'match': match,
                    'confidence': confidence
                })
            
            return jsonify({
                'status': 'no_match',
                'message': 'No suitable category match found'
            })
            
        except Exception as e:
            logger.error(f"Category matching API error: {str(e)}")
            return jsonify({
                'status': 'error',
                'error': str(e)
            }), 500

    @app.route("/api/stop", methods=['POST'])
    def stop_scripts():
        """Stop specific script or all scripts"""
        try:
            data = request.get_json()
            script_name = data.get('script') if data else None
            
            if script_name:
                # Stop specific script
                script_info = script_infos.get(script_name)
                if script_info and script_info.process:
                    log_to_ui(f"Stopping script: {script_name}")
                    terminate_process(script_info.process, script_name)  # Pass script_name
                    script_info.status = ScriptStatus.ERROR
                    script_info.end_time = datetime.now()
                    return jsonify({"status": "success", "message": f"Script {script_name} stopped"})
                return jsonify({"status": "error", "message": "Script not running"}), 404
            else:
                # Stop all scripts
                terminate_flag.set()  # Set terminate flag before stopping all
                terminate_scripts()
                return jsonify({"status": "success", "message": "All scripts stopped"})
                
        except Exception as e:
            log_to_ui(f"Error stopping scripts: {str(e)}")
            return jsonify({"status": "error", "message": str(e)}), 500

    @app.route("/api/start", methods=['POST'])
    def start_scripts():
        """Start specific script or initial batch of scripts"""
        try:
            data = request.get_json()
            script_name = data.get('script') if data else None
            
            if script_name:
                # Start specific script
                if script_name not in script_infos:
                    return jsonify({"status": "error", "message": "Script not found"}), 404
                    
                script_info = script_infos[script_name]
                if script_info.status != ScriptStatus.RUNNING:
                    # Clear the queue and add only scripts from this one onwards
                    with script_queue.mutex:
                        script_queue.queue.clear()
                    start_idx = scripts.index(script_name)
                    for script in scripts[start_idx:]:
                        script_queue.put(script)
                    
                    thread = threading.Thread(target=run_script, args=(script_name,))
                    thread.daemon = True
                    thread.start()
                    return jsonify({"status": "success", "message": f"Script {script_name} started"})
                return jsonify({"status": "error", "message": "Script already running"}), 400
            else:
                # Start initial batch of scripts
                # Clear any existing scripts in queue
                with script_queue.mutex:
                    script_queue.queue.clear()
                # Add all scripts to queue
                for script in scripts:
                    script_queue.put(script)
                
                for _ in range(min(4, len(scripts))):
                    if not terminate_flag.is_set():
                        start_next_script()
                return jsonify({"status": "success", "message": "Initial batch of scripts started"})
                
        except Exception as e:
            log_to_ui(f"Error starting scripts: {str(e)}")
            return jsonify({"status": "error", "message": str(e)}), 500

    @app.route("/api/app/status")
    def get_app_status():
        """Get the status of the app.py process"""
        try:
            # Check if app.py is running
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                if 'python' in proc.info['name'].lower() and 'app.py' in ' '.join(proc.info['cmdline']):
                    return jsonify({"status": "running", "pid": proc.info['pid']})
            return jsonify({"status": "stopped"})
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)})

    @app.route("/api/app/start", methods=['POST'])
    def start_app():
        """Start app.py if not running"""
        try:
            # Check if already running
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                if 'python' in proc.info['name'].lower() and 'app.py' in ' '.join(proc.info['cmdline']):
                    return jsonify({"status": "already_running"})
            
            # Start app.py
            subprocess.Popen([sys.executable, "app.py"])
            return jsonify({"status": "started"})
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)})

    @app.route("/api/app/stop", methods=['POST'])
    def stop_app():
        """Stop app.py if running"""
        try:
            stopped = False
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                if 'python' in proc.info['name'].lower() and 'app.py' in ' '.join(proc.info['cmdline']):
                    # Terminate the process
                    process = psutil.Process(proc.info['pid'])
                    process.terminate()
                    process.wait(timeout=5)
                    stopped = True
                    log_to_ui("App process terminated successfully")
            
            # Stop all running scripts
            terminate_scripts()
            
            return jsonify({
                "status": "stopped" if stopped else "not_running",
                "message": "App and all scripts stopped successfully" if stopped else "App was not running"
            })
        except Exception as e:
            log_to_ui(f"Error stopping app: {str(e)}")
            return jsonify({
                "status": "error",
                "message": str(e)
            }), 500

    @app.route("/api/logs/<path:script_name>")
    def get_script_logs(script_name):
        """Get logs for a specific script"""
        try:
            # Get yesterday's date folder
            yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
            log_folder = os.path.join(os.getcwd(), yesterday)
            
            # Clean up script name to get base name
            script_base = os.path.basename(script_name)
            if script_base.endswith('.py'):
                script_base = script_base[:-3]
            
            log_file = os.path.join(log_folder, f"{script_base}.log")
            
            if not os.path.exists(log_file):
                return jsonify({
                    "status": "not_found",
                    "message": "No logs available yet. Logs will appear here once the script starts running."
                }), 404
            
            # Read as binary first
            with open(log_file, 'rb') as f:
                content = f.read()
            
            # Remove BOM if present
            if content.startswith(b'\xff\xfe'):  # UTF-16 BOM
                content = content[2:].decode('utf-16')
            elif content.startswith(b'\xef\xbb\xbf'):  # UTF-8 BOM
                content = content[3:].decode('utf-8')
            else:
                content = content.decode('utf-8', errors='ignore')
            
            # Clean up any remaining invalid characters at start
            content = content.lstrip('\ufeff\ufffe')
            
            return jsonify({
                "status": "success",
                "content": content,
                "format": "powershell"
            })
            
        except Exception as e:
            logger.error(f"Error accessing logs: {str(e)}")
            return jsonify({
                "status": "error",
                "message": f"Error accessing logs: {str(e)}"
            }), 500

    return app

def remove_empty_folders(folder_path: str) -> bool:
    """Remove empty folders within the given folder path"""
    log_to_ui(f"Checking for empty folders in {folder_path}...")
    removed = False

    for root, dirs, files in os.walk(folder_path, topdown=False):
        for dir_name in dirs:
            dir_path = os.path.join(root, dir_name)
            try:
                # Check if directory is empty
                if not os.listdir(dir_path):
                    shutil.rmtree(dir_path)
                    log_to_ui(f"Removed empty folder: {dir_path}")
                    removed = True
            except Exception as e:
                log_to_ui(f"Error removing folder {dir_path}: {e}")

    if not removed:
        log_to_ui("No empty folders found.")
    return removed

def upload_data(completed_folder_path: str) -> tuple[bool, str]:
    """Upload Excel file and bid attachments to the server"""
    MAX_RETRIES = 3

    try:
        # First verify the folder exists and has content
        if not os.path.exists(completed_folder_path):
            return False, f"Folder not found: {completed_folder_path}"

        # Check if folder has any contents
        contents = os.listdir(completed_folder_path)
        if not contents:
            return False, f"Folder is empty: {completed_folder_path}"

        log_to_ui(f"Found {len(contents)} items in {completed_folder_path}")
        log_to_ui(f"Contents: {', '.join(contents)}")

        # Ensure UPLOAD_SCRIPT path is correct and exists
        upload_script = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "upload_bids.py"
        )
        if not os.path.exists(upload_script):
            return False, f"Upload script not found at: {upload_script}"

        # Set up environment for proper encoding
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONLEGACYWINDOWSSTDIO"] = "utf-8"

        # Try upload with retries
        for attempt in range(MAX_RETRIES):
            try:
                # Run upload_bids.py as a command line script
                cmd = [sys.executable, upload_script, completed_folder_path]
                log_to_ui(
                    f"Running upload command (attempt {attempt + 1}/{MAX_RETRIES}): {' '.join(cmd)}"
                )

                # Run the upload process
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

                result = subprocess.run(
                    cmd,
                    env=env,
                    startupinfo=startupinfo,
                    capture_output=True,
                    text=True,
                    check=True,
                )

                if result.returncode == 0:
                    log_to_ui("✅ Upload completed successfully")
                    
                    # Create upload success flag
                    with open(
                        os.path.join(completed_folder_path, ".upload_complete"), "w"
                    ) as f:
                        f.write(datetime.now().isoformat())
                    
                    # Remove empty folders after successful upload
                    log_to_ui("🧹 Cleaning up empty folders...")
                    remove_empty_folders(completed_folder_path)
                    
                    # Remove the original folder after successful upload and cleanup
                    try:
                        shutil.rmtree(completed_folder_path)
                        log_to_ui(f"✅ Removed original folder: {completed_folder_path}")
                    except Exception as e:
                        log_to_ui(f"⚠️ Error removing original folder: {str(e)}")
                    
                    return True, "Upload completed successfully and resources cleaned up"

            except subprocess.CalledProcessError as e:
                log_to_ui(f"❌ Upload failed (attempt {attempt + 1}): {e.stderr}")
                if attempt < MAX_RETRIES - 1:
                    log_to_ui(f"Retrying in 5 seconds...")
                    time.sleep(5)
                else:
                    return (
                        False,
                        f"Upload failed after {MAX_RETRIES} attempts: {e.stderr}",
                    )

            except Exception as e:
                log_to_ui(f"❌ Unexpected error during upload: {str(e)}")
                logger.error(f"Upload error: {traceback.format_exc()}")
                return False, f"Unexpected error: {str(e)}"

        return False, f"Upload failed after {MAX_RETRIES} attempts"

    except Exception as e:
        error_msg = f"Critical error in upload process: {str(e)}"
        log_to_ui(f"❌ {error_msg}")
        logger.error(f"Upload critical error: {traceback.format_exc()}")
        return False, error_msg

def update_log_file_status(script_name: str, status: str) -> None:
    """Update the log file name with the final status"""
    try:
        if script_infos[script_name].log_file:
            old_path = script_infos[script_name].log_file
            new_path = old_path.replace("IN_PROGRESS", status)
            if os.path.exists(old_path):
                os.rename(old_path, new_path)
                script_infos[script_name].log_file = new_path
    except Exception as e:
        logger.error(f"Error updating log file status: {str(e)}")

if __name__ == "__main__":
    try:
        # Parse command line arguments
        args = parse_arguments()
        log_to_ui(f"Starting scraper dashboard with {args.days} days of data")

        # Create the Flask app
        app = create_app()
        socketio.run(app, debug=False, port=5000)  # Disable debug mode
    except KeyboardInterrupt:
        terminate_scripts()
    except Exception as e:
        log_to_ui(f"Application error: {str(e)}")
        terminate_scripts()
