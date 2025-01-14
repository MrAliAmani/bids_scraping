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

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# Specify the path to your Python interpreter
PYTHON_PATH = r"C:\Users\AliAmani\Miniconda3\envs\bids\python.exe"

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


# Add these after other global variables
class ScriptStatus(Enum):
    PENDING = "Pending"
    RUNNING = "Running"
    SUCCESS = "Done"
    ERROR = "Done"


# Dictionary to track script statuses
script_statuses = {script: ScriptStatus.PENDING for script in scripts}


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
        print(f"\n Processing Excel files in {completed_folder_path}")
        processor = ExcelProcessor()

        if processor.process_completed_folder(completed_folder_path):
            print("✅ Successfully processed all Excel files")
            return True
        else:
            print("❌ Failed to process some Excel files")
            return False

    except Exception as e:
        print(f"❌ Error during Excel processing: {str(e)}")
        return False


def upload_data(completed_folder_path):
    """Upload Excel file and bid attachments to the server"""
    try:
        # First verify the folder exists and has content
        if not os.path.exists(completed_folder_path):
            return False, f"Folder not found: {completed_folder_path}"

        # Check if folder has any contents
        contents = os.listdir(completed_folder_path)
        if not contents:
            return False, f"Folder is empty: {completed_folder_path}"

        print(f"Found {len(contents)} items in {completed_folder_path}")
        print(f"Contents: {contents}")

        # Process Excel files before upload
        if not process_excel_files(completed_folder_path):
            return False, "Failed to process Excel files"

        # Set up environment for proper encoding
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONLEGACYWINDOWSSTDIO"] = "utf-8"

        # Run upload_bids.py with the specified folder
        cmd = [sys.executable, UPLOAD_SCRIPT, completed_folder_path]
        print(f"Running upload command: {' '.join(cmd)}")

        # Run the upload process with full output capture
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        upload_process = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            startupinfo=startupinfo,
            shell=False,
            cwd=os.path.dirname(os.path.abspath(__file__)),
        )

        # Always log the output for debugging
        if upload_process.stdout:
            print("Upload output:")
            print(upload_process.stdout)
        if upload_process.stderr:
            print("Upload errors:")
            print(upload_process.stderr)
        print(f"Upload return code: {upload_process.returncode}")

        if upload_process.returncode == 0:
            # Only remove the folder after successful upload
            try:
                shutil.rmtree(completed_folder_path)
                print(
                    f"🗑️ Removed completed folder after successful upload: {completed_folder_path}"
                )
            except Exception as e:
                print(f"⚠️ Warning: Could not remove folder after upload: {e}")
            return True, "Upload successful"
        else:
            error_msg = (
                upload_process.stderr or upload_process.stdout or "Unknown error"
            )
            return (
                False,
                f"Upload failed (code {upload_process.returncode}): {error_msg}",
            )

    except Exception as e:
        print(f"Exception during upload: {str(e)}")
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


# Modify the run_script function to use rich for output
def run_script(script_name):
    if terminate_flag.is_set():
        return

    with script_semaphore:
        start_time = datetime.now()
        script_progress.add_script(script_name)

        # Setup logging
        log_file = setup_script_logging(script_name, "IN_PROGRESS")

        with print_lock:
            script_statuses[script_name] = ScriptStatus.RUNNING
            console.print(
                f"\n[bold cyan]Starting {script_name} at {start_time}[/bold cyan]"
            )

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

            # Wait for process to complete
            return_code = process.wait()

            end_time = datetime.now()
            execution_time = end_time - start_time

            with print_lock:
                if return_code == 0:
                    print(f"\n{'='*50}")
                    print(f"Script {script_name} completed successfully")
                    print(f"{'='*50}")
                    script_statuses[script_name] = ScriptStatus.SUCCESS

                    try:
                        # Get the script base name without extension
                        script_base_name = script_name.split(".")[0].split("/")[-1]
                        completed_folder = f"{script_base_name}_COMPLETED"
                        date_folder = os.path.join(os.getcwd(), yesterday)
                        completed_folder_path = os.path.join(
                            date_folder, completed_folder
                        )

                        print(
                            f"\n🔍 Looking for completed folder at: {completed_folder_path}"
                        )

                        if os.path.exists(completed_folder_path):
                            print(f"Found completed folder: {completed_folder_path}")

                            # Process Excel files and upload
                            print(f"\n📤 Attempting to upload {completed_folder}")
                            success, message = upload_data(completed_folder_path)

                            if success:
                                print(f"\n✅ Successfully uploaded {completed_folder}")
                                # Folder is already removed in upload_data after successful upload
                            else:
                                print(f"\n❌ Upload failed for {completed_folder}")
                                print(f"Error details: {message}")
                                logging.error(
                                    f"Upload failed for {completed_folder}: {message}"
                                )
                        else:
                            print(
                                f"⚠️ No completed folder found at: {completed_folder_path}"
                            )
                            logging.warning(
                                f"No completed folder found at: {completed_folder_path}"
                            )

                    except Exception as e:
                        print(f"❌ Error during post-completion process: {str(e)}")
                        logging.error(
                            f"Error during post-completion process for {script_name}: {str(e)}"
                        )
                        print(f"Stack trace: {traceback.format_exc()}")
                else:
                    print(
                        f"\nScript {script_name} failed with return code {return_code}"
                    )
                    script_statuses[script_name] = ScriptStatus.ERROR

                print_status_report()

            # Play completion sound
            winsound.Beep(1000, 500)

        except Exception as e:
            with print_lock:
                console.print(
                    f"[bold red]Error running {script_name}: {str(e)}[/bold red]"
                )
                script_statuses[script_name] = ScriptStatus.ERROR
                print_status_report()

        finally:
            if script_name in running_processes:
                del running_processes[script_name]

            # Start next script after this one finishes
            if not terminate_flag.is_set():
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


def start_next_script():
    """Start the next script from the queue"""
    try:
        # Count currently running scripts
        running_count = len(
            [s for s in script_statuses.values() if s == ScriptStatus.RUNNING]
        )

        # Only start new script if we have less than max_concurrent_scripts running
        if (
            running_count < max_concurrent_scripts
            and not script_queue.empty()
            and not terminate_flag.is_set()
        ):
            next_script = script_queue.get()
            thread = threading.Thread(target=run_script, args=(next_script,))
            thread.daemon = True
            thread.start()
            return True
        return False
    except Exception as e:
        print(f"Error starting next script: {e}")
        return False


def terminate_scripts():
    terminate_flag.set()

    # Terminate all running processes
    for script_name, process in list(running_processes.items()):
        print(f"Terminating {script_name}")
        if sys.platform == "win32":
            subprocess.call(["taskkill", "/F", "/T", "/PID", str(process.pid)])
        else:
            os.killpg(os.getpgid(process.pid), signal.SIGTERM)

    # Clear the script queue
    while not script_queue.empty():
        script_queue.get()

    # Clean up any remaining windows
    cleanup_windows()


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

            if terminate_flag.is_set():
                console.print("[yellow]Script execution terminated by user.[/yellow]")
            else:
                console.print("[green]All scripts completed[/green]")
                winsound.Beep(2000, 1000)

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
