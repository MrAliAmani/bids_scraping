import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import functools

# Override the print function to always flush
print = functools.partial(print, flush=True)

import os

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import time
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException,
)
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import shutil
import glob
from selenium.webdriver.common.action_chains import ActionChains
import winsound  # For playing notification sounds
import sys
import argparse
import logging
import re
from utils.utils import safe_move, play_notification_sound
import requests
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue
import json
from pathlib import Path

# Script name
script_name = os.path.splitext(os.path.basename(__file__))[0]

# Set up logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Create a folder named after yesterday's date
yesterday = datetime.now() - timedelta(days=1)
main_folder = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    f"{yesterday.strftime('%Y-%m-%d')}\\{script_name}_IN_PROGRESS",
)
os.makedirs(main_folder, exist_ok=True)

# Create temporary download folder (without suffix)
temp_download_folder = os.path.join(main_folder, script_name)
os.makedirs(temp_download_folder, exist_ok=True)


def play_notification(frequency=2500, duration=1000):
    """Play a notification sound."""
    play_notification_sound()


def error_handler(error_message):
    """Handle errors by playing a notification and pausing the script."""
    logger.error(f"Error: {error_message}")
    play_notification()
    input("Press Enter to continue...")


def setup_driver():
    """Set up and configure the Chrome WebDriver."""
    logger.info("Setting up Chrome WebDriver")
    chrome_options = webdriver.ChromeOptions()

    # Use temp_download_folder for downloads
    prefs = {
        "download.default_directory": temp_download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")  # Disables GPU acceleration
    chrome_options.add_argument("--disable-software-rasterizer")  # Use CPU rendering

    # Add arguments to prevent disconnection
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-features=TranslateUI")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-background-networking")
    chrome_options.add_argument("--disable-background-timer-throttling")
    chrome_options.add_argument("--disable-backgrounding-occluded-windows")
    chrome_options.add_argument("--disable-renderer-backgrounding")
    chrome_options.add_argument("--disable-features=IsolateOrigins,site-per-process")

    # Set page load strategy
    chrome_options.set_capability("pageLoadStrategy", "normal")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # Set page load timeout
    driver.set_page_load_timeout(300)  # 5 minutes timeout

    # Execute JavaScript to prevent screen timeout
    driver.execute_script(
        """
    function preventTimeout() {
        window.setTimeout(preventTimeout, 60000);
        if (document.hasFocus()) {
            var event = new Event('mousemove');
            document.dispatchEvent(event);
        }
    }
    preventTimeout();
    """
    )

    logger.info("Chrome WebDriver set up successfully")
    return driver


def keep_session_alive(driver):
    """Perform actions to keep the session alive."""
    try:
        # Scroll slightly to simulate activity
        driver.execute_script("window.scrollTo(0, window.scrollY + 1)")
        driver.execute_script("window.scrollTo(0, window.scrollY - 1)")

        # Move mouse to prevent system idle
        actions = ActionChains(driver)
        actions.move_by_offset(1, 1).perform()
        actions.move_by_offset(-1, -1).perform()

    except Exception as e:
        logger.warning(f"Error in keep_session_alive: {str(e)}")


def extract_bid_links(driver, start_date):
    """Extract bid links from the main page."""
    logger.info(f"Extracting bid links from {start_date}")
    bid_links = []
    page = 1
    while True:
        logger.info(f"Processing page {page}")
        rows = driver.find_elements(By.XPATH, "//table[@id='body_x_grid_grd']/tbody/tr")
        for row in rows:
            try:
                posted_date = row.find_element(By.XPATH, "./td[6]").text
                posted_date = datetime.strptime(posted_date, "%m/%d/%Y %I:%M:%S %p")
                if posted_date >= start_date:
                    link = row.find_element(By.XPATH, "./td[3]/a").get_attribute("href")
                    response_date = row.find_element(By.XPATH, "./td[5]").text
                    bid_links.append((link, posted_date, response_date))
                    logger.debug(f"Added bid link: {link}")
                else:
                    logger.info(
                        f"Reached bids older than {start_date}, stopping extraction"
                    )
                    return bid_links
            except (NoSuchElementException, ValueError) as e:
                logger.warning(f"Error processing row: {str(e)}")
                continue

        if len(bid_links) >= 50 or page >= 5:
            logger.info(f"Reached limit: {len(bid_links)} links or {page} pages")
            break

        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "body_x_grid_PagerBtnNextPage"))
            )
            if "disabled" in next_button.get_attribute("class"):
                logger.info("No more pages available")
                break
            next_button.click()
            WebDriverWait(driver, 10).until(EC.staleness_of(rows[0]))
            time.sleep(2)
            page += 1
        except (NoSuchElementException, TimeoutException) as e:
            logger.warning(f"Error navigating to next page: {str(e)}")
            break

    logger.info(f"Extracted {len(bid_links)} bid links")
    return bid_links


def extract_bid_details(driver, bid_link):
    """Extract detailed information for a specific bid."""
    driver.get(bid_link)
    time.sleep(2)
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblProcessCode")
        )
    )

    bid_details = {}

    # Updated extraction of Notice Type (Solicitation Type)
    bid_details["Notice Type"] = driver.find_element(
        By.XPATH,
        "//div[@data-selector='body_x_tabc_rfp_ext_prxrfp_ext_x_selRfptypeCode']//div[@class='text']",
    ).text
    bid_details["Solicitation Number"] = driver.find_element(
        By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblProcessCode"
    ).text
    bid_details["Solicitation Title"] = driver.find_element(
        By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblLabel"
    ).text
    bid_details["Agency"] = driver.find_element(
        By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_txtOrgaLabel"
    ).get_attribute("value")
    bid_details["Category"] = driver.find_element(
        By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_txtFamLabel"
    ).get_attribute("value")

    # Extract Description
    try:
        description = driver.find_element(
            By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblSummary"
        ).text
        bid_details["Description"] = description.strip()
    except NoSuchElementException:
        bid_details["Description"] = ""

    try:
        procurement_officer = driver.find_element(
            By.XPATH,
            "//div[@data-selector='body_x_tabc_rfp_ext_prxrfp_ext_x_selContactId_1']//div[@class='text']",
        ).text
        email = driver.find_element(
            By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_txtContactEmail_1"
        ).get_attribute("value")
        bid_details["Contact Information"] = f"{procurement_officer}, Email: {email}"
    except NoSuchElementException:
        bid_details["Contact Information"] = ""

    bid_details["Bid Detail Page URL"] = bid_link

    try:
        additional_summary = driver.find_element(
            By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblAdditionalSummary"
        ).get_attribute("innerHTML")
        bid_details["Additional Summary"] = " ".join(additional_summary.split())
    except NoSuchElementException:
        bid_details["Additional Summary"] = ""

    bid_details["Contracting Office Address"] = ""

    # Update attachment finding
    attachments = driver.find_elements(By.CSS_SELECTOR, "a.iv-download-file")
    attachment_info = []
    for attachment in attachments:
        try:
            name = attachment.find_element(
                By.CSS_SELECTOR, "span[data-iv-role='label']"
            ).text
            attachment_info.append(name)
        except NoSuchElementException:
            continue
    bid_details["Attachments"] = " | ".join(attachment_info)

    return bid_details


def normalize_filename(filename):
    """Normalize filename by removing special characters and standardizing format."""
    if not filename or filename.isspace():
        return ""

    # Remove file extension for processing
    name, ext = os.path.splitext(filename)

    # Replace special characters with underscores
    normalized = re.sub(r"[^\w\s.-]", "_", name)

    # Replace multiple spaces/underscores with single underscore
    normalized = re.sub(r"[\s_]+", "_", normalized)

    # Remove duplicate indicators like (1), (2), etc.
    normalized = re.sub(r"_\(\d+\)", "", normalized)

    # Add extension back
    if ext:
        normalized = f"{normalized}{ext}"

    return normalized.strip("_")


def is_duplicate_file(file_path1, file_path2):
    """Check if two files are duplicates by comparing content."""
    if not (os.path.exists(file_path1) and os.path.exists(file_path2)):
        return False

    try:
        with open(file_path1, "rb") as f1, open(file_path2, "rb") as f2:
            return f1.read() == f2.read()
    except:
        return False


def handle_duplicate_files(bid_folder):
    """Remove duplicate files and standardize filenames."""
    files = os.listdir(bid_folder)
    processed_files = set()
    duplicates = []

    for file in files:
        normalized_name = normalize_filename(file)
        file_path = os.path.join(bid_folder, file)

        # Skip empty or invalid filenames
        if not normalized_name:
            continue

        normalized_path = os.path.join(bid_folder, normalized_name)

        if normalized_name in processed_files:
            # Check if it's a true duplicate
            if is_duplicate_file(file_path, normalized_path):
                duplicates.append(file_path)
            else:
                # If content is different, keep both with numbered suffix
                base, ext = os.path.splitext(normalized_name)
                counter = 1
                while os.path.exists(
                    os.path.join(bid_folder, f"{base}_{counter}{ext}")
                ):
                    counter += 1
                new_name = f"{base}_{counter}{ext}"
                os.rename(file_path, os.path.join(bid_folder, new_name))
        else:
            if file != normalized_name:
                os.rename(file_path, normalized_path)
            processed_files.add(normalized_name)

    # Remove duplicates
    for dup in duplicates:
        try:
            os.remove(dup)
            logger.info(f"Removed duplicate file: {dup}")
        except Exception as e:
            logger.error(f"Error removing duplicate file {dup}: {str(e)}")


def filename_matches(file, expected_filename):
    """Check if the file matches the expected filename."""
    normalized_expected = normalize_filename(expected_filename)
    normalized_file = normalize_filename(file)

    # Replace spaces with underscores in the expected filename
    normalized_expected_underscored = normalized_expected.replace(" ", "_")

    # Check for exact match after normalization (with and without underscores)
    if (
        normalized_file == normalized_expected
        or normalized_file == normalized_expected_underscored
    ):
        return True

    # Check if all parts before the last underscore match
    expected_parts = normalized_expected.rsplit("_", 1)
    file_parts = normalized_file.rsplit("_", 1)
    if len(expected_parts) > 1 and len(file_parts) > 1:
        if expected_parts[0] == file_parts[0]:
            return True

    # Check if the file starts with the expected filename (ignoring extension)
    expected_without_ext = os.path.splitext(normalized_expected)[0]
    file_without_ext = os.path.splitext(normalized_file)[0]
    if file_without_ext.startswith(expected_without_ext):
        return True

    # Check if the file starts with the expected filename with underscores (ignoring extension)
    expected_without_ext_underscored = expected_without_ext.replace(" ", "_")
    if file_without_ext.startswith(expected_without_ext_underscored):
        return True

    return False


def wait_for_download(download_folder, expected_filename, timeout=600):
    """
    Wait for download to complete with progress tracking and size checks.
    Checks every 30 seconds until file is found and verified.

    Args:
        download_folder: Path to download folder
        expected_filename: Expected name of downloaded file
        timeout: Maximum wait time in seconds (default 10 minutes)
    """
    start_time = time.time()
    last_size = 0
    stall_count = 0
    check_interval = 30  # Check every 30 seconds

    while time.time() - start_time < timeout:
        # First check for partial downloads
        partial_downloads = [
            f
            for f in os.listdir(download_folder)
            if f.endswith(".crdownload") or f.endswith(".tmp")
        ]

        if partial_downloads:
            for partial_file in partial_downloads:
                file_path = os.path.join(download_folder, partial_file)
                current_size = os.path.getsize(file_path)

                if current_size == last_size:
                    stall_count += 1
                    if stall_count >= 3:  # If size hasn't changed for 3 checks
                        logger.warning(f"Download appears stalled for {partial_file}")
                        try:
                            os.remove(file_path)  # Remove stalled download
                            logger.info(f"Removed stalled download: {partial_file}")
                        except:
                            pass
                        return False
                else:
                    stall_count = 0
                    last_size = current_size

                size_mb = current_size / (1024 * 1024)
                logger.info(f"Downloading: {partial_file} - {size_mb:.2f} MB")
                time.sleep(check_interval)
                continue

        # Check for completed download
        for file in os.listdir(download_folder):
            if filename_matches(file, expected_filename):
                file_path = os.path.join(download_folder, file)

                # Wait additional time to ensure file is completely written
                time.sleep(5)

                # Verify file integrity
                try:
                    file_size = os.path.getsize(file_path)
                    if file_size == 0:
                        logger.warning(f"Downloaded file is empty: {file}")
                        continue

                    # Try to open and read the file
                    with open(file_path, "rb") as f:
                        f.read(1024)  # Try reading first 1KB

                    logger.info(
                        f"Download completed and verified: {file} ({format_size(file_size)})"
                    )
                    return True

                except (IOError, OSError) as e:
                    logger.error(f"Error verifying downloaded file: {str(e)}")
                    try:
                        os.remove(file_path)
                        logger.info(f"Removed corrupted file: {file}")
                    except:
                        pass
                    return False

        # If we haven't found the file yet, log and wait
        logger.info(
            f"Waiting for download to complete... Time elapsed: {int(time.time() - start_time)}s"
        )
        time.sleep(check_interval)

    logger.error(f"Download timeout after {timeout} seconds for {expected_filename}")
    return False


def move_latest_download(source_folder, destination_folder, expected_filename):
    """Move the downloaded file to the bid-specific folder."""
    # Use temp_download_folder instead of main_folder
    for file in os.listdir(temp_download_folder):
        if filename_matches(file, expected_filename):
            source_path = os.path.join(temp_download_folder, file)
            destination_path = os.path.join(destination_folder, file)
            max_retries = 5
            for attempt in range(max_retries):
                try:
                    safe_move(source_path, destination_path)
                    logger.info(f"Successfully moved file: {file}")
                    return
                except PermissionError:
                    logger.warning(
                        f"Permission error when moving file. Attempt {attempt + 1} of {max_retries}. Waiting 2 seconds..."
                    )
                    time.sleep(2)
            logger.error(f"Failed to move file {file} after {max_retries} attempts")
            return
    logger.warning(f"No matching file found for {expected_filename}")


def get_file_size(url):
    """Get the file size from the Content-Length header, handling redirects."""
    try:
        with requests.get(url, stream=True, allow_redirects=True) as response:
            response.raise_for_status()
            size = int(response.headers.get("Content-Length", 0))
            if size == 0:
                # If Content-Length is not provided, try to get size from streaming
                size = sum(
                    len(chunk) for chunk in response.iter_content(chunk_size=8192)
                )
            return size
    except requests.RequestException as e:
        logger.error(f"Error getting file size: {str(e)}")
        return 0


def format_size(size_bytes):
    """Format file size in a human-readable format."""
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} PB"


def get_download_progress(file_path, total_size):
    """Get the current download progress."""
    if os.path.exists(file_path):
        return os.path.getsize(file_path)
    return 0


def wait_for_crdownload(download_folder, timeout=600):
    """Wait for all .crdownload files to complete downloading."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        crdownload_files = [
            f for f in os.listdir(download_folder) if f.endswith(".crdownload")
        ]
        if not crdownload_files:
            return True

        logger.info(f"Waiting for {len(crdownload_files)} downloads to complete...")
        for file in crdownload_files:
            base_name = file[:-10]  # Remove .crdownload extension
            logger.info(f"Still downloading: {base_name}")

        time.sleep(30)  # Check every 30 seconds

    return False


def is_file_already_downloaded(bid_folder, filename):
    """Check if a file with the same normalized name exists in the folder."""
    normalized_name = normalize_filename(filename)
    existing_files = [normalize_filename(f) for f in os.listdir(bid_folder)]
    return normalized_name in existing_files


def wait_for_download_complete(download_folder):
    """Wait indefinitely until download completes and return the downloaded filename."""
    while True:
        # Check for .crdownload or .tmp files
        downloading_files = [
            f
            for f in os.listdir(download_folder)
            if f.endswith(".crdownload") or f.endswith(".tmp")
        ]

        if downloading_files:
            # Log download in progress
            for file in downloading_files:
                base_name = file.replace(".crdownload", "").replace(".tmp", "")
                logger.info(f"Download in progress: {base_name}")
            time.sleep(2)
            continue

        # If no downloads in progress, look for the most recently modified file
        files = [
            f
            for f in os.listdir(download_folder)
            if os.path.isfile(os.path.join(download_folder, f))
        ]

        if files:
            # Sort by modification time to get the most recent file
            files.sort(
                key=lambda x: os.path.getmtime(os.path.join(download_folder, x)),
                reverse=True,
            )
            return files[0]

        # If no files found, wait before checking again
        time.sleep(2)


def download_attachments(driver, bid_number):
    """Download attachments with duplicate handling."""
    bid_folder = os.path.join(main_folder, bid_number.replace("/", "_"))

    # Only create folder if there are attachments to download
    attachments = driver.find_elements(By.CSS_SELECTOR, "a.iv-download-file")
    if not attachments:
        return []

    os.makedirs(bid_folder, exist_ok=True)

    bid_url = driver.current_url
    downloaded_attachments = []

    while True:
        attachments = driver.find_elements(By.CSS_SELECTOR, "a.iv-download-file")
        if not attachments:
            break

        # Get all attachment info first
        attachment_info = []
        for attachment in attachments:
            try:
                name = attachment.find_element(
                    By.CSS_SELECTOR, "span[data-iv-role='label']"
                ).text.strip()
                if not name:  # Skip empty names
                    continue

                href = attachment.get_attribute("href")
                normalized_name = normalize_filename(name)

                # Skip if already downloaded
                if not is_file_already_downloaded(bid_folder, normalized_name):
                    attachment_info.append((name, normalized_name, href))
                else:
                    logger.info(
                        f"Skipping already downloaded attachment: {normalized_name}"
                    )
            except Exception as e:
                logger.error(f"Error getting attachment info: {str(e)}")
                continue

        if not attachment_info:
            break

        total_attachments = len(attachment_info)
        logger.info(
            f"Found {total_attachments} new attachments to download for bid {bid_number}"
        )

        # Process each attachment
        for index, (name, normalized_name, href) in enumerate(attachment_info, 1):
            try:
                logger.info(
                    f"Processing attachment {index} of {total_attachments}: {normalized_name}"
                )

                # Clear download folder before each download
                for file in os.listdir(temp_download_folder):
                    try:
                        os.remove(os.path.join(temp_download_folder, file))
                    except:
                        pass

                # Navigate to download URL
                base_url = "https://emma.maryland.gov"
                download_url = base_url + href if href.startswith("/") else href
                driver.get(download_url)

                # Click download button and wait for download
                try:
                    download_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, "button#proxyActionBar_x__cmdEnd")
                        )
                    )
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center'});",
                        download_button,
                    )
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", download_button)
                    logger.info(f"Started download for: {name}")

                    # Wait indefinitely for download to complete
                    downloaded_file = wait_for_download_complete(temp_download_folder)

                    # Move file to bid folder
                    if downloaded_file:
                        source_path = os.path.join(
                            temp_download_folder, downloaded_file
                        )
                        dest_path = os.path.join(bid_folder, normalized_name)

                        try:
                            safe_move(source_path, dest_path)
                            downloaded_attachments.append(normalized_name)
                            logger.info(
                                f"Successfully downloaded and moved: {normalized_name}"
                            )
                        except Exception as e:
                            logger.error(f"Error moving file: {str(e)}")
                            continue

                except TimeoutException:
                    logger.error(
                        f"Download button not found or not clickable for: {name}"
                    )
                    continue

            except Exception as e:
                logger.error(f"Error processing attachment {normalized_name}: {str(e)}")

            finally:
                # Return to bid page after each download
                driver.get(bid_url)
                time.sleep(2)

    # After all downloads complete, handle duplicates and update Excel
    handle_duplicate_files(bid_folder)

    # Get final list of normalized filenames for Excel
    final_files = [
        f for f in os.listdir(bid_folder) if os.path.isfile(os.path.join(bid_folder, f))
    ]

    # Update Excel with downloaded attachments
    if downloaded_attachments:
        excel_file = os.path.join(main_folder, f"{script_name}.xlsx")
        update_excel_attachments(excel_file, bid_number, downloaded_attachments)

    return final_files


def move_files_to_bid_folder(main_folder, bid_folder):
    """Move downloaded files to the bid-specific folder."""
    for file in os.listdir(main_folder):
        source = os.path.join(main_folder, file)
        if os.path.isfile(source) and not file.endswith(".crdownload"):
            destination = os.path.join(bid_folder, file)
            safe_move(source, destination)


def update_excel(bid_data):
    """Update the Excel file with the latest bid data."""
    # Put Excel file in main_folder
    excel_file = os.path.join(main_folder, f"{script_name}.xlsx")

    column_order = [
        "SL No",
        "Posted Date",
        "Response Date",
        "Notice Type",
        "Solicitation Number",
        "Solicitation Title",
        "Agency",
        "Category",
        "Description",
        "Additional Summary",
        "Contracting Office Address",
        "Contact Information",
        "Bid Detail Page URL",
        "Attachments",
    ]

    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file, dtype=str)
    else:
        df = pd.DataFrame(columns=column_order)
        logger.info(f"Created new Excel file: {excel_file}")

    # Update the Attachments column for the current bid
    bid_folder = os.path.join(
        main_folder, bid_data["Solicitation Number"].replace("/", "_")
    )
    if os.path.exists(bid_folder):
        attachments = [
            f
            for f in os.listdir(bid_folder)
            if os.path.isfile(os.path.join(bid_folder, f))
        ]
        bid_data["Attachments"] = ", ".join(attachments)

    # Check if the solicitation number already exists
    existing_row = df[df["Solicitation Number"] == bid_data["Solicitation Number"]]
    if not existing_row.empty:
        # Update the existing row
        index = existing_row.index[0]
        for column in df.columns:
            if column in bid_data:
                df.at[index, column] = bid_data[column]
        logger.info(
            f"Updated existing entry for Solicitation Number: {bid_data['Solicitation Number']}"
        )
    else:
        # Add a new row
        df = pd.concat([df, pd.DataFrame([bid_data])], ignore_index=True)
        logger.info(
            f"Added new entry for Solicitation Number: {bid_data['Solicitation Number']}"
        )

    df = df.reindex(columns=column_order)

    for date_col in ["Posted Date", "Response Date"]:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.strftime(
            "%Y-%m-%d"
        )

    df["SL No"] = range(1, len(df) + 1)

    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        for column in worksheet.columns:
            column[0].style = "Normal"
            column[0].number_format = "@"

    logger.info(f"Updated Excel file: {excel_file}")


def update_excel_attachments(excel_file, bid_number, attachments):
    """Update Excel with normalized attachment names."""
    if not os.path.exists(excel_file):
        df = pd.DataFrame(columns=["Solicitation Number", "Attachments"])
    else:
        df = pd.read_excel(excel_file, dtype=str)

    # Normalize attachment names
    normalized_attachments = [normalize_filename(att) for att in attachments if att]

    bid_index = df[df["Solicitation Number"] == bid_number].index
    if not bid_index.empty:
        df.at[bid_index[0], "Attachments"] = " | ".join(normalized_attachments)
        logger.info(f"Updated Attachments for bid {bid_number} in Excel file")
    else:
        logger.warning(f"Solicitation Number {bid_number} not found in Excel file")

    df.to_excel(excel_file, index=False)


def cleanup_script_download_folder():
    """This function is no longer needed since we're using main_folder directly"""
    pass


def mark_folder_completed():
    """Rename the folder to indicate completion."""
    try:
        completed_folder = main_folder.replace("_IN_PROGRESS", "_COMPLETED")
        if os.path.exists(main_folder):
            os.rename(main_folder, completed_folder)
            logger.info(f"Marked folder as completed: {completed_folder}")
    except Exception as e:
        logger.error(f"Error marking folder as completed: {str(e)}")


def load_processed_bids():
    """Load previously processed bid numbers from JSON file."""
    processed_bids_file = os.path.join(main_folder, "processed_bids.json")
    if os.path.exists(processed_bids_file):
        with open(processed_bids_file, "r") as f:
            return set(json.load(f))
    return set()


def save_processed_bids(processed_bids):
    """Save processed bid numbers to JSON file."""
    processed_bids_file = os.path.join(main_folder, "processed_bids.json")
    with open(processed_bids_file, "w") as f:
        json.dump(list(processed_bids), f)


def load_cache():
    """Load the cache from JSON file."""
    cache_file = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        "cache",
        f"{os.path.splitext(os.path.basename(__file__))[0]}_cache.json",
    )
    if os.path.exists(cache_file):
        with open(cache_file, "r") as f:
            return json.load(f)
    return {}


def save_to_cache(bid_url, bid_details):
    """Save bid details to cache."""
    cache_file = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        "cache",
        f"{os.path.splitext(os.path.basename(__file__))[0]}_cache.json",
    )

    cache = load_cache()

    # Convert date to YYYY-MM-DD format
    try:
        posted_date = datetime.strptime(
            bid_details.get("Posted Date", ""), "%Y-%m-%d"
        ).strftime("%Y-%m-%d")
    except ValueError:
        posted_date = bid_details.get("Posted Date", "")

    # Add/update the bid in cache
    cache[bid_url] = {
        "posted_date": posted_date,
        "last_checked": datetime.now().strftime("%Y-%m-%d"),
    }

    # Remove bids older than 3 months
    three_months_ago = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
    cache = {
        k: v for k, v in cache.items() if v.get("posted_date", "") >= three_months_ago
    }

    # Save updated cache
    os.makedirs(os.path.dirname(cache_file), exist_ok=True)
    with open(cache_file, "w") as f:
        json.dump(cache, f, indent=2)


def should_process_bid_link(bid_url):
    """Check if bid should be processed based on cache."""
    cache = load_cache()
    if bid_url in cache:
        bid_data = cache[bid_url]
        logger.info(
            f"⏭️ Skipping previously processed bid from {bid_data['posted_date']}"
        )
        return False
    return True


def main(days_to_scrape=2):
    logger.info(f"🚀 Bids Extraction Started for {script_name}")
    driver = setup_driver()

    # Create a thread to keep the session alive
    def session_keeper():
        while True:
            try:
                keep_session_alive(driver)
                time.sleep(60)  # Perform actions every minute
            except:
                break

    session_thread = threading.Thread(target=session_keeper, daemon=True)
    session_thread.start()

    # Load previously processed bids
    processed_bids = load_processed_bids()

    # Calculate the start date based on the number of days to scrape
    start_date = datetime.now() - timedelta(days=days_to_scrape)

    logger.info(
        f"Scraping bids from {start_date.strftime('%Y-%m-%d')} to {datetime.now().strftime('%Y-%m-%d')}"
    )

    try:
        driver.get("https://emma.maryland.gov/page.aspx/en/rfp/request_browse_public")
        logger.info("Navigated to the main page")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "body_x_grid_grd"))
        )

        try:
            status_dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "body_x_selStatusCode_search"))
            )
            status_dropdown.click()
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[text()='Open']"))
            ).click()
            time.sleep(2)
            logger.info("Set 'Open' filter successfully")
        except TimeoutException:
            logger.warning("'Open' filter is already set.")

        bid_links = extract_bid_links(driver, start_date)
        logger.info(f"{len(bid_links)} bid links found, checking for new bids...")

        new_bids = []
        for link, posted_date, response_date in bid_links:
            try:
                driver.get(link)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located(
                        (By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblProcessCode")
                    )
                )
                bid_number = driver.find_element(
                    By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblProcessCode"
                ).text

                if bid_number not in processed_bids:
                    new_bids.append((link, posted_date, response_date))
                    processed_bids.add(bid_number)
            except Exception as e:
                logger.error(f"Error checking bid number for {link}: {str(e)}")
                continue

        logger.info(f"Found {len(new_bids)} new bids to process")

        for index, (link, posted_date, response_date) in enumerate(new_bids, start=1):
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    logger.info(
                        f"Processing bid {index} of {len(new_bids)} (Attempt {attempt + 1})"
                    )

                    # Add cache check here
                    if not should_process_bid_link(link):
                        break

                    driver.get(link)  # Navigate to the bid page
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located(
                            (By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_x_lblProcessCode")
                        )
                    )

                    bid_details = extract_bid_details(driver, link)
                    if bid_details:
                        bid_details["SL No"] = index
                        bid_details["Posted Date"] = posted_date.strftime("%Y-%m-%d")
                        bid_details["Response Date"] = datetime.strptime(
                            response_date, "%m/%d/%Y"
                        ).strftime("%Y-%m-%d")

                        # Download attachments
                        downloaded_attachments = download_attachments(
                            driver, bid_details["Solicitation Number"]
                        )

                        # Update bid_details with correct attachment filenames
                        bid_details["Attachments"] = " | ".join(downloaded_attachments)

                        # Update Excel (Attachments column will be updated here)
                        update_excel(bid_details)

                        # Save to cache after successful processing
                        save_to_cache(link, bid_details)

                        logger.info(
                            f"✅ Processed bid: {bid_details['Solicitation Number']}"
                        )
                        break  # Break the retry loop if successful
                    else:
                        raise Exception("Could not extract bid details")

                except (TimeoutException, WebDriverException) as e:
                    logger.warning(f"Attempt {attempt + 1} failed: {str(e)}")
                    if attempt < max_retries - 1:
                        logger.info("Restarting the WebDriver and retrying...")
                        driver.quit()
                        driver = setup_driver()
                    else:
                        error_handler(
                            f"Failed to process bid at {link} after {max_retries} attempts"
                        )

                except Exception as e:
                    error_handler(f"Error processing bid at {link}: {str(e)}")
                    break  # Break the retry loop for non-WebDriver related errors

            time.sleep(2)

        logger.info(
            "✅ Bids posted within 1 day successfully extracted and saved to Excel with proper formatting."
        )
        logger.info(
            f"🎉 All Bids and Attachments Extraction Successfully Completed for '{script_name}'"
        )
        play_notification()  # Play notification sound on successful completion
        mark_folder_completed()  # Mark the folder as completed

    except Exception as e:
        error_handler(f"An unexpected error occurred: {str(e)}")
        import traceback

        logger.error(traceback.format_exc())

    finally:
        # Stop the session keeper thread
        session_thread = None
        driver.quit()
        logger.info("WebDriver closed")
        cleanup_script_download_folder()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scrape eMaryland eMMA bids")
    parser.add_argument(
        "--days", type=int, default=2, help="Number of days to scrape (default: 2)"
    )
    args = parser.parse_args()

    main(days_to_scrape=args.days)
