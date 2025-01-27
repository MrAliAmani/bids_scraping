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
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
import requests
import threading
import urllib3
import openpyxl
from openpyxl import Workbook
from bs4 import BeautifulSoup
import concurrent.futures
import argparse
from urllib.parse import urljoin
import shutil
import winsound  # For playing notification sounds
import sys
from selenium.webdriver.common.actions.action_builder import ActionBuilder
import glob
import zipfile
from utils.utils import safe_move, play_notification_sound
import base64
from PIL import Image
import tempfile
import atexit
import logging
import json
import random
from fake_useragent import UserAgent
import undetected_chromedriver as uc

# Set up logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuration
DAYS_TO_SEARCH = 2  # Default to 2 days, can be changed for testing

# Constants
SCRIPT_NAME = "08_SFCityPartner"
BASE_URL = "https://sfcitypartner.sfgov.org/pages/index.aspx"

# Create folder structure
yesterday = datetime.now() - timedelta(days=1)
BASE_FOLDER = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
date_folder = os.path.join(BASE_FOLDER, yesterday.strftime("%Y-%m-%d"))
MAIN_FOLDER = os.path.join(date_folder, f"{SCRIPT_NAME}_IN_PROGRESS")
SCRIPT_FOLDER = MAIN_FOLDER
TEMP_DOWNLOAD_FOLDER = os.path.join(MAIN_FOLDER, f"{SCRIPT_NAME}")
EXCEL_FILE = os.path.join(MAIN_FOLDER, f"{SCRIPT_NAME}.xlsx")

# Define progress file path after MAIN_FOLDER is defined
PROGRESS_FILE = os.path.join(MAIN_FOLDER, "scraping_progress.json")

# Create necessary directories
os.makedirs(date_folder, exist_ok=True)
os.makedirs(MAIN_FOLDER, exist_ok=True)
os.makedirs(TEMP_DOWNLOAD_FOLDER, exist_ok=True)

# Chrome options setup
chrome_options = Options()
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--ignore-ssl-errors")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Enable safe browsing
chrome_options.add_argument("--safebrowsing-disable-download-protection")
chrome_options.add_argument("--safebrowsing-disable-extension-blacklist")

# Set download preferences
prefs = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True,
    "plugins.always_open_pdf_externally": True,
    "download.default_directory": TEMP_DOWNLOAD_FOLDER,
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()

# Add this global variable near the top of the file
scraped_bids = set()


# Function to handle errors
def handle_error(error_message):
    """Handle errors by playing a notification and pausing the script."""
    print(f"❗ Error: {error_message}")
    play_notification_sound()
    input("Press Enter to continue...")


def get_random_user_agent():
    """Return a random user agent string."""
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    ]
    return random.choice(user_agents)


def get_random_browser_config():
    """Generate random browser configuration."""
    ua = UserAgent()

    # Random viewport sizes
    viewports = [(1920, 1080), (1366, 768), (1536, 864), (1440, 900), (1280, 720)]

    # Random platform
    platforms = ["Windows", "Macintosh", "Linux"]

    viewport = random.choice(viewports)
    return {
        "user_agent": ua.random,
        "viewport_width": viewport[0],
        "viewport_height": viewport[1],
        "platform": random.choice(platforms),
    }


def setup_driver():
    """Set up and return a configured Chrome WebDriver."""
    chrome_options = Options()

    # Add random user agent
    user_agent = get_random_user_agent()
    chrome_options.add_argument(f"user-agent={user_agent}")

    # Basic options
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Anti-detection measures
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    # Download settings
    prefs = {
        "download.default_directory": TEMP_DOWNLOAD_FOLDER,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()

    # Execute CDP commands to prevent detection
    driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": user_agent})
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        """
        },
    )

    return driver


def wait_for_element(driver, by, value, timeout=60):
    """Wait for an element to be present and visible."""
    return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, value)))


def perform_advanced_search(driver):
    """Perform advanced search on the SF City Partner website."""
    logging.info("Starting advanced search...")

    try:
        # Wait for page to be fully loaded
        time.sleep(random.uniform(3, 5))

        # Wait and click Advanced button with retry
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # Wait for Advanced button
                advanced_button = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//button[@id='RESP_INQA_WK_AUC_SRCH_ADVBSC_PB']")
                    )
                )

                # Random delay before clicking
                time.sleep(random.uniform(1, 2))

                # Scroll into view with offset
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    advanced_button,
                )
                time.sleep(random.uniform(0.5, 1))

                # Move mouse to button with random offset
                action = ActionChains(driver)
                action.move_to_element_with_offset(
                    advanced_button, random.randint(-5, 5), random.randint(-5, 5)
                )
                action.perform()
                time.sleep(random.uniform(0.5, 1))

                # Click using different methods
                try:
                    advanced_button.click()
                except:
                    try:
                        action.click().perform()
                    except:
                        driver.execute_script("arguments[0].click();", advanced_button)

                logging.info("Clicked Advanced search button")
                time.sleep(random.uniform(8, 10))
                break

            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                logging.warning(
                    f"Attempt {attempt + 1} failed to click Advanced button: {str(e)}"
                )
                time.sleep(random.uniform(3, 5))

        # Wait for and handle status dropdown
        status_select = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "RESP_INQA_WK_AUC_STATUS"))
        )
        time.sleep(random.uniform(1, 2))

        # Select status using different methods
        try:
            Select(status_select).select_by_value("P")
        except:
            driver.execute_script(
                "arguments[0].value = 'P'; arguments[0].dispatchEvent(new Event('change'));",
                status_select,
            )

        logging.info("Selected Event Status: Posted")
        time.sleep(random.uniform(3, 5))

        # Handle start date input
        start_date_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "RESP_INQA_WK_AUC_FROM_START_DT"))
        )
        start_date = (datetime.now() - timedelta(days=DAYS_TO_SEARCH)).strftime(
            "%m/%d/%Y"
        )

        # Clear and set date with random delays
        start_date_input.clear()
        time.sleep(random.uniform(0.3, 0.7))
        for char in start_date:
            start_date_input.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))

        logging.info(f"Set start date to: {start_date}")
        time.sleep(random.uniform(3, 5))

        # Handle search button
        for attempt in range(max_retries):
            try:
                search_button = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//button[@id='RESP_INQA_WK_INQ_AUC_GO_PB']")
                    )
                )

                # Random delay before clicking
                time.sleep(random.uniform(1, 2))

                # Scroll and move to button
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    search_button,
                )
                time.sleep(random.uniform(0.5, 1))

                action = ActionChains(driver)
                action.move_to_element_with_offset(
                    search_button, random.randint(-5, 5), random.randint(-5, 5)
                )
                action.perform()
                time.sleep(random.uniform(0.5, 1))

                # Try different click methods
                try:
                    search_button.click()
                except:
                    try:
                        action.click().perform()
                    except:
                        driver.execute_script("arguments[0].click();", search_button)

                logging.info("Clicked search button")

                # Wait for results with longer timeout
                WebDriverWait(driver, 90).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//ul[@class='list-group']")
                    )
                )
                time.sleep(random.uniform(12, 15))
                logging.info("Search results loaded")
                break

            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                logging.warning(
                    f"Attempt {attempt + 1} failed to click Search button: {str(e)}"
                )
                time.sleep(random.uniform(3, 5))

    except Exception as e:
        logging.error(f"Error in advanced search: {str(e)}")
        logging.error(f"Current URL: {driver.current_url}")
        raise

    if "data:," in driver.current_url:
        raise Exception("Search results page did not load properly")


def extract_auc_ids(driver):
    """Extract AUC IDs from the search results page."""
    auc_ids = []
    try:
        print("Extracting AUC IDs...")
        wait_for_element(driver, By.XPATH, "//ul[@class='list-group']", timeout=60)

        # Scroll through search results
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # Extract AUC IDs
        auc_id_elements = driver.find_elements(
            By.CSS_SELECTOR,
            "#l0RESP_INQA_HD_VW_GR\\$0 li span[data-if-label='tdEventId']",
        )

        # Check if there are any bids (excluding header)
        if len(auc_id_elements) <= 1:  # Only header present or no elements
            print("No bids found in search results")
            return None  # Return None instead of empty list to indicate no bids

        for element in auc_id_elements[1:]:  # Skip the first element
            auc_id = element.text.split("-")[-1]  # Extract the numeric part
            auc_ids.append(auc_id)

        print(f"📊 AUC IDs Successfully Extracted: {len(auc_ids)}")
    except Exception as e:
        print(f"Error extracting AUC IDs: {str(e)}")
        return None  # Return None on error
    return auc_ids


def construct_bid_url(auc_id):
    """Construct the bid detail URL using the AUC ID."""
    base_url = "https://sfcitypartner.sfgov.org/pages/Events-BS3/event-details.aspx"
    params = {
        "Page": "AUC_RESP_INQ_DTL",
        "Action": "U",
        "AUC_ID": auc_id,
        "AUC_ROUND": "1",
        "AUC_VERSION": "1",
        "BIDDER_ID": "0000000001",
        "BIDDER_LOC": "1",
        "BIDDER_SETID": "SHARE",
        "BIDDER_TYPE": "B",
        "BUSINESS_UNIT": "SFGOV",
    }
    return f"{base_url}?{'&'.join(f'{k}={v}' for k, v in params.items())}"


def get_attachment_names_from_page(driver):
    """Get the list of attachment names from the page."""
    attachment_names = []
    attachment_elements = driver.find_elements(
        By.CSS_SELECTOR, "span[data-if-label='ViewAttachFileName']"
    )
    for element in attachment_elements:
        attachment_names.append(element.text)
    return attachment_names


def move_downloaded_files(driver, bid_folder, event_id, max_retries=3, retry_delay=5):
    """Move downloaded files to the bid folder with retry mechanism."""
    files_moved = False

    for attempt in range(max_retries):
        try:
            # Ensure both folders exist
            os.makedirs(TEMP_DOWNLOAD_FOLDER, exist_ok=True)
            os.makedirs(bid_folder, exist_ok=True)

            # Get list of files in download folder
            files = glob.glob(os.path.join(TEMP_DOWNLOAD_FOLDER, "*"))

            if not files:
                print(
                    f"No files found in {TEMP_DOWNLOAD_FOLDER} (Attempt {attempt + 1})"
                )
                time.sleep(retry_delay)
                continue

            moved_files = []
            for file in files:
                try:
                    file_name = os.path.basename(file)
                    destination = os.path.join(bid_folder, file_name)

                    # Handle zip files
                    if file_name.lower().endswith(".zip"):
                        with zipfile.ZipFile(file, "r") as zip_ref:
                            zip_ref.extractall(bid_folder)
                        os.remove(file)
                        moved_files.extend(zip_ref.namelist())
                        print(f"Extracted {file_name} to {bid_folder}")
                    else:
                        # Move file with overwrite if exists
                        if os.path.exists(destination):
                            os.remove(destination)
                        shutil.move(file, destination)
                        moved_files.append(file_name)
                        print(f"Moved {file_name} to {bid_folder}")

                    files_moved = True

                except Exception as e:
                    print(f"Error processing file {file}: {str(e)}")

            if files_moved:
                print(f"Successfully moved {len(moved_files)} files")
                update_excel_with_attachments(EXCEL_FILE, event_id, moved_files)
                break

        except Exception as e:
            print(f"Error during move attempt {attempt + 1}: {str(e)}")
            if attempt == max_retries - 1:
                print("Failed to move files after all retries")
            time.sleep(retry_delay)

    return files_moved


def download_attachments(driver, event_id):
    """Download attachments for a specific bid using XPath selectors and JavaScript."""
    try:
        print(f"Downloading attachments for bid {event_id}")

        # Click the "See Attachments" button
        attachments_button = wait_for_element(
            driver, By.ID, "RESP_INQ_DL0_WK_AUC_DOWNLOAD_PB"
        )
        if attachments_button:
            print("Attachments button found.")
            ActionChains(driver).move_to_element(attachments_button).click().perform()
            print("Clicked on the 'See Attachments' button.")
        else:
            print("Attachments button not found.")
            return "Attachments button not found"

        time.sleep(5)  # Wait for the attachments to load

        # Create the bid folder with the correct name
        bid_folder = os.path.join(MAIN_FOLDER, f"SFGOV-{event_id}")
        os.makedirs(bid_folder, exist_ok=True)

        attachment_count = 0

        # Loop through attachment buttons
        for i in range(2, 100):  # Adjust the range as needed
            try:
                # Find download button in current row
                xpath = f"/html/body/div[8]/div/div[2]/div[1]/form/div[5]/div/div/div[3]/table/tbody/tr[{i}]/td[5]/button[2]"
                attachment_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )

                # Click download button
                driver.execute_script("arguments[0].click();", attachment_button)
                print(f"Clicked attachment button {i-1}")
                time.sleep(5)  # Wait for popup

                # Handle download popup
                try:
                    # Wait for and click download button in popup
                    popup_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "downloadButton"))
                    )
                    driver.execute_script("arguments[0].click();", popup_button)
                    print("Clicked download button using JavaScript")
                    attachment_count += 1

                    # Wait for download to complete
                    wait_for_download(driver, 60)  # Wait up to 60 seconds

                    # Move downloaded files to bid folder
                    move_downloaded_files(driver, bid_folder, event_id)

                except Exception as e:
                    print(f"Error handling attachment: {str(e)}")
                    continue

                time.sleep(2)  # Wait between downloads

            except TimeoutException:
                print(f"No more attachment buttons found after {i-2} attachments.")
                break
            except Exception as e:
                print(f"Error with attachment {i-1}: {str(e)}")
                continue

        # Get final list of downloaded files
        attachment_names = get_attachment_names(bid_folder)

        if not attachment_names:
            print("No attachments downloaded")
            return "No attachments downloaded"

        # Update Excel file with attachment names
        update_excel_with_attachments(EXCEL_FILE, event_id, attachment_names)

        print(
            f"Successfully downloaded {len(attachment_names)} attachments for bid {event_id}"
        )
        return ",".join(attachment_names)

    except Exception as e:
        error_message = f"Error downloading attachments for bid {event_id}: {str(e)}"
        handle_error(error_message)
        return "Error downloading attachments"


def get_attachment_names(bid_folder):
    """Get the list of attachment names from the bid folder, including contents of unzipped folders."""
    attachment_names = []
    for root, dirs, files in os.walk(bid_folder):
        for file in files:
            attachment_names.append(file)
    return attachment_names


def update_excel(bid_data):
    """Update the Excel file with new bid data, creating the file if it doesn't exist."""
    try:
        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE)
        else:
            df = pd.DataFrame(
                columns=[
                    "SL No",
                    "Posted Date",
                    "Response Date",
                    "Notice Type",
                    "Solicitation Number",
                    "Solicitation Title",
                    "Agency",
                    "Category",
                    "Description",
                    "Additional Summary, if any",
                    "Contracting Office Address",
                    "Contact Information",
                    "Bid Detail Page URL",
                    "Attachments",
                ]
            )

        new_row = pd.DataFrame([bid_data])
        df = pd.concat([df, new_row], ignore_index=True)

        # Ensure the directory exists
        os.makedirs(os.path.dirname(EXCEL_FILE), exist_ok=True)

        df.to_excel(EXCEL_FILE, index=False)
        print(
            f"Successfully updated Excel file with bid {bid_data['Solicitation Number']}"
        )
    except Exception as e:
        print(f"Error updating Excel file: {str(e)}")


def update_excel_with_attachments(excel_file, event_id, attachment_names):
    """Update the Excel file with the correct attachment names for a specific bid."""
    try:
        if os.path.exists(excel_file):
            df = pd.read_excel(excel_file)
            mask = df["Solicitation Number"] == f"SFGOV-{event_id}"
            if mask.any():
                df.loc[mask, "Attachments"] = ",".join(attachment_names)
                df.to_excel(excel_file, index=False)
                print(
                    f"Updated Excel file with correct attachment names for bid {event_id}"
                )
            else:
                print(f"Bid {event_id} not found in Excel file")
        else:
            print(f"Excel file not found: {excel_file}")
    except Exception as e:
        print(f"Error updating Excel file with attachments: {str(e)}")


def handle_missing_date(bid_id, date_type):
    """
    Handle missing date by prompting user for action.

    Args:
        bid_id (str): The bid ID
        date_type (str): Type of missing date (Posted/Response)

    Returns:
        bool: True to continue with current bid, False to skip it
    """
    print(f"\n⚠️ Warning: {date_type} Date is missing for bid {bid_id}")
    print("Options:")
    print("1. Press Enter to continue scraping this bid")
    print("2. Press 'S' to skip this bid")

    while True:
        choice = input("Your choice (Enter/S): ").strip().upper()
        if choice == "":
            return True
        elif choice == "S":
            return False
        else:
            print("Invalid input. Please press Enter to continue or 'S' to skip.")


def is_lookup_page(driver):
    """Check if we're on the lookup page."""
    return "lookup.aspx" in driver.current_url


def handle_lookup_page(driver, bid_id):
    """Handle the case when we hit the lookup page."""
    if is_lookup_page(driver):
        print(f"\n⚠️ Hit lookup page for bid {bid_id}")
        print("Skipping this bid due to lookup page redirect")
        return True
    return False


def extract_bid_details(driver, event_id, sl_no, total_bids):
    """Extract details from a specific bid page."""
    # Check for lookup page before proceeding
    if handle_lookup_page(driver, event_id):
        return None

    global scraped_bids

    print(f"Processing bid {sl_no} out of {total_bids}")

    if event_id in scraped_bids:
        print(f"Bid {event_id} has already been scraped. Skipping...")
        return None

    bid_details = {
        "SL No": str(sl_no),
        "Posted Date": "",
        "Response Date": "",
        "Notice Type": "",
        "Solicitation Number": "",
        "Solicitation Title": "",
        "Agency": "",
        "Category": "",
        "Description": "",
        "Additional Summary, if any": "",
        "Contracting Office Address": "",
        "Contact Information": "",
        "Bid Detail Page URL": driver.current_url,
        "Attachments": "",
    }

    try:
        print(f"Extracting details for bid {event_id}")

        # Add lookup page check after each major operation
        if handle_lookup_page(driver, event_id):
            return None

        # Extract Posted Date
        try:
            posted_date_element = wait_for_element(
                driver,
                By.XPATH,
                "//div[@class='text-muted']/span[@data-if-label='startDate']",
            )
            posted_date = posted_date_element.text.split()[0]
            if posted_date != "This":
                bid_details["Posted Date"] = datetime.strptime(
                    posted_date, "%m/%d/%Y"
                ).strftime("%Y-%m-%d")
            else:
                # Handle missing Posted Date
                if not handle_missing_date(event_id, "Posted"):
                    print(f"Skipping bid {event_id} due to missing Posted Date")
                    return None
        except (TimeoutException, ValueError) as e:
            if not handle_missing_date(event_id, "Posted"):
                print(f"Skipping bid {event_id} due to missing Posted Date")
                return None

        # Extract Response Date
        try:
            response_date_element = wait_for_element(
                driver,
                By.XPATH,
                "//div[@class='bold']/span[@data-if-label='expireDate']",
                timeout=5,
            )
            response_date = response_date_element.text.split()[0]
            if response_date != "This":
                bid_details["Response Date"] = datetime.strptime(
                    response_date, "%m/%d/%Y"
                ).strftime("%Y-%m-%d")
            else:
                # Handle missing Response Date
                if not handle_missing_date(event_id, "Response"):
                    print(f"Skipping bid {event_id} due to missing Response Date")
                    return None
        except (TimeoutException, ValueError) as e:
            if not handle_missing_date(event_id, "Response"):
                print(f"Skipping bid {event_id} due to missing Response Date")
                return None

        # Extract Solicitation Number
        bid_details["Solicitation Number"] = wait_for_element(
            driver,
            By.XPATH,
            "//span[@data-if-source='#RESP_AUC_H0B_WK_AUC_ID_BUS_UNIT']",
        ).text

        # Extract Solicitation Title
        bid_details["Solicitation Title"] = wait_for_element(
            driver, By.XPATH, "//h3[@data-if-label='eventName']"
        ).text

        # Extract Description
        bid_details["Description"] = wait_for_element(
            driver, By.XPATH, "//div[@data-if-label='descriptiondetails']"
        ).text

        # Extract Contact Information
        contact_name = wait_for_element(
            driver, By.XPATH, "//strong[@data-if-label='contactName']"
        ).text
        contact_phone = wait_for_element(
            driver, By.XPATH, "//span[@data-if-label='phoneText']"
        ).text
        contact_email = wait_for_element(
            driver, By.XPATH, "//span[@data-if-label='emffffailText']"
        ).text
        bid_details["Contact Information"] = (
            f"{contact_name}, {contact_phone}, {contact_email}"
        )

        # Download attachments
        bid_details["Attachments"] = download_attachments(driver, event_id)

        print(f"Successfully extracted details for bid {event_id}")

    except Exception as e:
        error_message = f"Error extracting bid details for {event_id}: {str(e)}"
        handle_error(error_message)

    finally:
        # Check one final time before returning
        if handle_lookup_page(driver, event_id):
            return None

        update_excel(bid_details)
        scraped_bids.add(event_id)

        bid_folder = os.path.join(SCRIPT_FOLDER, f"SFGOV-{event_id}")
        os.makedirs(bid_folder, exist_ok=True)
        move_downloaded_files(driver, bid_folder, event_id)

        attachment_names = get_attachment_names(bid_folder)
        update_excel_with_attachments(EXCEL_FILE, event_id, attachment_names)

    return bid_details


def click_at_coordinates(driver, x, y):
    """
    Clicks at the specified (x, y) coordinates using JavaScriptExecutor.

    Args:
        driver (webdriver.Chrome): The Selenium WebDriver instance.
        x (int): The x-coordinate in pixels.
        y (int): The y-coordinate in pixels.

    Returns:
        None
    """
    try:
        # Execute JavaScript to click at (x, y)
        driver.execute_script(f"document.elementFromPoint({x}, {y}).click();")
        print(f"📍 Clicked at coordinates ({x}, {y})")
        time.sleep(1)  # Wait for the action to complete
    except Exception as e:
        handle_error(f"Failed to click at ({x}, {y}): {e}")


def ensure_element_visibility(driver, element):
    """
    Ensures that the element is visible and not overlapped by other elements.

    Args:
        driver (webdriver.Chrome): The Selenium WebDriver instance.
        element (WebElement): The target WebElement.

    Returns:
        bool: True if the element is visible and not overlapped, False otherwise.
    """
    try:
        # Get the bounding rectangle of the element
        rect = element.rect
        x, y = rect["x"], rect["y"]
        width, height = rect["width"], rect["height"]

        # Use JavaScript to get the element at the center of the target element
        center_x = x + width / 2
        center_y = y + height / 2
        element_at_point = driver.execute_script(
            "return document.elementFromPoint(arguments[0], arguments[1]);",
            center_x,
            center_y,
        )

        if element_at_point == element:
            return True
        else:
            print(f"❗ Element is overlapped by {element_at_point.tag_name}.")
            return False
    except Exception as e:
        handle_error(f"❗ Error while verifying element visibility: {e}")
        return False


def wait_for_download(driver, timeout):
    """Wait for download to complete."""
    print(f"Waiting for download to complete (up to {timeout} seconds)...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        if not any(
            filename.endswith(".crdownload")
            for filename in os.listdir(TEMP_DOWNLOAD_FOLDER)
        ):
            print("Download completed.")
            return True
        time.sleep(1)
    print("Download timed out.")
    return False


def fetch_html(url):
    """Fetch the HTML content of the given URL."""
    response = requests.get(url)
    response.raise_for_status()
    return response.text


def parse_html(html):
    """Parse the HTML content and extract the required fields."""
    soup = BeautifulSoup(html, "html.parser")

    data = {
        "SL No": 1,  # This will be incremented for each bid
        "Start Date": soup.select_one('span[data-if-label="startDate"]').text.split()[
            0
        ],
        "Due by": soup.select_one('span[data-if-label="expireDate"]').text.split()[0],
        "Notice Type": "",  # Not available on website
        "Event": soup.select_one(
            'span[data-if-source="#RESP_AUC_H0B_WK_AUC_ID_BUS_UNIT"]'
        ).text,
        "Solicitation Title": soup.select_one('h3[data-if-label="eventName"]').text,
        "Agency": "",  # Not available on website
        "Category": "",  # Not available on website
        "Details": soup.select_one('div[data-if-label="descriptiondetails"]').text,
        "Contact Name": soup.select_one('strong[data-if-label="contactName"]').text,
        "Phone": soup.select_one('span[data-if-label="phoneText"]').text,
        "Email": soup.select_one('span[data-if-label="emffffailText"]').text,
        "Bid Detail Page URL": url,
        "See Attachments": (
            "Yes" if soup.select_one('button[data-if-ps-clickable="true"]') else "No"
        ),
    }

    return data


def save_to_excel(data, filename):
    """Save the extracted data to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Bids"

    headers = [
        "SL No",
        "Start Date",
        "Due by",
        "Notice Type",
        "Event",
        "Solicitation Title",
        "Agency",
        "Category",
        "Details",
        "Contact Name",
        "Phone",
        "Email",
        "Bid Detail Page URL",
        "See Attachments",
    ]

    ws.append(headers)

    for row in data:
        ws.append(
            [
                row["SL No"],
                row["Start Date"],
                row["Due by"],
                row["Notice Type"],
                row["Event"],
                row["Solicitation Title"],
                row["Agency"],
                row["Category"],
                row["Details"],
                row["Contact Name"],
                row["Phone"],
                row["Email"],
                row["Bid Detail Page URL"],
                row["See Attachments"],
            ]
        )

    wb.save(filename)


def get_bid_count(driver):
    """Get the number of bids from the search results page."""
    ul_element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//ul[@class='list-group']"))
    )
    li_elements = ul_element.find_elements(By.TAG_NAME, "li")
    return len(li_elements) - 1  # Subtract 1 to exclude the header li element


def open_bid_tabs(driver, bid_links):
    """Open each bid in a new tab."""
    for _, bid_url, _ in bid_links:
        driver.execute_script(f"window.open('{bid_url}', '_blank');")

    # Switch back to the first tab
    driver.switch_to.window(driver.window_handles[0])


def process_bid_tab(driver, tab_index, event_id):
    """Process a single bid tab."""
    driver.switch_to.window(driver.window_handles[tab_index])
    return extract_bid_details(driver, event_id, tab_index)


def move_remaining_files():
    """Move any remaining files to their respective bid folders."""
    print("Moving remaining files to their respective bid folders...")

    for filename in os.listdir(TEMP_DOWNLOAD_FOLDER):
        file_path = os.path.join(TEMP_DOWNLOAD_FOLDER, filename)

        # Skip directories and the Excel file
        if os.path.isdir(file_path) or filename.endswith(".xlsx"):
            continue

        # Extract the event ID from the filename (assuming it's at the start)
        event_id = filename.split("_")[0]

        # Construct the bid folder path
        bid_folder = os.path.join(MAIN_FOLDER, f"Event_{event_id}")

        # If the bid folder exists, move the file
        if os.path.exists(bid_folder):
            destination = os.path.join(bid_folder, filename)
            safe_move(file_path, destination)
            print(f"Moved {filename} to {bid_folder}")

            # Update Excel file with the new attachment
            update_excel_with_attachments(EXCEL_FILE, event_id, [filename])
        else:
            print(f"Could not find folder for {filename}")


def crop_images(images, left, right, top, bottom):
    """Crop all images in the list."""
    cropped_images = []
    for img in images:
        width, height = img.size
        cropped = img.crop((left, top, width - right, height - bottom))
        cropped_images.append(cropped)
    return cropped_images


def save_pdf_from_viewer(driver, event_id):
    """Save PDF from the viewer page by clicking the download button."""
    try:
        print(f"Attempting to save PDF for event {event_id}...")

        # JavaScript to click the download button
        download_button_js = """
        // Access the 'download' button within multiple shadow roots and click it
        const downloadButton = document
          .querySelector("#viewer")
          .shadowRoot
          .querySelector("#toolbar")
          .shadowRoot
          .querySelector("#downloads")
          .shadowRoot
          .querySelector("#download");

        // Click the button if it exists
        if (downloadButton) {
            downloadButton.click();
            return "Download button clicked successfully.";
        } else {
            return "Download button not found.";
        }
        """

        # Execute the JavaScript to click the download button
        result = driver.execute_script(download_button_js)
        print(result)

        # Wait for the download to complete
        wait_for_download(driver, 60)  # Wait up to 60 seconds for download

        # Move the downloaded file to the event folder
        event_folder = os.path.join(MAIN_FOLDER, f"Event_{event_id}")
        os.makedirs(event_folder, exist_ok=True)
        move_downloaded_files(driver, event_folder)

        print(f"PDF saved for event {event_id}")
        return True

    except Exception as e:
        print(f"Error saving PDF from viewer for event {event_id}: {str(e)}")
        print(f"Current URL: {driver.current_url}")
        return False


def cleanup_temp_folder():
    """Remove the temporary download folder safely."""
    try:
        if os.path.exists(TEMP_DOWNLOAD_FOLDER):
            # Try to remove all files first
            for file in glob.glob(os.path.join(TEMP_DOWNLOAD_FOLDER, "*")):
                try:
                    if os.path.isfile(file):
                        os.remove(file)
                    elif os.path.isdir(file):
                        shutil.rmtree(file)
                except Exception as e:
                    print(f"Error removing file {file}: {str(e)}")

            # Then try to remove the directory
            try:
                shutil.rmtree(TEMP_DOWNLOAD_FOLDER)
                print(f"Removed temporary download folder: {TEMP_DOWNLOAD_FOLDER}")
            except Exception as e:
                print(f"Error removing temp folder: {str(e)}")
    except Exception as e:
        print(f"Error in cleanup_temp_folder: {str(e)}")


# Register the cleanup function to run at exit
atexit.register(cleanup_temp_folder)


def final_cleanup():
    """Move any remaining files to their respective bid folders and update Excel."""
    print("Performing final cleanup...")

    if not os.path.exists(TEMP_DOWNLOAD_FOLDER):
        print(f"Temporary download folder not found: {TEMP_DOWNLOAD_FOLDER}")
        return

    for filename in os.listdir(TEMP_DOWNLOAD_FOLDER):
        file_path = os.path.join(TEMP_DOWNLOAD_FOLDER, filename)

        # Skip directories and the Excel file
        if os.path.isdir(file_path) or filename.endswith(".xlsx"):
            continue

        # Find the most recently created bid folder
        bid_folders = [f for f in os.listdir(MAIN_FOLDER) if f.startswith("Event_")]
        if not bid_folders:
            print("No bid folders found.")
            return

        latest_bid_folder = max(
            bid_folders, key=lambda f: os.path.getctime(os.path.join(MAIN_FOLDER, f))
        )
        bid_folder = os.path.join(MAIN_FOLDER, latest_bid_folder)
        event_id = latest_bid_folder.split("_")[1]

        try:
            if filename.lower().endswith(".zip"):
                # Unzip the file
                with zipfile.ZipFile(file_path, "r") as zip_ref:
                    zip_ref.extractall(bid_folder)
                print(f"Unzipped {filename} to {bid_folder}")
                os.remove(file_path)  # Remove the zip file after extraction
                moved_files = zip_ref.namelist()
            else:
                # Move non-zip files directly
                safe_move(file_path, os.path.join(bid_folder, filename))
                print(f"Moved {filename} to {bid_folder}")
                moved_files = [filename]

            # Update Excel file with new attachment names
            update_excel_with_attachments(EXCEL_FILE, event_id, moved_files)
        except Exception as e:
            print(f"Error processing file {filename}: {str(e)}")

    # After moving all files, attempt to remove the temporary folder
    try:
        shutil.rmtree(TEMP_DOWNLOAD_FOLDER)
        print(f"Removed temporary download folder: {TEMP_DOWNLOAD_FOLDER}")
    except Exception as e:
        print(f"Error removing temporary download folder: {str(e)}")


def mark_folder_as_completed():
    """Rename the script folder to indicate completion."""
    try:
        completed_folder = SCRIPT_FOLDER.replace("_IN_PROGRESS", "_COMPLETED")
        if os.path.exists(SCRIPT_FOLDER):
            os.rename(SCRIPT_FOLDER, completed_folder)
            logging.info(f"Marked folder as completed: {completed_folder}")
    except Exception as e:
        logging.error(f"Error marking folder as completed: {str(e)}")


def load_progress():
    """Load the list of scraped and skipped bid IDs from the progress file."""
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r") as f:
                data = json.load(f)
                return (
                    set(data.get("scraped_bids", [])),
                    set(data.get("skipped_bids", [])),  # Add skipped bids
                )
        except Exception as e:
            print(f"Error loading progress: {e}")
            return set(), set()
    return set(), set()


def save_progress(scraped_bids, skipped_bids, total_bids=None):
    """Save the list of scraped and skipped bid IDs to the progress file."""
    try:
        data = {
            "scraped_bids": list(scraped_bids),
            "skipped_bids": list(skipped_bids),  # Add skipped bids
            "total_bids": total_bids,
        }
        with open(PROGRESS_FILE, "w") as f:
            json.dump(data, f)
    except Exception as e:
        print(f"Error saving progress: {e}")


def cleanup_drivers():
    """Close only Chrome instances created by this script."""
    try:
        # Store the process ID of the current script
        script_pid = os.getpid()

        # Use wmic to get Chrome processes and their parent PIDs
        chrome_cmd = (
            "wmic process where \"name='chrome.exe'\" get ProcessId,ParentProcessId"
        )
        chromedriver_cmd = "wmic process where \"name='chromedriver.exe'\" get ProcessId,ParentProcessId"

        # Get process info
        chrome_output = (
            os.popen(chrome_cmd).read().strip().split("\n")[1:]
        )  # Skip header
        chromedriver_output = (
            os.popen(chromedriver_cmd).read().strip().split("\n")[1:]
        )  # Skip header

        # Function to kill process if it's child of our script
        def kill_if_child(pid, parent_pid):
            try:
                if int(parent_pid) == script_pid:
                    os.system(f"taskkill /F /PID {pid} 2>nul")
                    logging.info(f"Closed Chrome process {pid}")
            except:
                pass

        # Kill Chrome processes that are children of our script
        for line in chrome_output:
            if line.strip():
                try:
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        pid, parent_pid = parts
                        kill_if_child(pid, parent_pid)
                except:
                    continue

        # Kill ChromeDriver processes that are children of our script
        for line in chromedriver_output:
            if line.strip():
                try:
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        pid, parent_pid = parts
                        kill_if_child(pid, parent_pid)
                except:
                    continue

        time.sleep(2)  # Wait for processes to close
        logging.info("Cleaned up Chrome instances created by this script")

    except Exception as e:
        logging.error(f"Error cleaning up drivers: {e}")


def main():
    """Main function to execute the SF City Partner bid extraction process."""
    global scraped_bids, DAYS_TO_SEARCH

    # Parse command line arguments
    parser = argparse.ArgumentParser()
    parser.add_argument('--days', type=int, default=2, help='Number of days to search')
    args = parser.parse_args()
    DAYS_TO_SEARCH = args.days

    # Load previously scraped and skipped bids
    scraped_bids, skipped_bids = load_progress()
    logging.info(
        f"🔄 Loaded {len(scraped_bids)} scraped bids and {len(skipped_bids)} skipped bids"
    )

    try:
        # Force cleanup any existing Chrome instances
        cleanup_drivers()

        # Start fresh driver instance with random configuration
        driver = setup_driver()

        try:
            # Add random delay before first request
            time.sleep(random.uniform(2, 5))

            # Load the page
            driver.get("https://sfcitypartner.sfgov.org/pages/Events-BS3/event-search.aspx")

            # Add random delay before search
            time.sleep(random.uniform(1, 3))

            perform_advanced_search(driver)

            # Extract all AUC IDs
            auc_ids = extract_auc_ids(driver)
            if auc_ids is None:  # Changed from if not auc_ids
                logging.warning("No bids found")
                # Clean up before marking as completed
                try:
                    driver.quit()
                except:
                    pass
                cleanup_drivers()
                cleanup_temp_folder()
                # Mark folder as completed since there are no bids to process
                mark_folder_as_completed()
                return

            # Save total number of bids
            save_progress(scraped_bids, skipped_bids, len(auc_ids))

            # Find the first unprocessed bid
            next_bid = None
            for auc_id in auc_ids:
                if auc_id not in scraped_bids and auc_id not in skipped_bids:
                    next_bid = auc_id
                    break

            if next_bid is None:
                logging.info("✅ All bids have been processed!")
                mark_folder_as_completed()
                return

            # Process the next unprocessed bid
            bid_url = construct_bid_url(next_bid)
            driver.get(bid_url)

            # Check for lookup page
            if handle_lookup_page(driver, next_bid):
                skipped_bids.add(next_bid)
                save_progress(scraped_bids, skipped_bids, len(auc_ids))
                print(f"Moving on to next bid...")
                return

            # Extract bid details
            result = extract_bid_details(
                driver, next_bid, len(scraped_bids) + 1, len(auc_ids)
            )

            # Handle the result
            if result is None:  # Bid was skipped
                skipped_bids.add(next_bid)
            else:  # Bid was successfully scraped
                scraped_bids.add(next_bid)

            # Save progress
            save_progress(scraped_bids, skipped_bids, len(auc_ids))

            # Show remaining bids
            remaining = len(auc_ids) - len(scraped_bids) - len(skipped_bids)
            if remaining > 0:
                print(f"\n🔄 {remaining} bids remaining to process")
                print(f"Scraped: {len(scraped_bids)}, Skipped: {len(skipped_bids)}")
                print("Please run the script again for the next bid.")
            else:
                print("\n✅ All bids have been processed!")
                print(
                    f"Final count - Scraped: {len(scraped_bids)}, Skipped: {len(skipped_bids)}"
                )
                # Clean up before marking as completed
                try:
                    driver.quit()
                except:
                    pass
                cleanup_drivers()
                cleanup_temp_folder()
                mark_folder_as_completed()

        finally:
            # Clean up driver in all cases
            try:
                driver.quit()
            except:
                pass
            time.sleep(1)
            cleanup_drivers()
            cleanup_temp_folder()

    except Exception as e:
        error_message = f"An error occurred: {str(e)}"
        if 'driver' in locals():
            error_message += f"\nCurrent URL: {driver.current_url}"
        handle_error(error_message)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.exception("An unexpected error occurred:")
        handle_error(f"An unexpected error occurred: {str(e)}")
    finally:
        # Make absolutely sure we clean up
        cleanup_drivers()
