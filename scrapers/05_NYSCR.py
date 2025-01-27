import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import functools
import argparse

# Override the print function to always flush
print = functools.partial(print, flush=True)

import os

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import re
import time
import random
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
from dotenv import load_dotenv
import concurrent.futures
import requests
import shutil
from urllib.parse import urljoin
from fake_useragent import UserAgent
from selenium_stealth import stealth
import pickle
import zipfile
from utils.utils import safe_move, play_notification_sound
import json

# Load environment variables
load_dotenv()

# Constants
BASE_URL = "https://www.nyscr.ny.gov"
LOGIN_URL = f"{BASE_URL}/login.cfm?page=/business/adsOpenAll.cfm"
SEARCH_URL = f"{BASE_URL}/business/adsOpenAll.cfm"
MAX_BIDS = 30
SCRIPT_NAME = "05_NYSCR"
yesterday = datetime.now() - timedelta(days=1)
date_folder = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), yesterday.strftime("%Y-%m-%d")
)
main_folder = os.path.join(date_folder, f"{SCRIPT_NAME}_IN_PROGRESS")
temp_download_folder = os.path.join(main_folder, SCRIPT_NAME)

# Create all necessary directories
os.makedirs(date_folder, exist_ok=True)
os.makedirs(main_folder, exist_ok=True)
os.makedirs(temp_download_folder, exist_ok=True)

# Configure Chrome options
chrome_options = Options()
# chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-gpu")  # Disables GPU acceleration
chrome_options.add_argument("--disable-software-rasterizer")  # Use CPU rendering
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)

# Set default download folder to the script-specific folder
chrome_options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": temp_download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    },
)

# Rotate User Agents
ua = UserAgent()
chrome_options.add_argument(f"user-agent={ua.random}")

# Create a session for requests
session = requests.Session()


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description="Scrape NYSCR bids.")
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Number of days to look back for bids (default: 2)",
    )
    return parser.parse_args()


def random_wait(base_time, variation=0.2):
    """Wait for a random time around the base time."""
    wait_time = base_time + random.uniform(
        -base_time * variation, base_time * variation
    )
    time.sleep(wait_time)


def is_element_visible(driver, element):
    return element.is_displayed() and element.is_enabled()


def move_mouse_to_element(driver, element):
    action = ActionChains(driver)
    action.move_to_element(element)
    action.perform()
    random_wait(0.5, 0.3)


def random_scroll(driver):
    total_height = driver.execute_script("return document.body.scrollHeight")
    for i in range(1, total_height, random.randint(100, 200)):
        driver.execute_script(f"window.scrollTo(0, {i});")
        random_wait(0.1, 0.5)


def load_cookies(driver, path):
    if os.path.exists(path):
        with open(path, "rb") as cookiesfile:
            cookies = pickle.load(cookiesfile)
            for cookie in cookies:
                driver.add_cookie(cookie)


def save_cookies(driver, path):
    with open(path, "wb") as filehandler:
        pickle.dump(driver.get_cookies(), filehandler)


def login(driver, max_attempts=3):
    """Login to the NYSCR website."""
    for attempt in range(max_attempts):
        driver.get(LOGIN_URL)
        random_scroll(driver)

        try:
            WebDriverWait(driver, random.uniform(18, 22)).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
        except TimeoutException:
            print(
                f"Login page did not load within the expected time. Attempt {attempt + 1}/{max_attempts}"
            )
            continue

        username = driver.find_element(By.ID, "username")
        password = driver.find_element(By.ID, "password")
        submit_button = driver.find_element(By.NAME, "Submit")

        if not all(
            is_element_visible(driver, elem)
            for elem in [username, password, submit_button]
        ):
            print("Login elements are not visible. Retrying...")
            continue

        username.clear()
        password.clear()

        for char in os.getenv("NYSCR_EMAIL"):
            username.send_keys(char)
            random_wait(0.2, 0.5)

        random_wait(0.5, 0.3)

        for char in os.getenv("NYSCR_PASSWORD"):
            password.send_keys(char)
            random_wait(0.2, 0.5)

        random_wait(0.5, 0.3)
        move_mouse_to_element(driver, submit_button)
        submit_button.click()

        try:
            WebDriverWait(driver, random.uniform(28, 32)).until(
                lambda d: d.current_url.startswith(SEARCH_URL)
            )
            print("Successfully logged in.")
            return True
        except TimeoutException:
            print(f"Login failed. Current URL: {driver.current_url}")
            if (
                "unable to determine that you are not a robot"
                in driver.page_source.lower()
            ):
                print("CAPTCHA detected. Please solve the CAPTCHA manually.")
                input("Press Enter after solving the CAPTCHA...")
            else:
                print("Unexpected error. Retrying...")
                random_wait(5, 0.2)

    print("Login failed after maximum attempts. Exiting script.")
    return False


def sort_by_issue_date(driver):
    """Sort the bids by issue date (latest first)."""
    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "orderBy"))
    )
    for option in select_element.find_elements(By.TAG_NAME, "option"):
        if option.get_attribute("value") == "55":
            option.click()
            break
    time.sleep(2)  # Wait for the page to reload


def get_bid_links(driver, max_bids=MAX_BIDS, days_back=2):
    """Get bid links from the search page."""
    bid_links = []
    page_number = 1
    cutoff_date = datetime.now() - timedelta(days=days_back)

    while len(bid_links) < max_bids:
        print(f"Current Page Number: {page_number}")
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.r1, tr.r2")
        for row in rows:
            issue_date_elem = row.find_element(
                By.XPATH,
                ".//div[contains(@class, 'labelText') and text()='Issue Date:']/following-sibling::div",
            )
            issue_date = datetime.strptime(issue_date_elem.text, "%m/%d/%Y")
            if issue_date < cutoff_date:
                print(
                    f"Bids with Posting Range within {days_back} Days Successfully Extracted"
                )
                return bid_links

            link_elem = row.find_element(
                By.XPATH, ".//div[@class='result_btn tblColor']/a[@class='view']"
            )
            bid_links.append(link_elem.get_attribute("href"))

            if len(bid_links) >= max_bids:
                print(
                    f"Bids with Posting Range within {days_back} Days Successfully Extracted"
                )
                return bid_links

        try:
            next_button = driver.find_element(
                By.XPATH, "//div[@class='paging']/a[text()='Next »']"
            )
            next_button.click()
            time.sleep(2)
            page_number += 1
        except NoSuchElementException:
            print("Reached the last page")
            break

    print(f"Bids with Posting Range within {days_back} Days Successfully Extracted")
    return bid_links


def extract_bid_details(driver, url):
    """Extract bid details from a given URL."""
    driver.get(url)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "adSummary"))
    )

    details = {}

    def safe_find_element(xpath, default="Not available"):
        try:
            return driver.find_element(By.XPATH, xpath).text.strip()
        except NoSuchElementException:
            return default

    details["SL No"] = ""  # This will be filled in the main loop
    details["Date of Issue"] = safe_find_element(
        "//ul/li/span[text()='Date of Issue:']/../../following-sibling::p[1]"
    )
    details["Due Date/Time"] = safe_find_element(
        "//ul/li/span[text()='Due Date/Time:']/../../following-sibling::p[1]"
    ).split()[
        0
    ]  # Only extract the date
    details["Opportunity Type"] = safe_find_element(
        "//ul/li/span[text()='Opportunity Type:']/../../following-sibling::p[1]"
    )
    details["CR Number"] = safe_find_element(
        "//ul/li/span[text()='CR Number:']/../../following-sibling::p[1]"
    )
    details["Title"] = safe_find_element(
        "//ul/li[@class='firstPub']/span[text()='Title:']/../../following-sibling::p[1]"
    )
    details["Agency"] = safe_find_element(
        "//ul/li/span[text()='Agency:']/../../following-sibling::p[1]"
    )
    details["Classification(s)"] = safe_find_element(
        "//ul/li/span[text()='Classification(s):']/../../following-sibling::p[1]"
    )

    description_elements = driver.find_elements(By.XPATH, "//div[@id='Description']/p")
    details["Description"] = "\n".join(
        [
            p.text
            for p in description_elements
            if not p.find_elements(By.TAG_NAME, "strong")
        ]
    )

    details["Additional Summary"] = (
        ""  # Not available on website, but required in excel
    )
    details["Location"] = safe_find_element(
        "//ul/li/span[text()='Location:']/../../following-sibling::p[1]"
    )

    try:
        contact_tab = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "ui-id-2"))
        )
        contact_tab.click()
        contact_info = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "addContentReview"))
        )
        details["Primary contact"] = contact_info.text
    except TimeoutException:
        details["Primary contact"] = "Not available"

    details["Bid Detail Page URL"] = url
    details["Attachments"] = ""  # This will be filled after downloading attachments

    return details


def wait_for_download(download_folder):
    """Wait indefinitely until a file appears in the download folder and is complete."""
    while True:
        # Check for any .crdownload or .tmp files
        temp_files = [
            f
            for f in os.listdir(download_folder)
            if f.endswith((".crdownload", ".tmp"))
        ]
        if temp_files:
            time.sleep(2)
            continue

        # Look for completed downloads
        files = [
            f
            for f in os.listdir(download_folder)
            if not f.endswith((".crdownload", ".tmp"))
        ]
        if files:
            # Get the most recently modified file
            newest_file = max(
                [os.path.join(download_folder, f) for f in files], key=os.path.getmtime
            )

            # Verify file is completely written
            initial_size = os.path.getsize(newest_file)
            time.sleep(2)
            if os.path.getsize(newest_file) == initial_size:
                return newest_file

        time.sleep(1)


def download_attachment(driver, link, download_folder):
    """Download an attachment and wait until it's complete."""
    try:
        # Clear the download folder before each download
        for file in os.listdir(download_folder):
            file_path = os.path.join(download_folder, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Error clearing file {file}: {e}")

        # Click the download link
        actions = ActionChains(driver)
        actions.move_to_element(link).click().perform()

        # Wait indefinitely for the download to complete
        print("Waiting for download to complete...")
        downloaded_file = wait_for_download(download_folder)

        if downloaded_file:
            print(f"Download completed: {os.path.basename(downloaded_file)}")
            return {
                "name": os.path.basename(downloaded_file),
                "path": downloaded_file,
                "size": os.path.getsize(downloaded_file),
            }

    except Exception as e:
        print(f"Error downloading attachment: {str(e)}")
    return None


def move_file_with_retry(source, destination, max_retries=5, delay=10):
    """Move a file with retry mechanism."""
    for attempt in range(max_retries):
        try:
            if os.path.exists(source):
                safe_move(source, destination)
                print(
                    f"Successfully moved {os.path.basename(source)} to {os.path.dirname(destination)}"
                )
                return True
            else:
                print(f"Source file {source} not found. Retrying in {delay} seconds...")
                time.sleep(delay)
        except Exception as e:
            print(
                f"Error moving file {source} to {destination}: {str(e)}. Retrying in {delay} seconds..."
            )
            time.sleep(delay)
    print(f"Failed to move {source} after {max_retries} attempts.")
    return False


def get_attachments(driver, download_folder, cr_number, main_folder):
    """Get attachment links for a bid and download them."""
    try:
        # Wait for the attachments tab to be present and click it
        print(f"Looking for attachments for bid {cr_number}")

        # Try different selectors for the attachments tab
        tab_selectors = [
            "//a[@id='ui-id-3']",
            "//li[@aria-controls='Documents']/a",
            "//a[contains(text(), 'Documents')]",
        ]

        attachments_tab = None
        for selector in tab_selectors:
            try:
                attachments_tab = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                break
            except TimeoutException:
                continue

        if not attachments_tab:
            print("Could not find attachments tab")
            return []

        # Try different methods to click the tab
        try:
            # Try regular click
            attachments_tab.click()
        except:
            try:
                # Try JavaScript click
                driver.execute_script("arguments[0].click();", attachments_tab)
            except:
                # Try ActionChains
                actions = ActionChains(driver)
                actions.move_to_element(attachments_tab).click().perform()

        time.sleep(3)  # Wait for tab content to load

        # Check for "no documents" message
        try:
            no_docs_element = driver.find_element(
                By.XPATH,
                "//p[contains(text(), 'No documents have been uploaded to this ad')]",
            )
            if no_docs_element:
                print("No attachments are present for this bid")
                return []
        except NoSuchElementException:
            pass  # Continue if no such message is found

        # Wait for the documents table
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "Documents"))
            )
        except TimeoutException:
            print("No attachments table found")
            return []

        # Try different selectors for attachment links
        attachment_selectors = [
            "//table[@class='categoriesTable']//tr/td[1]/a[contains(@href, 'adsViewPDF.cfm')]",
            "//div[@id='Documents']//table//tr/td[1]/a[contains(@href, 'adsViewPDF.cfm')]",
            "#Documents table tr td:first-child a[href*='adsViewPDF.cfm']",
        ]

        attachment_links = []
        for selector in attachment_selectors:
            try:
                if selector.startswith("/"):
                    attachment_links = driver.find_elements(By.XPATH, selector)
                else:
                    attachment_links = driver.find_elements(By.CSS_SELECTOR, selector)
                if attachment_links:
                    break
            except:
                continue

        if not attachment_links:
            print("No attachments found for this bid.")
            return []

        print(f"Found {len(attachment_links)} attachment links.")

        # Create bid folder if attachments are found
        bid_folder = os.path.join(main_folder, cr_number)
        os.makedirs(bid_folder, exist_ok=True)

        downloaded_files = []

        # Process attachments one at a time
        for i, link in enumerate(attachment_links, 1):
            print(
                f"Processing attachment {i}/{len(attachment_links)} for bid {cr_number}"
            )

            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    print(f"Download attempt {attempt + 1} for attachment {i}")

                    # Scroll the link into view
                    driver.execute_script("arguments[0].scrollIntoView(true);", link)
                    time.sleep(1)

                    # Try different click methods
                    try:
                        # Try ActionChains with move and click
                        actions = ActionChains(driver)
                        actions.move_to_element(link)
                        actions.click()
                        actions.perform()
                    except:
                        try:
                            # Try JavaScript click
                            driver.execute_script("arguments[0].click();", link)
                        except:
                            # Try direct click
                            link.click()

                    time.sleep(3)  # Wait for download to start

                    result = download_attachment(driver, link, download_folder)

                    if result:
                        source = result["path"]
                        destination = os.path.join(bid_folder, result["name"])

                        # Wait until file exists and is complete
                        print(
                            f"Waiting for file to be ready for moving: {result['name']}"
                        )
                        while True:
                            if os.path.exists(source) and os.path.getsize(source) > 0:
                                if move_file_with_retry(source, destination):
                                    downloaded_files.append(result["name"])
                                    print(
                                        f"Successfully downloaded and moved attachment {i}"
                                    )
                                    break
                            time.sleep(1)
                        break  # Break attempt loop if successful

                    if attempt < max_attempts - 1:
                        print(f"Retrying download for attachment {i}")
                        time.sleep(5)

                except Exception as e:
                    print(
                        f"Error on attempt {attempt + 1} for attachment {i}: {str(e)}"
                    )
                    if attempt < max_attempts - 1:
                        time.sleep(5)

            # Wait between attachments
            time.sleep(5)

        print(
            f"Successfully downloaded {len(downloaded_files)} out of {len(attachment_links)} attachments for bid {cr_number}"
        )
        return downloaded_files

    except Exception as e:
        print(f"Error accessing attachments tab or downloading files: {str(e)}")
        return []


def move_remaining_files(download_folder, main_folder):
    """Move any leftover attachments to their respective bid folders."""
    for filename in os.listdir(download_folder):
        source = os.path.join(download_folder, filename)
        if os.path.isfile(source):
            moved = False
            for bid_folder in os.listdir(main_folder):
                bid_folder_path = os.path.join(main_folder, bid_folder)
                if os.path.isdir(bid_folder_path) and not bid_folder.endswith(
                    "_IN_PROGRESS"
                ):
                    destination = os.path.join(bid_folder_path, filename)
                    try:
                        if filename.lower().endswith((".zip", ".rar", ".7z")):
                            # Unzip compressed files
                            with zipfile.ZipFile(source, "r") as zip_ref:
                                zip_ref.extractall(bid_folder_path)
                            print(
                                f"Extracted compressed file {filename} to {bid_folder}"
                            )
                            os.remove(source)  # Remove the original compressed file
                            moved = True
                        else:
                            # Move other files using move_file_with_retry
                            if move_file_with_retry(source, destination):
                                print(f"Moved leftover file {filename} to {bid_folder}")
                                moved = True
                        if moved:
                            break
                    except Exception as e:
                        print(
                            f"Error processing file {filename} for {bid_folder}: {str(e)}"
                        )
            if not moved:
                print(f"Could not find a suitable bid folder for {filename}")


def select_ad_types(driver):
    """Select specific ad types before extracting bids."""
    print("Selecting ad types...")

    # Dictionary of ad types to select with their values
    ad_types = {
        "type1": "0",  # General
        "type2": "10",  # Discretionary under $50,000
        "type3": "20",  # Discretionary between $50,000 and $750,000
        "type6": "50",  # Requests for information
        "type7": "60",  # Grants/Notice of funds
        "type8": "100",  # Contractor Ads
    }

    try:
        # Wait for the filter form to be present
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "adsFilterTypes"))
        )

        # First uncheck all existing selections
        checkboxes = driver.find_elements(By.CSS_SELECTOR, "input[name='adType']")
        for checkbox in checkboxes:
            if checkbox.is_selected():
                checkbox.click()
                time.sleep(1)  # Wait for the form to update

        # Now select our desired ad types
        for type_id, value in ad_types.items():
            try:
                checkbox = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, f"input[name='adType'][value='{value}']")
                    )
                )

                if not checkbox.is_selected():
                    # Scroll the checkbox into view
                    driver.execute_script(
                        "arguments[0].scrollIntoView(true);", checkbox
                    )
                    time.sleep(0.5)

                    # Click the checkbox
                    checkbox.click()
                    print(f"Selected ad type: {type_id} (value: {value})")

                    # Wait for the page to update after each selection
                    time.sleep(2)

            except Exception as e:
                print(f"Error selecting ad type {type_id}: {str(e)}")

        print("Ad types selection completed")

        # Wait for the page to fully update after all selections
        time.sleep(5)

    except Exception as e:
        print(f"Error in select_ad_types: {str(e)}")
        raise


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
        issue_date = datetime.strptime(
            bid_details.get("Date of Issue", ""), "%m/%d/%Y"
        ).strftime("%Y-%m-%d")
    except ValueError:
        issue_date = bid_details.get("Date of Issue", "")

    # Add/update the bid in cache
    cache[bid_url] = {
        "posted_date": issue_date,
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
        print(f"⏭️ Skipping previously processed bid from {bid_data['posted_date']}")
        return False
    return True


def main():
    # Parse command line arguments
    args = parse_arguments()
    days_back = args.days

    print(f"Bids Extraction Started (looking back {days_back} days)")

    # Create date folder and subfolder with IN_PROGRESS suffix
    yesterday = datetime.now() - timedelta(days=1)
    date_folder = os.path.join(
        os.path.dirname(os.path.dirname(__file__)), yesterday.strftime("%Y-%m-%d")
    )
    main_folder = os.path.join(date_folder, f"{SCRIPT_NAME}_IN_PROGRESS")
    temp_download_folder = os.path.join(main_folder, SCRIPT_NAME)

    # Create all necessary directories
    os.makedirs(date_folder, exist_ok=True)
    os.makedirs(main_folder, exist_ok=True)
    os.makedirs(temp_download_folder, exist_ok=True)
    print(f"📁 Saving data in folder: {main_folder}")

    # Update the download folder path
    global SCRIPT_DOWNLOAD_FOLDER
    SCRIPT_DOWNLOAD_FOLDER = temp_download_folder

    driver = webdriver.Chrome(options=chrome_options)

    stealth(
        driver,
        languages=["en-US", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
    )

    try:
        if not login(driver):
            print("Login failed. Exiting script.")
            return

        select_ad_types(driver)
        sort_by_issue_date(driver)

        bid_links = get_bid_links(driver, days_back=days_back)
        if not bid_links:
            print(f"No bids found within the last {days_back} days.")
            return

        print(f"Found {len(bid_links)} bids within the last {days_back} days.")

        data = []
        for i, url in enumerate(bid_links, 1):
            print(f"Processing bid {i}/{len(bid_links)}")
            try:
                # Add cache check here
                if not should_process_bid_link(url):
                    continue

                details = extract_bid_details(driver, url)
                details["SL No"] = i
                cr_number = details.get("CR Number", f"unknown_bid_{i}")

                # Get attachments without creating folder initially
                attachment_names = get_attachments(
                    driver, SCRIPT_DOWNLOAD_FOLDER, cr_number, main_folder
                )
                print(f"Downloading {len(attachment_names)} attachments for bid {i}")

                # Add additional wait time after processing each bid
                time.sleep(10)

                # Update the Attachments column with file names
                details["Attachments"] = (
                    ", ".join(attachment_names) if attachment_names else ""
                )
                data.append(details)

                # Update Excel file after each bid
                df = pd.DataFrame(data)
                # Map the old column names to the new standardized column names
                column_mapping = {
                    "SL No": "SL No",
                    "Date of Issue": "Posted Date",
                    "Due Date/Time": "Response Date",
                    "Opportunity Type": "Notice Type",
                    "CR Number": "Solicitation Number",
                    "Title": "Solicitation Title",
                    "Agency": "Agency",
                    "Classification(s)": "Category",
                    "Description": "Description",
                    "Additional Summary": "Additional Summary, if any",
                    "Location": "Contracting Office Address",
                    "Primary contact": "Contact Information",
                    "Bid Detail Page URL": "Bid Detail Page URL",
                    "Attachments": "Attachments"
                }
                
                # Rename the columns
                df = df.rename(columns=column_mapping)
                
                # Ensure the correct order of columns
                columns_order = [
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
                    "Attachments"
                ]
                df = df[columns_order]
                df["Posted Date"] = pd.to_datetime(
                    df["Posted Date"], format="%m/%d/%Y", errors="coerce"
                ).dt.strftime("%Y-%m-%d")
                df["Response Date"] = pd.to_datetime(
                    df["Response Date"], format="%m/%d/%Y", errors="coerce"
                ).dt.strftime("%Y-%m-%d")

                # Save Excel file in the main folder
                excel_file = os.path.join(main_folder, f"{SCRIPT_NAME}.xlsx")
                df.to_excel(excel_file, index=False)

                # Save to cache after successful processing
                save_to_cache(url, details)

                # Rewrite the Attachments column based on the actual files in the bid folder
                if os.path.exists(os.path.join(main_folder, cr_number)):
                    actual_attachments = [
                        f
                        for f in os.listdir(os.path.join(main_folder, cr_number))
                        if os.path.isfile(os.path.join(main_folder, cr_number, f))
                    ]
                    df.loc[df["Solicitation Number"] == cr_number, "Attachments"] = (
                        ", ".join(actual_attachments)
                        if actual_attachments
                        else "No attachments"
                    )
                    df.to_excel(excel_file, index=False)

            except Exception as e:
                print(f"Error processing bid {i}: {str(e)}")
                play_notification_sound()
                input("Press Enter to continue...")
                continue

        # Add additional wait time before moving remaining files
        time.sleep(30)

        # Move any remaining files
        move_remaining_files(SCRIPT_DOWNLOAD_FOLDER, main_folder)

        print("All Bids and Attachments Extraction Successfully Completed")
        play_notification_sound()

        # Clean up temporary download folder before renaming
        try:
            if os.path.exists(temp_download_folder):
                shutil.rmtree(temp_download_folder)
                print("✅ Removed temporary download folder")
        except Exception as e:
            print(f"⚠️ Error removing temporary folder: {str(e)}")

        # When renaming the completed folder, maintain the same structure
        completed_folder = os.path.join(date_folder, f"{SCRIPT_NAME}_COMPLETED")
        if os.path.exists(completed_folder):
            shutil.rmtree(completed_folder)
        os.rename(main_folder, completed_folder)
        print(f"Renamed folder to: {completed_folder}")

    finally:
        # Save cookies before quitting
        cookie_file = "cookies.pkl"
        if os.path.exists(cookie_file):
            try:
                os.remove(cookie_file)
                print("🗑️ Removed cookies file")
            except Exception as e:
                print(f"Error removing cookies file: {e}")
                
        driver.quit()
        # Clean up the temporary download folder
        shutil.rmtree(temp_download_folder, ignore_errors=True)


if __name__ == "__main__":
    main()
