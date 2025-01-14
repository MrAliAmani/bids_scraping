import sys
import io
import functools

# Override the print function to always flush
print = functools.partial(print, flush=True)

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import os

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.utils import safe_move, play_notification_sound
import time
from datetime import datetime, timedelta
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import re
from dotenv import load_dotenv
import math
from fake_useragent import UserAgent
import shutil
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import urllib3
from urllib3.exceptions import InsecureRequestWarning
import glob
import threading
import queue
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from dateutil import parser
import numpy as np
import winsound  # For playing notification sounds
import sys  # For pausing the script
import mimetypes
import traceback

# Disable SSL warnings (use with caution in production)
urllib3.disable_warnings(InsecureRequestWarning)

# Load environment variables
load_dotenv()

# Create a directory named after yesterday's date
yesterday = datetime.now() - timedelta(days=1)
main_folder = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), yesterday.strftime("%Y-%m-%d")
)
os.makedirs(main_folder, exist_ok=True)

# Create a script-specific folder with IN_PROGRESS suffix
script_folder = os.path.join(main_folder, "02_NYC_IN_PROGRESS")
os.makedirs(script_folder, exist_ok=True)

# Create a script-specific download folder inside the script folder
download_folder = os.path.join(script_folder, "02_NYC")
os.makedirs(download_folder, exist_ok=True)

# Update the excel_path to be in the script folder with IN_PROGRESS suffix
excel_path = os.path.join(script_folder, "02_NYC.xlsx")


def setup_driver():
    """
    Set up and configure the Chrome WebDriver for NYC procurement website.

    Returns:
        webdriver.Chrome: Configured Chrome WebDriver instance.
    """
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")  # Run in headless mode
    chrome_options.add_argument("--disable-gpu")  # Disables GPU acceleration
    chrome_options.add_argument("--disable-software-rasterizer")  # Use CPU rendering)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")  # Set a default window size
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--ignore-ssl-errors=yes")
    chrome_options.add_argument("--ignore-certificate-errors")

    # Set download preferences to use the script-specific folder
    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,  # Disable safe browsing download protection
        "profile.default_content_setting_values.automatic_downloads": 1,  # Allow automatic downloads
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # Disable the download restrictions
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    # Set up capabilities to ignore SSL errors
    chrome_options.set_capability("acceptInsecureCerts", True)

    # Use webdriver_manager to automatically download and manage the ChromeDriver
    service = Service(ChromeDriverManager().install())

    # Set a specific desktop user agent
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    )

    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver


def setup_session():
    """
    Set up a requests session with retry mechanism and random user agent.

    Returns:
        requests.Session: Configured session object.
    """
    session = requests.Session()
    retry = Retry(total=5, backoff_factor=0.1, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    ua = UserAgent()
    session.headers.update({"User-Agent": ua.random})
    return session


def random_sleep(min_seconds, max_seconds):
    """
    Sleep for a random duration between min_seconds and max_seconds.

    This function is used to introduce random delays between actions to mimic human behavior
    and avoid detection as a bot.

    Args:
        min_seconds (float): Minimum sleep duration in seconds.
        max_seconds (float): Maximum sleep duration in seconds.
    """
    time.sleep(random.uniform(min_seconds, max_seconds))


def login(driver, email, password):
    print("Attempting to log in...")
    driver.get("https://a856-cityrecord.nyc.gov/Search/Advanced")
    random_sleep(2, 4)

    try:
        # Click the login link
        login_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@href, '/Visitor/LogIn')]")
            )
        )
        login_link.click()
        random_sleep(5, 7)  # Increased wait time after clicking login link

        # Wait for the login form to be present and visible
        WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, "gigya-loginID"))
        )

        # Find the email field
        print("Locating email field...")
        email_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "gigya-loginID"))
        )
        email_field.clear()
        email_field.send_keys(email)
        print("Email entered")
        random_sleep(2, 3)

        # Find the password field
        print("Locating password field...")
        password_field = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "gigya-password"))
        )
        password_field.clear()
        password_field.send_keys(password)
        print("Password entered")
        random_sleep(2, 3)

        # Click the login button
        print("Clicking login button...")
        login_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//input[@type='submit' and @value='Login']")
            )
        )
        login_button.click()
        print("Login button clicked")

        # Wait for the page to load after login
        WebDriverWait(driver, 60).until(
            EC.url_contains("cityrecord.nyc.gov/Notification")
        )

        print("Login successful")
        return True

    except Exception as e:
        print(f"Login failed: {str(e)}")
        print(f"Current URL: {driver.current_url}")
        driver.save_screenshot("login_failure.png")
        return False


def navigate_to_advanced_search(driver):
    """
    Navigate to the Advanced Search page on the NYC procurement website.

    This function clicks the "Advanced" button and waits for the Advanced Search page to load.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.

    Returns:
        bool: True if navigation is successful, False otherwise.
    """
    print("Navigating to Advanced Search page...")
    try:
        # Ensure we're on the Notification page
        WebDriverWait(driver, 30).until(
            EC.url_to_be("https://a856-cityrecord.nyc.gov/Notification")
        )

        # Wait for the "Advanced" button to be clickable
        advanced_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "advancedSearch"))
        )
        advanced_button.click()

        # Wait for the Advanced Search page to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "ddlSection"))
        )
        print("Successfully navigated to Advanced Search page")
        return True
    except Exception as e:
        print(f"Failed to navigate to Advanced Search page: {str(e)}")
        print(f"Current URL: {driver.current_url}")
        driver.save_screenshot("navigation_failure.png")
        return False


def apply_filters(driver):
    """
    Apply filters on the Advanced Search page to narrow down the search results.

    This function sets various filters such as section, notice type, date range, etc.
    to refine the search results for procurement bids.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.

    Raises:
        Exception: If an error occurs while applying filters.
    """
    try:
        if driver.current_url != "https://a856-cityrecord.nyc.gov/Search/Advanced":
            print("Not on the Advanced Search page. Redirecting...")
            if not navigate_to_advanced_search(driver):
                raise Exception("Failed to navigate to Advanced Search page")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "ddlSection"))
        )

        # Select "Procurement" from the section dropdown
        section_dropdown = Select(driver.find_element(By.ID, "ddlSection"))
        section_dropdown.select_by_value("6")  # Value for "Procurement"
        random_sleep(1, 2)

        # Select "Solicitation" from the notice type dropdown
        notice_type_dropdown = Select(driver.find_element(By.ID, "ddlNoticeType"))
        notice_type_dropdown.select_by_value("1")  # Value for "Solicitation"
        random_sleep(1, 2)

        # Click the checkbox for current ads
        checkbox_label = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//label[@for='chkCurrentAds']"))
        )
        checkbox_label.click()
        random_sleep(1, 2)

        # Set the date range for the last 2 days
        today = datetime.now()
        two_days_ago = today - timedelta(days=2)

        start_date_input = driver.find_element(By.ID, "txtStartDate")
        start_date_input.clear()
        start_date_input.send_keys(two_days_ago.strftime("%m/%d/%Y"))

        end_date_input = driver.find_element(By.ID, "txtEndDate")
        end_date_input.clear()
        end_date_input.send_keys(today.strftime("%m/%d/%Y"))

        random_sleep(1, 2)

        # Submit the form
        submit_button = driver.find_element(By.ID, "AdvancedSubmitButton")
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
        submit_button.click()

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "notice-container"))
        )

        print("Filters applied successfully")
    except Exception as e:
        print(f"Error applying filters: {str(e)}")
        driver.save_screenshot("apply_filters_error.png")
        raise


def extract_bid_data(driver, max_bids=30):
    """
    Extract bid data from the search results page.

    This function navigates through the search results pages and extracts bid information
    up to the specified maximum number of bids.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        max_bids (int): Maximum number of bids to extract.

    Returns:
        list: A list of BeautifulSoup objects representing the extracted bid data.
    """
    all_bids = []
    page = 1

    while len(all_bids) < max_bids:
        print(f"📄 Current Page Number: {page}")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "notice-container"))
        )

        soup = BeautifulSoup(driver.page_source, "html.parser")
        bids = soup.find_all("div", class_="notice-container")

        if not bids:
            print("No bids found on this page. Ending extraction.")
            break

        for bid in bids:
            all_bids.append(bid)
            if len(all_bids) >= max_bids:
                break

        if len(all_bids) >= max_bids:
            break

        next_page = soup.find("a", class_="page-link next")
        if not next_page:
            break

        next_page_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@class, 'page-link next')]")
            )
        )

        driver.execute_script("arguments[0].scrollIntoView(true);", next_page_link)
        time.sleep(2)

        try:
            overlay = driver.find_element(By.CLASS_NAME, "overlay")
            driver.execute_script("arguments[0].style.display='none';", overlay)
        except NoSuchElementException:
            pass

        driver.execute_script("arguments[0].click();", next_page_link)

        page += 1
        random_sleep(2, 4)

    print(f"✅ Bids Extracted: {len(all_bids)} bids")
    return all_bids


def check_available_date(driver, bid_link):
    """
    Check if the bid has an "Available Date" on its detail page.

    This function navigates to the bid detail page and checks for the presence of an "Available Date" field.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        bid_link (str): The URL of the bid detail page.

    Returns:
        tuple: A tuple containing a boolean indicating whether the "Available Date" is found,
               and the "Available Date" string if found, or None otherwise.
    """
    driver.get("https://a856-cityrecord.nyc.gov" + bid_link)
    random_sleep(2, 4)

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "portlet-body"))
    )

    soup = BeautifulSoup(driver.page_source, "html.parser")
    available_date_label = soup.find("label", string="Available Date")

    if available_date_label:
        available_date = available_date_label.find_next(
            "div", class_="form-control-static"
        ).text.strip()
        return True, available_date
    return False, None


def extract_bid_details(driver, bid, index):
    """
    Extract details from a specific bid page.

    This function parses the bid detail page and extracts various pieces of information
    such as publication date, due date, agency name, etc.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        bid (BeautifulSoup): The BeautifulSoup object representing the bid data.
        index (int): The index of the bid in the overall extraction process.

    Returns:
        dict: A dictionary containing the extracted bid details.
    """
    details = {
        "SL No": index,
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

    soup = BeautifulSoup(driver.page_source, "html.parser")

    form_body = soup.find("div", class_="portlet-body form")
    if form_body:
        divs = form_body.find_all("div", class_="form-control form-control-static")
        labels = form_body.find_all("label")

        for div, label in zip(divs, labels):
            text = div.text.strip()
            label_text = label.text.strip()

            if label_text == "Publication Date":
                details["Posted Date"] = text
            elif label_text == "Selection Method":
                details["Notice Type"] = text
            elif label_text == "Agency Name":
                details["Agency"] = text
            elif label_text == "Category":
                details["Category"] = text
            elif label_text == "Request Address":
                details["Contracting Office Address"] = text
            elif label_text == "Contact Information":
                details["Contact Information"] = text

    pin_due_date_div = soup.find("label", string="PIN - Due Date")
    if pin_due_date_div:
        pin_due_date_text = pin_due_date_div.find_previous(
            "div", class_="form-control-static"
        ).text.strip()
        pin_match = re.search(r"PIN#([\w-]+)", pin_due_date_text)
        due_date_match = re.search(
            r"Due:\s*([\d/]+\s*[\d:]+\s*[APM]+)", pin_due_date_text
        )

        if pin_match:
            details["Solicitation Number"] = pin_match.group(
                1
            )  # Extract only the number after PIN#
        if due_date_match:
            details["Response Date"] = due_date_match.group(1)

    title_span = soup.find("span", class_="caption-subject bold")
    if title_span:
        details["Solicitation Title"] = title_span.text.strip()

    description_div = soup.find(
        "span", string=" Description", class_="caption-subject bold uppercase"
    )
    if description_div:
        description_content = description_div.find_parent(
            "div", class_="portlet light"
        ).find("div", class_="col-md-12")
        if description_content:
            details["Description"] = " ".join(
                [p.text.strip() for p in description_content.find_all("p")]
            )

    attachments_section = soup.find(
        "span", string=" Attachments", class_="caption-subject bold uppercase"
    )
    if attachments_section:
        external_link = soup.find(
            "a", href=lambda x: x and "passport.cityofnewyork.us" in x
        )
        if external_link:
            details["Attachments"] = (
                "Attachments Available"  # We'll handle this in download_attachments
            )
        else:
            details["Attachments"] = ""
    else:
        details["Attachments"] = ""

    # Parse dates
    for date_field in ["Posted Date", "Response Date"]:
        if details[date_field]:
            try:
                parsed_date = parse_date(details[date_field])
                details[date_field] = parsed_date.strftime("%Y-%m-%d")
            except ValueError as e:
                print(f"Error parsing {date_field}: {e}")
                details[date_field] = "Invalid Date"

    return details


def update_excel(data, excel_path):
    """
    Update the Excel file with new bid data.

    This function takes the extracted bid details and updates the Excel file with the new data.
    If the Excel file already exists, it appends the new data to the existing file.
    If the Excel file doesn't exist, it creates a new one.

    Args:
        data (dict): The bid details to be added to the Excel file.
        excel_path (str): The path to the Excel file.
    """
    columns = [
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

    if os.path.exists(excel_path):
        existing_df = pd.read_excel(excel_path)
        new_sl_no = len(existing_df) + 1
        data["SL No"] = new_sl_no
        new_df = pd.DataFrame([data])
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        data["SL No"] = 1
        updated_df = pd.DataFrame([data])

    # Ensure all columns are present and in the correct order
    for col in columns:
        if col not in updated_df.columns:
            updated_df[col] = ""
    updated_df = updated_df[columns]

    updated_df.to_excel(excel_path, index=False)
    print(f"Excel file updated: {excel_path}")


def wait_for_download(download_folder, timeout=600):
    """
    Wait for a file to finish downloading by checking for .crdownload files.

    Args:
        download_folder (str): The path to the download folder.
        timeout (int): The maximum time to wait for the file to download in seconds.

    Returns:
        str: The name of the downloaded file if successful, None otherwise.
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        crdownload_files = [
            f for f in os.listdir(download_folder) if f.endswith(".crdownload")
        ]
        if not crdownload_files:
            # Check for new files that are not .crdownload
            files = [
                f for f in os.listdir(download_folder) if not f.endswith(".crdownload")
            ]
            if files:
                return files[0]  # Return the name of the first non-.crdownload file
        time.sleep(1)
    return None


def download_and_move_file(download_queue, download_folder, target_folder):
    while True:
        task = download_queue.get()
        if task is None:
            break  # Exit signal

        file_name, file_url = task
        try:
            downloaded_file = wait_for_download(download_folder)
            if downloaded_file:
                source_path = os.path.join(download_folder, downloaded_file)
                target_path = os.path.join(target_folder, file_name)
                safe_move(source_path, target_path)
                print(f"Moved file: {file_name}")
            else:
                print(f"Failed to download {file_name}")
        except Exception as e:
            print(f"Error processing {file_name}: {str(e)}")
        finally:
            download_queue.task_done()


def move_remaining_files(download_folder, output_dir):
    remaining_files = [
        f for f in os.listdir(download_folder) if not f.endswith(".crdownload")
    ]
    for file_name in remaining_files:
        source_path = os.path.join(download_folder, file_name)
        for bid_folder in os.listdir(output_dir):
            if bid_folder in file_name:
                target_folder = os.path.join(output_dir, bid_folder)
                target_path = os.path.join(target_folder, file_name)
                safe_move(source_path, target_path)
                print(f"Moved remaining file: {file_name} to {target_folder}")
                break
        else:
            target_path = os.path.join(output_dir, file_name)
            safe_move(source_path, target_path)
            print(f"Moved remaining file: {file_name} to main folder")


def move_leftover_attachments(download_folder, main_folder):
    """
    Check for any leftover attachments in the download folder and move them to appropriate bid folders.
    Returns True if any files were moved, False otherwise.
    """
    print("\n🔍 Checking for leftover attachments in download folder...")
    leftover_files = [
        f for f in os.listdir(download_folder) if not f.endswith(".crdownload")
    ]

    if not leftover_files:
        print("✅ No leftover files found in download folder")
        return False

    print(f"⚠️ Found {len(leftover_files)} leftover files in download folder")

    # Get list of all bid folders
    bid_folders = [
        d
        for d in os.listdir(main_folder)
        if os.path.isdir(os.path.join(main_folder, d))
    ]

    for filename in leftover_files:
        print(f"\nProcessing leftover file: {filename}")
        source_path = os.path.join(download_folder, filename)
        moved = False

        # Try to match the file with a bid folder
        for bid_folder in bid_folders:
            if bid_folder in filename:
                destination_folder = os.path.join(main_folder, bid_folder)
                destination_path = os.path.join(destination_folder, filename)
                try:
                    safe_move(source_path, destination_path)
                    print(f"✅ Moved leftover file to matching bid folder:")
                    print(f"   From: {source_path}")
                    print(f"   To: {destination_path}")
                    moved = True
                    break
                except Exception as e:
                    print(f"❌ Error moving leftover file: {str(e)}")

        if not moved:
            # If no matching folder found, move to a special 'unmatched_attachments' folder
            unmatched_folder = os.path.join(main_folder, "unmatched_attachments")
            os.makedirs(unmatched_folder, exist_ok=True)
            try:
                safe_move(source_path, os.path.join(unmatched_folder, filename))
                print(f"⚠️ Moved to unmatched_attachments folder: {filename}")
                moved = True
            except Exception as e:
                print(f"❌ Error moving to unmatched folder: {str(e)}")

    return True


def wait_for_downloads_complete(download_folder, timeout=300):
    """
    Wait for all downloads to complete in the download folder.
    Returns True if all downloads completed, False if timeout reached.
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        # Check for any .crdownload or .tmp files
        incomplete_files = [
            f
            for f in os.listdir(download_folder)
            if f.endswith((".crdownload", ".tmp"))
        ]
        if not incomplete_files:
            # Check if any files exist in the folder
            files = [f for f in os.listdir(download_folder)]
            if files:
                return True
            # If no files found, wait a bit longer
            if (
                time.time() - start_time < 10
            ):  # Give extra 10 seconds for file to appear
                time.sleep(1)
                continue
            return False
        time.sleep(1)
    return False


def download_attachments(
    driver, solicitation_number, output_dir, download_folder, total_bids, current_bid
):
    print("\n" + "=" * 80, flush=True)
    print(f"🔄 Processing Bid {current_bid}/{total_bids}", flush=True)
    print(f"📋 Solicitation Number: {solicitation_number}", flush=True)
    print("=" * 80 + "\n", flush=True)

    # Check and move any leftover attachments before processing new bid
    move_leftover_attachments(download_folder, script_folder)

    try:
        # First check for the "Please visit" text in the portlet-body form
        try:
            portlet_body = driver.find_element(By.CLASS_NAME, "portlet-body.form")
            portlet_text = portlet_body.text.strip()

            # Check for "Please visit" text
            if "Please visit" in portlet_text:
                try:
                    link = driver.find_element(
                        By.XPATH, "/html/body/div[3]/div/div[2]/div[2]/div[2]/a"
                    )
                except NoSuchElementException:
                    link = portlet_body.find_element(By.TAG_NAME, "a")

                url = link.get_attribute("href")
                print(f"📎 Found external link with 'Please visit' text")
                print(f"🔗 URL: {url}")
                print("✅ Saving URL to Excel without downloading")
                return url

        except NoSuchElementException:
            pass

        # Check for direct download links
        direct_links = driver.find_elements(
            By.XPATH, "/html/body/div[3]/div/div[2]/div[2]/div[2]/ul/li/a"
        )
        if direct_links:
            # Create bid folder only when we know there are attachments to download
            bid_folder = os.path.join(script_folder, solicitation_number)
            os.makedirs(bid_folder, exist_ok=True)
            print(f"📁 Created folder for attachments: {bid_folder}")

            print(f"📎 Found {len(direct_links)} direct attachment links")
            downloaded_files = []

            for index, link in enumerate(direct_links, 1):
                file_name = link.text.strip()
                print(f"\n📥 Downloading attachment {index}/{len(direct_links)}")
                print(f"   File Name: {file_name}")

                # Clear download folder before each download
                for file in os.listdir(download_folder):
                    try:
                        os.remove(os.path.join(download_folder, file))
                    except Exception as e:
                        print(f"Error clearing download folder: {str(e)}")

                try:
                    # Try multiple click methods to ensure download starts
                    try:
                        # First try ActionChains
                        actions = ActionChains(driver)
                        actions.move_to_element(link).click().perform()
                        print("   🖱️ Clicked using ActionChains")
                    except:
                        try:
                            # Try JavaScript click if ActionChains fails
                            driver.execute_script("arguments[0].click();", link)
                            print("   🖱️ Clicked using JavaScript")
                        except:
                            # Direct click as last resort
                            link.click()
                            print("   🖱️ Clicked directly")

                    # Wait for download to complete
                    print("⌛ Waiting for download to complete...")
                    if wait_for_downloads_complete(download_folder):
                        # Get the downloaded file(s)
                        downloaded = [
                            f
                            for f in os.listdir(download_folder)
                            if not f.endswith((".crdownload", ".tmp"))
                        ]

                        for downloaded_file in downloaded:
                            source_path = os.path.join(download_folder, downloaded_file)
                            # Keep original extension if file was downloaded with one
                            _, ext = os.path.splitext(downloaded_file)
                            if ext:
                                target_file = f"{file_name}{ext}"
                            else:
                                target_file = file_name

                            target_path = os.path.join(bid_folder, target_file)
                            safe_move(source_path, target_path)
                            downloaded_files.append(target_file)
                            print(f"✅ Moved file: {target_file}")
                    else:
                        print("❌ Download timed out")

                except Exception as e:
                    print(f"❌ Error downloading file: {str(e)}")

                time.sleep(2)  # Short pause between downloads

            if downloaded_files:
                return ", ".join(downloaded_files)
            return ""

        # Check for the link to the attachments page
        try:
            attachments_link = driver.find_element(
                By.XPATH, "/html/body/div[3]/div/div[2]/div[2]/div[2]/a"
            )
            attachments_url = attachments_link.get_attribute("href")
            print("🔗 Found link to attachments page")

            # Check if this is a passport.cityofnewyork.us URL
            if "passport.cityofnewyork.us" in attachments_url:
                print(
                    "📌 Found passport.cityofnewyork.us URL, saving to Excel without downloading"
                )
                return attachments_url

            print(f"🌐 Navigating to attachments URL: {attachments_url}")
            driver.get(attachments_url)
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # Check if redirected to mobile version
            if "mobile.aspx" in driver.current_url:
                print("📱 Detected mobile version, switching to desktop...")
                desktop_url = driver.current_url.replace(
                    "mobile.aspx", "page.aspx"
                ).replace("?detectDevice=true", "")
                driver.get(desktop_url)
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )

            # Check if the table is present on the new page
            table = driver.find_elements(
                By.ID, "body_x_tabc_rfp_ext_prxrfp_ext_grdDocs_grd"
            )
            if table:
                # Create bid folder only when we know there are attachments to download
                bid_folder = os.path.join(script_folder, solicitation_number)
                os.makedirs(bid_folder, exist_ok=True)
                print(f"📁 Created folder for attachments: {bid_folder}")

                print("\n📑 Found attachment table")
                rows = table[0].find_elements(By.TAG_NAME, "tr")
                total_attachments = len(rows) - 1
                print(f"   Total attachments found: {total_attachments}")

                for index, row in enumerate(rows[1:], 1):
                    try:
                        # Clear download folder before each download
                        for file in os.listdir(download_folder):
                            try:
                                os.remove(os.path.join(download_folder, file))
                            except Exception as e:
                                print(f"Error clearing download folder: {str(e)}")

                        link = row.find_element(By.CSS_SELECTOR, "td:first-child a")
                        file_name = link.text.strip()
                        print(
                            f"\n📥 Downloading attachment {index}/{total_attachments}"
                        )
                        print(f"   File Name: {file_name}")

                        actions = ActionChains(driver)
                        actions.move_to_element(link).click().perform()

                        # Wait for download to complete
                        print("⌛ Waiting for download to complete...")
                        if wait_for_downloads_complete(download_folder):
                            # Move downloaded file to bid folder
                            move_downloaded_files(download_folder, bid_folder)
                        else:
                            print("❌ Download timed out")

                        time.sleep(2)  # Short pause between downloads

                    except Exception as e:
                        print(f"❌ Error downloading attachment: {str(e)}")

                print("\n✅ All table attachments processed")
                return ", ".join(
                    [
                        os.path.basename(f)
                        for f in os.listdir(bid_folder)
                        if os.path.isfile(os.path.join(bid_folder, f))
                    ]
                )
            else:
                print("❌ No attachment table found on the page.")
                return attachments_url

        except NoSuchElementException:
            print("❌ No attachments link found")
            return ""

        print("❌ No attachments found.")
        return ""

    except Exception as e:
        print(f"❌ Error processing attachments: {str(e)}")
        return ""


def download_file(url, folder, file_name):
    """
    Download a file from a URL and save it to the specified folder.
    """
    print(f"\n📥 Downloading: {file_name}")
    print(f"   URL: {url}")
    print(f"   Destination: {folder}")

    try:
        response = requests.get(url, verify=False)
        if response.status_code == 200:
            # Get content type and size
            content_type = response.headers.get("content-type", "unknown")
            file_size = len(response.content)
            print(f"   Content Type: {content_type}")
            print(f"   File Size: {file_size/1024:.2f} KB")

            # Ensure the file has an extension
            file_name, extension = os.path.splitext(file_name)
            if not extension:
                extension = mimetypes.guess_extension(content_type) or ".bin"
            file_name = f"{file_name}{extension}"

            file_path = os.path.join(folder, file_name)
            with open(file_path, "wb") as file:
                file.write(response.content)
            print(f"✅ Successfully downloaded: {file_name}")
        else:
            print(f"❌ Failed to download: {file_name}")
            print(f"   Status Code: {response.status_code}")
    except Exception as e:
        print(f"❌ Error downloading file: {str(e)}")


def move_downloaded_files(download_folder, bid_folder):
    """
    Move downloaded files from the download directory to the bid folder.
    """
    print("\n📦 Moving downloaded files...")
    moved_files = []

    for filename in os.listdir(download_folder):
        if filename.endswith(".crdownload"):
            print(f"⏳ Skipping incomplete download: {filename}")
            continue

        source = os.path.join(download_folder, filename)
        destination = os.path.join(bid_folder, filename)

        try:
            safe_move(source, destination)
            moved_files.append(filename)
            print(f"✅ Moved: {filename}")
            print(f"   From: {source}")
            print(f"   To: {destination}")
        except Exception as e:
            print(f"❌ Error moving {filename}: {str(e)}")

    if moved_files:
        print(f"\n✅ Successfully moved {len(moved_files)} files")
    else:
        print("\n⚠️ No files were moved")


def move_specific_files(download_folder, bid_folder, file_names):
    for file_name in file_names:
        found = False
        for downloaded_file in os.listdir(download_folder):
            if file_name.lower() in downloaded_file.lower():
                source_path = os.path.join(download_folder, downloaded_file)
                target_path = os.path.join(bid_folder, downloaded_file)
                try:
                    safe_move(source_path, target_path)
                    print(f"Moved file: {downloaded_file} to {bid_folder}")
                    found = True
                    break
                except Exception as e:
                    print(f"Error moving file {downloaded_file}: {str(e)}")
        if not found:
            print(f"File not found: {file_name}")


def move_remaining_files(download_folder, output_dir):
    print("Checking for remaining files in the download folder...")
    remaining_files = [
        f for f in os.listdir(download_folder) if not f.endswith(".crdownload")
    ]

    for file_name in remaining_files:
        source_path = os.path.join(download_folder, file_name)

        for bid_folder in os.listdir(output_dir):
            if bid_folder in file_name:
                target_folder = os.path.join(output_dir, bid_folder)
                target_path = os.path.join(target_folder, file_name)
                try:
                    if not os.path.exists(target_path):
                        safe_move(source_path, target_path)
                        print(f"Moved remaining file: {file_name} to {target_folder}")
                    else:
                        new_file_name = f"duplicate_{file_name}"
                        new_target_path = os.path.join(target_folder, new_file_name)
                        safe_move(source_path, new_target_path)
                        print(
                            f"Moved duplicate file: {file_name} as {new_file_name} to {target_folder}"
                        )
                    break
                except Exception as e:
                    print(f"Error moving remaining file {file_name}: {str(e)}")
        else:
            target_path = os.path.join(output_dir, file_name)
            try:
                if not os.path.exists(target_path):
                    safe_move(source_path, target_path)
                    print(f"Moved remaining file: {file_name} to main folder")
                else:
                    new_file_name = f"duplicate_{file_name}"
                    new_target_path = os.path.join(output_dir, new_file_name)
                    safe_move(source_path, new_target_path)
                    print(
                        f"Moved duplicate file: {file_name} as {new_file_name} to main folder"
                    )
            except Exception as e:
                print(f"Error moving remaining file {file_name}: {str(e)}")


def remove_empty_folders(output_dir):
    """
    Remove empty folders in the output directory.
    """
    for folder_name in os.listdir(output_dir):
        folder_path = os.path.join(output_dir, folder_name)
        if os.path.isdir(folder_path) and not os.listdir(folder_path):
            os.rmdir(folder_path)
            print(f"Removed empty folder: {folder_path}")


def update_attachments_for_bid(excel_path, solicitation_number, output_dir):
    """
    Update the Attachments column in the Excel file for a specific bid.
    """
    df = pd.read_excel(excel_path)
    bid_folder = os.path.join(script_folder, solicitation_number)

    if solicitation_number not in df["Solicitation Number"].values:
        print(f"Solicitation Number {solicitation_number} not found in Excel file.")
        return

    if os.path.exists(bid_folder):
        attachments = [
            f
            for f in os.listdir(bid_folder)
            if os.path.isfile(os.path.join(bid_folder, f))
        ]
        attachments_str = ", ".join(attachments) if attachments else ""
    else:
        attachments_str = ""

    # Convert the column to object type before assigning the value
    df["Attachments"] = df["Attachments"].astype(object)
    df.loc[df["Solicitation Number"] == solicitation_number, "Attachments"] = (
        attachments_str
    )
    df.to_excel(excel_path, index=False)
    print(
        f"Updated Attachments for bid {solicitation_number} in Excel file: {excel_path}"
    )


def parse_date(date_string):
    """
    Parse date string and return a datetime object.

    :param date_string: String representation of the date
    :return: datetime object
    """
    try:
        # Handle range format
        if " - " in date_string:
            date_string = date_string.split(" - ")[
                0
            ]  # Take the first date in the range

        # Use dateutil parser to handle various date formats
        parsed_date = parser.parse(date_string, fuzzy=True)
        return parsed_date.date()  # Return only the date part
    except ValueError:
        # If parsing fails, raise an exception
        raise ValueError(f"Unable to parse date string: {date_string}")


def move_downloaded_files(download_folder, bid_folder):
    """
    Move downloaded files from the download directory to the bid folder.
    """
    for filename in os.listdir(download_folder):
        if filename.endswith(".crdownload"):
            continue  # Skip files that are still downloading
        source = os.path.join(download_folder, filename)
        destination = os.path.join(bid_folder, filename)
        safe_move(source, destination)
        print(f"Moved {filename} to {bid_folder}")


def move_remaining_downloads(download_folder, main_folder):
    """
    Move any remaining files in the download directory to their respective bid folders.
    """
    for filename in os.listdir(download_folder):
        source_path = os.path.join(download_folder, filename)
        if os.path.isfile(source_path):
            # Try to find a matching bid folder
            for bid_folder in os.listdir(main_folder):
                if (
                    os.path.isdir(os.path.join(main_folder, bid_folder))
                    and bid_folder in filename
                ):
                    destination_path = os.path.join(main_folder, bid_folder, filename)
                    safe_move(source_path, destination_path)
                    print(f"Moved remaining file {filename} to {bid_folder}")
                    break
            else:
                # If no matching bid folder found, move to main folder
                destination_path = os.path.join(main_folder, filename)
                safe_move(source_path, destination_path)
                print(f"Moved remaining file {filename} to main folder")


def scrape_city_record():
    """
    Main function to execute the NYC procurement bid extraction process.
    """
    print("Bids Extraction Started")
    print(f"📁 Saving data in folder: {main_folder}")

    driver = setup_driver()
    print("Browser initialized")

    try:
        login_attempts = 0
        max_login_attempts = 3

        while login_attempts < max_login_attempts:
            print(f"Attempting login (attempt {login_attempts + 1})")
            if login(driver, os.getenv("NYC_EMAIL"), os.getenv("NYC_PASSWORD")):
                print("Login successful")
                break
            login_attempts += 1
            print(f"Login attempt {login_attempts} failed. Retrying...")
            random_sleep(5, 10)

        if login_attempts == max_login_attempts:
            print("Failed to log in after multiple attempts. Exiting.")
            return

        if not navigate_to_advanced_search(driver):
            print("Failed to navigate to Advanced Search page. Exiting.")
            return

        apply_filters(driver)

        total_bids_processed = 0
        # Excel file saved in the script folder with IN_PROGRESS suffix
        excel_path = os.path.join(script_folder, "02_NYC.xlsx")

        while True:
            bids = extract_bid_data(driver, max_bids=30)
            total_bids = len(bids)

            for index, bid in enumerate(bids, start=1):
                try:
                    bid_link = bid.find("a")["href"]
                    driver.get("https://a856-cityrecord.nyc.gov" + bid_link)
                    random_sleep(2, 4)

                    has_available_date, available_date = check_available_date(
                        driver, bid_link
                    )

                    if has_available_date:
                        print(
                            f"Found 'Available Date': {available_date}. Stopping extraction."
                        )
                        break

                    bid_details = extract_bid_details(
                        driver, bid, total_bids_processed + index
                    )

                    # Pass main_folder instead of a subfolder path
                    attachments_result = download_attachments(
                        driver,
                        bid_details["Solicitation Number"],
                        main_folder,  # Changed from download_folder to main_folder
                        download_folder,
                        total_bids,
                        index,
                    )
                    bid_details["Attachments"] = attachments_result

                    update_excel(bid_details, excel_path)
                    update_attachments_for_bid(
                        excel_path, bid_details["Solicitation Number"], main_folder
                    )

                    total_bids_processed += 1
                except Exception as e:
                    print(f"Error processing bid {index}: {str(e)}")
                    traceback.print_exc()  # Print the full traceback for debugging
                    play_notification_sound()
                    user_input = input(
                        "Press Enter to continue or type 'exit' to stop: "
                    )
                    if user_input.lower() == "exit":
                        break  # Exit the loop if user chooses to stop

            print(
                f"Bids with Posting Range within 2 Days Successfully Extracted: {total_bids_processed}"
            )

            if has_available_date or len(bids) < 30:
                print("Ending extraction.")
                break

            # Navigate back to the search results page
            driver.execute_script("window.history.go(-1)")
            random_sleep(2, 4)

            # Navigate to the next page
            next_page_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(@class, 'page-link next')]")
                )
            )
            driver.execute_script("arguments[0].click();", next_page_link)
            random_sleep(2, 4)

        # After all processing is done:
        move_remaining_downloads(download_folder, main_folder)
        remove_empty_folders(main_folder)

        # Remove the bulk update of SL No as it's now handled in update_excel function

        if total_bids_processed > 0:
            print(f"💾 Final bid results saved to 02_NYC.xlsx in {script_folder}")
        else:
            print("❌ No bids were processed. No Excel file created.")

        print("🎉 All Bids and Attachments Extraction Successfully Completed")

        # Rename the script folder to indicate completion
        completed_folder = script_folder.replace("_IN_PROGRESS", "_COMPLETED")
        if os.path.exists(script_folder):
            os.rename(script_folder, completed_folder)
            print(f"Renamed folder to indicate completion: {completed_folder}")

            # Move the Excel file to the completed folder
            old_excel_path = excel_path
            new_excel_path = os.path.join(completed_folder, "02_NYC.xlsx")
            if os.path.exists(old_excel_path):
                safe_move(old_excel_path, new_excel_path)
                print(f"Moved Excel file to completed folder: {new_excel_path}")

        play_notification_sound()
    except Exception as e:
        print(f"Error during scraping: {str(e)}")
        play_notification_sound()
        input("Press Enter to continue...")
        traceback.print_exc()
        driver.save_screenshot("scraping_error.png")
        print("Screenshot saved as scraping_error.png")
    finally:
        driver.quit()
        print("Browser closed")

        # Remove the temporary download folder
        shutil.rmtree(download_folder, ignore_errors=True)
        print(f"Removed temporary download folder: {download_folder}")


if __name__ == "__main__":
    scrape_city_record()
