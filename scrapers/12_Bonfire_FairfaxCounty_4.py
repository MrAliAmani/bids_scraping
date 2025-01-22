from asyncio.log import logger
import sys
import os
import io

import urllib3

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.utils import safe_move

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import time
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    WebDriverException,
)
import shutil
import re
from dotenv import load_dotenv
import argparse
import winsound
import logging
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
import tempfile
import atexit
from selenium.webdriver.chrome.options import Options
from selenium_stealth import stealth
from fake_useragent import UserAgent
import json
from pathlib import Path

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Set up logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Load environment variables
load_dotenv()

# Add this near the top of your script, after loading environment variables
if not os.getenv("FAIRFAX_EMAIL") or not os.getenv("FAIRFAX_PASSWORD"):
    logger.error("FAIRFAX_EMAIL or FAIRFAX_PASSWORD environment variables are not set.")
    sys.exit(1)

# Add this list of URLs at the beginning of the script, after the imports
BONFIRE_URLS = [
    "https://cityofmilwaukee.bonfirehub.com/portal/?tab=openOpportunities",
    "https://utrgv.bonfirehub.com/portal/?tab=openOpportunities",
    "https://fortbendisd.bonfirehub.com/portal/?tab=openOpportunities",
    "https://tuhsd.bonfirehub.com/portal/?tab=openOpportunities",
    "https://npc.bonfirehub.com/portal/?tab=openOpportunities",
    "https://fairfaxcounty.bonfirehub.com/portal/?tab=openOpportunities",
]

# Update the script_name variable
script_name = os.path.splitext(os.path.basename(__file__))[0]

# Parse command-line arguments
parser = argparse.ArgumentParser(description="Scrape Bonfire FairfaxCounty bids.")
parser.add_argument(
    "--days", type=int, default=2, help="Number of days to scrape (default: 2)"
)
args = parser.parse_args()

# Use the provided number of days or default to 2
days_to_scrape = args.days

# Create a folder named after yesterday's date
yesterday = datetime.now() - timedelta(days=1)
main_folder = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), yesterday.strftime("%Y-%m-%d")
)
os.makedirs(main_folder, exist_ok=True)
logging.info(f"Main folder created at: {main_folder}")


def play_notification():
    """Play a notification sound."""
    frequency = 2500
    duration = 1000
    winsound.Beep(frequency, duration)


def handle_error(error_message):
    """Handle errors by logging, playing a notification and pausing the script."""
    logging.error(f"Error: {error_message}")
    play_notification()
    input("Press Enter to continue...")


def create_temp_folder():
    """Create a temporary folder with IN_PROGRESS suffix and a downloads subfolder"""
    # Create the script folder with IN_PROGRESS suffix
    script_folder = os.path.join(main_folder, f"{script_name}_IN_PROGRESS")
    os.makedirs(script_folder, exist_ok=True)

    # Create downloads subfolder within the script folder
    downloads_folder = os.path.join(script_folder, script_name)
    os.makedirs(downloads_folder, exist_ok=True)

    def cleanup_temp(path):
        try:
            if os.path.exists(path):
                # Rename the folder to indicate completion
                completed_folder = path.replace("_IN_PROGRESS", "_COMPLETED")
                if os.path.exists(completed_folder):
                    shutil.rmtree(completed_folder)
                os.rename(path, completed_folder)
                logging.info(f"Successfully renamed folder to: {completed_folder}")
        except Exception as e:
            logging.error(f"Warning: Could not handle temp directory {path}: {e}")

    atexit.register(cleanup_temp, script_folder)
    return downloads_folder


def setup_driver():
    """Set up and configure the Chrome WebDriver."""
    chrome_options = webdriver.ChromeOptions()
    temp_download_dir = create_temp_folder()
    prefs = {
        "download.default_directory": temp_download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Add these lines to bypass SSL certificate verification
    chrome_options.add_argument("--ignore-certificate-errors-spki-list")
    chrome_options.add_argument("--ignore-ssl-errors")

    # Use ChromeDriverManager to automatically download and manage the ChromeDriver
    service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver, temp_download_dir


def is_logged_in(driver):
    try:
        # Check for elements that are visible when logged in
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "opportunities-table"))
        )
        # Additional check for user menu or profile indicator
        user_menu = driver.find_element(
            By.XPATH, "//button[contains(@class, 'user-menu-button')]"
        )
        return True
    except (NoSuchElementException, TimeoutException):
        try:
            # Check for the login button, which should not be present if logged in
            login_button = driver.find_element(
                By.XPATH, "//a[@class='login-button' and contains(text(), 'Log In')]"
            )
            return False
        except NoSuchElementException:
            # If login button is not found, assume we're logged in
            return True


def login(driver, max_retries=3):
    logger.info("Checking login status...")

    for attempt in range(max_retries):
        try:
            driver.get(
                "https://co-newton-ga.bonfirehub.com/portal/?tab=openOpportunities"
            )
            time.sleep(5)  # Wait for page to load

            if is_logged_in(driver):
                logger.info("Already logged in")
                return True

            logger.info(f"Not logged in. Proceeding with login attempt {attempt + 1}")

            # Click the "Log In" button
            try:
                login_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//a[@class='login-button' and contains(text(), 'Log In')]",
                        )
                    )
                )
                driver.execute_script("arguments[0].click();", login_button)
                logger.info("Clicked initial Log In button")
            except Exception as e:
                logger.error(f"Failed to click initial Log In button: {str(e)}")
                continue

            # Wait for the login page to load
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, "input-email"))
                )
                logger.info("Login page loaded")
            except Exception as e:
                logger.error(f"Login page did not load as expected: {str(e)}")
                continue

            # Enter email
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, "input-email"))
                )
                email_input = driver.find_element(By.ID, "input-email")
                if email_input:
                    email_input.clear()
                    email = os.getenv("FAIRFAX_EMAIL")
                    if email:
                        email_input.send_keys(email)
                        logger.info(f"Entered email: {email}")
                    else:
                        logger.error("FAIRFAX_EMAIL environment variable is not set")
                        return False
                else:
                    logger.error("Email input field not found")
                    return False
            except Exception as e:
                logger.error(f"Failed to enter email: {str(e)}")
                logger.error(f"Page source: {driver.page_source}")
                return False

            # Click the "Continue" button
            try:
                continue_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//button[@type='submit' and contains(text(), 'Continue')]",
                        )
                    )
                )
                driver.execute_script("arguments[0].click();", continue_button)
                logger.info("Clicked Continue button")
            except Exception as e:
                logger.error(f"Failed to click Continue button: {str(e)}")
                continue

            # Wait for password field to appear
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, "input-password"))
                )
                logger.info("Password field appeared")
            except Exception as e:
                logger.error(f"Password field did not appear: {str(e)}")
                continue

            # Enter password
            try:
                password_input = driver.find_element(By.ID, "input-password")
                password_input.clear()
                password = os.getenv("FAIRFAX_PASSWORD")
                password_input.send_keys(password)
                logger.info("Entered password")
            except Exception as e:
                logger.error(f"Failed to enter password: {str(e)}")
                continue

            # Click the final "Log In" button
            try:
                login_submit_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//button[@type='submit' and contains(text(), 'Log In')]",
                        )
                    )
                )
                driver.execute_script("arguments[0].click();", login_submit_button)
                logger.info("Clicked final Log In button")
            except Exception as e:
                logger.error(f"Failed to click final Log In button: {str(e)}")
                continue

            # Wait for a short time and then check if we're logged in
            time.sleep(10)
            if is_logged_in(driver):
                logger.info("Login successful")
                return True
            else:
                logger.warning("Login may have failed. Retrying...")

        except Exception as e:
            logger.warning(f"Login attempt {attempt + 1} failed: {str(e)}")

        if attempt < max_retries - 1:
            logger.info(f"Retrying login in 5 seconds...")
            time.sleep(5)
        else:
            logger.error("All login attempts failed")
            return False

    return False


def extract_bid_links(driver):
    """Extract bid links from the search results page."""
    bid_links = []
    max_retries = 3
    max_scrolls = 10  # Set a maximum number of scrolls to prevent infinite loops
    scroll_wait_time = 10  # Increase wait time between scrolls
    initial_page_load_wait = 15  # Increase initial page load wait time

    logging.info("Starting to extract bid links")

    # Wait for initial page load
    try:
        WebDriverWait(driver, initial_page_load_wait).until(
            EC.presence_of_element_located(
                (By.XPATH, "//table[@id='DataTables_Table_1']/tbody/tr")
            )
        )
    except TimeoutException:
        logging.warning(f"Page did not load within {initial_page_load_wait} seconds")
        return bid_links

    for scroll_attempt in range(max_scrolls):
        logging.info(f"Scroll attempt {scroll_attempt + 1} of {max_scrolls}")

        for attempt in range(max_retries):
            try:
                # Wait for table rows to load
                rows = WebDriverWait(driver, 60).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//table[@id='DataTables_Table_1']/tbody/tr")
                    )
                )

                # Extract links
                new_links = []
                for row in rows:
                    try:
                        link = row.find_element(By.XPATH, ".//td[6]/a").get_attribute(
                            "href"
                        )
                        if link not in bid_links:
                            new_links.append(link)
                    except NoSuchElementException:
                        continue

                bid_links.extend(new_links)
                logging.info(
                    f"Found {len(new_links)} new links (Total: {len(bid_links)} / {len(rows)} visible bids)"
                )

                if not new_links:
                    logging.info(
                        "No new links found. All visible bids have been processed."
                    )
                    return bid_links

                break  # Success, move to next scroll

            except TimeoutException:
                if attempt < max_retries - 1:
                    logging.warning(
                        f"Timeout on scroll attempt {scroll_attempt + 1}, retry {attempt + 1}. Retrying..."
                    )
                    time.sleep(scroll_wait_time)  # Wait longer before retrying
                else:
                    logging.error(
                        f"Failed to load content after {max_retries} attempts. Stopping extraction."
                    )
                    return bid_links

        # Scroll to bottom of page more gradually
        previous_height = driver.execute_script("return document.body.scrollHeight")
        driver.execute_script(f"window.scrollTo(0, {previous_height});")
        time.sleep(scroll_wait_time)  # Wait longer for content to load

        # Check if we've reached the bottom
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == previous_height:
            logging.info("Reached bottom of page - no more bids to load")
            break

    logging.info(f"Completed bid extraction. Total links found: {len(bid_links)}")
    return bid_links


def extract_bid_details(driver, bid_link, url):
    """Extract details from a specific bid page."""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            driver.get(bid_link)
            bid_details = {}

            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "opportunityMetaSection"))
            )

            def safe_extract(xpath, key):
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    if key in [
                        "Notice Type",
                        "Solicitation Number",
                        "Solicitation Title",
                    ]:
                        bid_details[key] = (
                            element.text.split(":", 1)[1].strip()
                            if ":" in element.text
                            else ""
                        )
                    else:
                        bid_details[key] = element.text.strip()
                    logging.info(f"Successfully extracted {key}: {bid_details[key]}")
                except (TimeoutException, NoSuchElementException):
                    logging.warning(
                        f"Could not find element for {key} using xpath: {xpath}"
                    )
                    bid_details[key] = ""

            safe_extract(
                "//span[contains(@class, 'js-opportunity-date-open')]", "Posted Date"
            )
            safe_extract(
                "//span[contains(@class, 'js-opportunity-date-close')]", "Response Date"
            )
            safe_extract(
                "//div[contains(@class, 'modalSection projectDetailSection')]/b[text()='Type:']/..",
                "Notice Type",
            )
            safe_extract(
                "//div[contains(@class, 'modalSection projectDetailSection')]/b[text()='Ref. #:']/..",
                "Solicitation Number",
            )
            safe_extract(
                "//div[contains(@class, 'modalSection projectDetailSection')]/b[text()='Project:']/..",
                "Solicitation Title",
            )
            safe_extract("//div[@id='publicTitle']//b", "Agency")

            # Extract NIGP codes and category text
            try:
                # Get NIGP codes and their corresponding texts
                category_pairs = []

                # Find all commodity code elements and their corresponding text elements
                commodity_sections = driver.find_elements(
                    By.XPATH, "//div[contains(@class, 'commodityCodeSection')]"
                )

                for i, section in enumerate(
                    commodity_sections[:2]
                ):  # Limit to first 2 pairs
                    try:
                        # Get NIGP code
                        nigp_code = section.find_element(
                            By.XPATH, ".//span[contains(@class, 'commodityCodeBadge')]"
                        ).text.strip()

                        # Get corresponding text
                        category_text = section.find_element(
                            By.XPATH, ".//strong"
                        ).text.strip()

                        if nigp_code and category_text:
                            category_pairs.append(f"NIGP {nigp_code} - {category_text}")
                    except NoSuchElementException:
                        continue

                # Combine the pairs
                if category_pairs:
                    bid_details["Category"] = " | ".join(category_pairs)
                else:
                    bid_details["Category"] = ""

                logging.info(f"Extracted Category: {bid_details['Category']}")
            except Exception as e:
                logging.warning(f"Error extracting category information: {str(e)}")
                bid_details["Category"] = ""

            safe_extract(
                "//div[contains(@class, 'bfMarkdown markdown_formatted')]",
                "Description",
            )
            bid_details["Bid Detail Page URL"] = bid_link

            bid_details["Additional Summary"] = ""
            bid_details["Contracting Office Address"] = ""
            bid_details["Contact Information"] = ""

            for date_field in ["Posted Date", "Response Date"]:
                if bid_details[date_field]:
                    date_match = re.search(
                        r"(\w+)\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\d{4})",
                        bid_details[date_field],
                    )
                    if date_match:
                        month, day, year = date_match.groups()
                        date_obj = datetime.strptime(
                            f"{month} {day} {year}", "%b %d %Y"
                        )
                        bid_details[date_field] = date_obj.strftime("%Y-%m-%d")
                        logging.info(
                            f"Formatted {date_field}: {bid_details[date_field]}"
                        )
                    else:
                        logging.warning(
                            f"Could not parse date for {date_field}: {bid_details[date_field]}"
                        )

            attachments = driver.find_elements(
                By.XPATH, "//table[@id='publicProjectDocumentsTable']//td[1]"
            )
            bid_details["Attachments"] = ", ".join([att.text for att in attachments])
            logging.info(f"Extracted Attachments: {bid_details['Attachments']}")

            return bid_details
        except WebDriverException as e:
            if "target window already closed" in str(e) and attempt < max_retries - 1:
                logging.warning(
                    f"WebDriver lost connection. Attempt {attempt + 1} of {max_retries}. Restarting driver..."
                )
                driver.quit()
                driver = setup_driver()
                login(driver)
            elif attempt < max_retries - 1:
                logging.warning(
                    f"Error extracting bid details. Attempt {attempt + 1} of {max_retries}. Retrying..."
                )
                time.sleep(5)
            else:
                handle_error(
                    f"Error extracting bid details after {max_retries} attempts: {str(e)}"
                )
                return None


def download_attachments(driver, bid_number, temp_download_dir):
    """Download attachments for a specific bid."""
    # Create bid folder inside the script folder (with IN_PROGRESS suffix), not main folder
    script_folder = os.path.join(main_folder, f"{script_name}_IN_PROGRESS")
    bid_folder = os.path.join(script_folder, bid_number.replace("/", "_"))
    os.makedirs(bid_folder, exist_ok=True)

    download_buttons = driver.find_elements(
        By.XPATH, "//button[contains(@onclick, 'BFUtil.downloadDocument')]"
    )
    total_attachments = len(download_buttons)
    logging.info(f"Total attachments to download: {total_attachments}")

    for index, button in enumerate(download_buttons, start=1):
        max_retries = 3
        for attempt in range(max_retries):
            try:
                logging.info(f"Downloading attachment {index} of {total_attachments}")
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button))

                # Click the download button
                button.click()

                # Wait for alert and accept it
                try:
                    alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
                    alert_text = alert.text
                    logging.info(f"Alert detected: {alert_text}")
                    alert.accept()
                    logging.info("Alert accepted")
                except TimeoutException:
                    logging.info("No alert detected, continuing with download")

                time.sleep(2)
                break
            except StaleElementReferenceException:
                if attempt < max_retries - 1:
                    logging.warning(f"Stale element reference: {button}. Retrying...")
                    download_buttons = driver.find_elements(
                        By.XPATH,
                        "//button[contains(@onclick, 'BFUtil.downloadDocument')]",
                    )
                    continue
                else:
                    logging.error(
                        f"Failed to click download button after {max_retries} attempts"
                    )
            except Exception as e:
                if attempt < max_retries - 1:
                    logging.warning(
                        f"Error clicking download button: {str(e)}. Retrying..."
                    )
                    time.sleep(2)

                    # Try to handle any lingering alerts
                    try:
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        logging.info(f"Found lingering alert: {alert_text}")
                        alert.accept()
                        logging.info("Lingering alert accepted")
                    except:
                        pass
                else:
                    handle_error(
                        f"Error clicking download button after {max_retries} attempts: {str(e)}"
                    )

    logging.info("Waiting for downloads to complete...")
    time.sleep(10)  # Wait longer for downloads to complete

    # Check for and handle any remaining alerts
    try:
        alert = driver.switch_to.alert
        alert.accept()
        logging.info("Accepted final alert")
    except:
        pass

    moved_files = []
    max_wait = 30  # Maximum seconds to wait for downloads
    start_time = time.time()

    while time.time() - start_time < max_wait:
        # Check for ongoing downloads
        downloads_in_progress = False
        for file in os.listdir(temp_download_dir):
            if file.endswith(".crdownload"):
                downloads_in_progress = True
                break

        if not downloads_in_progress:
            break

        time.sleep(2)
        logging.info("Waiting for downloads to complete...")

    # Move completed downloads
    for file in os.listdir(temp_download_dir):
        source = os.path.join(temp_download_dir, file)
        if os.path.isfile(source) and not file.endswith(".crdownload"):
            destination = os.path.join(bid_folder, file)
            safe_move(source, destination)
            moved_files.append(file)

    if moved_files:
        logging.info(
            f"Successfully downloaded and moved {len(moved_files)} attachments"
        )
    else:
        logging.warning("No attachments were successfully downloaded and moved")

    return moved_files


def update_excel(bid_data, downloaded_attachments):
    """Update the Excel file with new bid data."""
    # Excel file should be in the script folder (with IN_PROGRESS suffix), not main folder
    script_folder = os.path.join(main_folder, f"{script_name}_IN_PROGRESS")
    excel_file = os.path.join(script_folder, f"{script_name}.xlsx")
    logging.info(f"Attempting to save Excel file at: {excel_file}")

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

    bid_data["Attachments"] = (
        ", ".join(downloaded_attachments) if downloaded_attachments else ""
    )

    try:
        if os.path.exists(excel_file):
            logging.info(f"Existing Excel file found. Updating: {excel_file}")
            df = pd.read_excel(excel_file)
            df = pd.concat([df, pd.DataFrame([bid_data])], ignore_index=True)
        else:
            logging.info(f"Creating new Excel file: {excel_file}")
            df = pd.DataFrame([bid_data])

        df = df.reindex(columns=column_order)
        df.to_excel(excel_file, index=False)
        logging.info(f"Successfully updated Excel file: {excel_file}")
    except Exception as e:
        handle_error(f"Error updating Excel file: {str(e)}")


def load_bid_cache():
    """Load the cached bid data from JSON file."""
    cache_file = Path("cache/bonfire_bids_cache.json")
    if not cache_file.exists():
        cache_file.parent.mkdir(exist_ok=True)
        return {}

    try:
        with cache_file.open("r") as f:
            cache = json.load(f)
            # Convert string dates back to datetime objects
            for url, data in cache.items():
                if data["posted_date"]:
                    data["posted_date"] = datetime.strptime(
                        data["posted_date"], "%Y-%m-%d"
                    )
            return cache
    except Exception as e:
        logging.warning(f"Error loading bid cache: {e}")
        return {}


def save_bid_cache(cache):
    """Save the bid cache to JSON file."""
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    cache_file = Path(f"cache/{script_name}.json")
    try:
        # Convert datetime objects to strings for JSON serialization
        serializable_cache = {}
        for url, data in cache.items():
            try:
                posted_date = None
                if data["posted_date"]:
                    if isinstance(data["posted_date"], str):
                        posted_date = data["posted_date"]
                    else:
                        posted_date = data["posted_date"].strftime("%Y-%m-%d")

                last_checked = (
                    data["last_checked"].strftime("%Y-%m-%d")
                    if hasattr(data["last_checked"], "strftime")
                    else data["last_checked"]
                )

                serializable_cache[url] = {
                    "posted_date": posted_date,
                    "last_checked": last_checked,
                }
            except Exception as e:
                logging.warning(f"Error processing cache entry for {url}: {e}")
                continue

        with cache_file.open("w") as f:
            json.dump(serializable_cache, f, indent=2)
            logging.info(
                f"Successfully saved {len(serializable_cache)} entries to cache"
            )
    except Exception as e:
        logging.warning(f"Error saving bid cache: {e}")


def should_process_bid(bid_link, bid_cache):
    """
    Determine if a bid should be processed based on cache data.
    Returns (should_process, reason) tuple.
    """
    today = datetime.now().date()

    if bid_link not in bid_cache:
        return True, "New bid - not in cache"

    cache_entry = bid_cache[bid_link]

    # Handle string or datetime objects for last_checked
    try:
        last_checked = (
            datetime.strptime(cache_entry["last_checked"], "%Y-%m-%d").date()
            if isinstance(cache_entry["last_checked"], str)
            else cache_entry["last_checked"].date()
        )
    except (TypeError, AttributeError):
        return True, "Invalid last_checked date in cache"

    # If we haven't checked this bid in the last 30 days, check it again
    if (today - last_checked).days > 30:
        return (
            True,
            f"Last checked {(today - last_checked).days} days ago - needs refresh",
        )

    # If the bid has a posted date, check if it's too old
    if cache_entry["posted_date"]:
        try:
            posted_date = (
                datetime.strptime(cache_entry["posted_date"], "%Y-%m-%d").date()
                if isinstance(cache_entry["posted_date"], str)
                else cache_entry["posted_date"].date()
            )

            days_old = (today - posted_date).days
            if days_old > days_to_scrape:
                return (
                    False,
                    f"Bid is {days_old} days old (older than {days_to_scrape} days cutoff)",
                )
        except (TypeError, AttributeError):
            return True, "Invalid posted_date in cache"

    return (
        True,
        f"Bid was checked {(today - last_checked).days} days ago but no posted date found",
    )


def process_single_site(driver, url, temp_download_dir, url_number, total_urls):
    """Process a single Bonfire site."""
    logging.info(f"Processing URL {url_number}/{total_urls}: {url}")

    # Load the bid cache
    bid_cache = load_bid_cache()
    cache_modified = False

    try:
        driver.get(url)
        time.sleep(10)

        bid_links = extract_bid_links(driver)
        logging.info(f"Found {len(bid_links)} bid links")

        for index, link in enumerate(bid_links, start=1):
            # Check if we should process this bid
            should_process, reason = should_process_bid(link, bid_cache)

            if link in bid_cache:
                cache_info = bid_cache[link]
                posted_date_str = (
                    cache_info["posted_date"].strftime("%Y-%m-%d")
                    if hasattr(cache_info["posted_date"], "strftime")
                    else cache_info["posted_date"]
                )
                last_checked_str = (
                    cache_info["last_checked"].strftime("%Y-%m-%d")
                    if hasattr(cache_info["last_checked"], "strftime")
                    else cache_info["last_checked"]
                )

                logging.info(
                    f"[URL {url_number}/{total_urls}] Bid {index}/{len(bid_links)}:"
                )
                logging.info(
                    f"  └─ Cache Status: Posted: {posted_date_str}, Last Checked: {last_checked_str}"
                )
                logging.info(
                    f"  └─ Decision: {'Will process' if should_process else 'Skipping'} - {reason}"
                )
            else:
                logging.info(
                    f"[URL {url_number}/{total_urls}] Bid {index}/{len(bid_links)}: Not in cache - will process"
                )

            if not should_process:
                continue

            bid_details = extract_bid_details(driver, link, url)
            if bid_details:
                posted_date = (
                    datetime.strptime(bid_details["Posted Date"], "%Y-%m-%d")
                    if bid_details["Posted Date"]
                    else None
                )

                # Update cache with this bid's information
                bid_cache[link] = {
                    "posted_date": posted_date,
                    "last_checked": datetime.now(),
                }
                cache_modified = True

                if posted_date and posted_date >= datetime.now() - timedelta(
                    days=days_to_scrape
                ):
                    bid_details["SL No"] = index
                    downloaded_attachments = download_attachments(
                        driver, bid_details["Solicitation Number"], temp_download_dir
                    )
                    update_excel(bid_details, downloaded_attachments)
                    logging.info(
                        f"✅ Successfully processed bid: {bid_details['Solicitation Number']}"
                    )
                else:
                    logging.info(
                        f"Skipping bid {bid_details.get('Solicitation Number', 'Unknown')} - older than {days_to_scrape} days or no date"
                    )

            time.sleep(5)

        # Save the updated cache if it was modified
        if cache_modified:
            save_bid_cache(bid_cache)

    except Exception as e:
        logging.error(
            f"Error processing site {url_number}/{total_urls} - {url}: {str(e)}"
        )
        # Still try to save cache if modified
        if cache_modified:
            save_bid_cache(bid_cache)


def take_screenshot(driver, filename):
    try:
        screenshot_path = os.path.join(main_folder, f"{filename}.png")
        driver.save_screenshot(screenshot_path)
        logger.info(f"Screenshot saved: {screenshot_path}")
    except Exception as e:
        logger.error(f"Failed to take screenshot {filename}: {str(e)}")


def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    ua = UserAgent()
    chrome_options.add_argument(f"user-agent={ua.random}")

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

    return driver


def login_to_first_portal(driver, url):
    driver.get(url)
    # Add login logic here if needed
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "some-login-element-id"))
    )
    print("Login successful.")


def scrape_portal(driver, url):
    driver.get(url)
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "some-element-id"))
    )
    print(f"Scraping portal: {url}")
    # Add scraping logic here


def main():
    logging.info(
        f"🚀 Bids Extraction Started (Scraping bids from the last {days_to_scrape} days)"
    )
    driver, temp_download_dir = setup_driver()

    try:
        # Attempt to log in once at the beginning
        if not login(driver):
            raise Exception("Failed to log in after multiple attempts")

        # Double-check login status
        if is_logged_in(driver):
            logging.info("Login confirmed successful")
        else:
            logging.warning("Login may have failed. Proceeding with caution.")

        total_urls = len(BONFIRE_URLS)
        for url_number, url in enumerate(BONFIRE_URLS, start=1):
            process_single_site(driver, url, temp_download_dir, url_number, total_urls)
            time.sleep(10)  # Wait between processing different sites

        logging.info("🎉 All Bids and Attachments Extraction Successfully Completed")
        play_notification()

    except Exception as e:
        handle_error(f"An unexpected error occurred: {str(e)}")
        logging.error("Traceback:", exc_info=True)

    finally:
        driver.quit()
        shutil.rmtree(temp_download_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
