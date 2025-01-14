import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", write_through=True)

import os
import shutil

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
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from selenium.webdriver.common.action_chains import ActionChains
import argparse
from utils.utils import safe_move, play_notification_sound
from concurrent.futures import ThreadPoolExecutor
import concurrent
from urllib.parse import urlparse
import json

# Add these imports at the top
import signal
import threading

# Add these global variables after imports
stop_processing = threading.Event()
active_drivers = []
thread_pool = None
base_folder = None
main_folder = None
script_download_folder = None
script_name = None

# Add log_lock definition after other global variables
log_lock = threading.Lock()

# Add these constants after the imports
SITES_CONFIG = {
    "gvibuy.buyspeed.com": {"name": "GVI Buy", "prefix": "GVI"},
    "longbeachbuys.buyspeed.com": {"name": "Long Beach Buys", "prefix": "LB"},
    "procure.portlandoregon.gov": {"name": "Portland Procurement", "prefix": "PRT"},
    "knoxbuys.buyspeed.com": {"name": "Knox Buys", "prefix": "KNOX"},
    "oregon-uat.buyspeed.com": {"name": "Oregon UAT", "prefix": "OR_UAT"},
    "www.njstart.gov": {"name": "NJ Start", "prefix": "NJ"},
    "epro.sbcounty.gov": {"name": "SB County", "prefix": "SBC"},
    "oregonbuys.gov": {"name": "Oregon Buys", "prefix": "OR"},
    "arbuy.arkansas.gov": {"name": "Arkansas Buy", "prefix": "AR"},
    "nevadaepro.com": {"name": "Nevada ePro", "prefix": "NV"},
    "www.nttamarketplace.org": {"name": "NTTA Marketplace", "prefix": "NTTA"},
    "www.phlcontracts.phila.gov": {"name": "PHL Contracts", "prefix": "PHL"},
    "alohaebuys-uat.hawaii.gov": {"name": "Hawaii eBuys", "prefix": "HI"},
    "www.commbuys.com": {"name": "COMMBUYS", "prefix": "CB"},
}


# Update the folder setup code
def setup_folders():
    """Set up the required folder structure."""
    yesterday = datetime.now() - timedelta(days=1)
    base_folder = os.path.join(
        os.path.dirname(os.path.dirname(__file__)), yesterday.strftime("%Y-%m-%d")
    )
    os.makedirs(base_folder, exist_ok=True)

    script_name = os.path.splitext(os.path.basename(__file__))[0]
    main_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
    os.makedirs(main_folder, exist_ok=True)

    # Create temporary download folder inside main_folder
    script_download_folder = os.path.join(main_folder, script_name)
    os.makedirs(script_download_folder, exist_ok=True)

    return base_folder, main_folder, script_download_folder


# Add this function to get site info
def get_site_info(url):
    """Get site name and prefix from URL."""
    domain = urlparse(url).netloc
    return SITES_CONFIG.get(domain, {"name": "Unknown", "prefix": "UNK"})


# Function to handle errors
def handle_error(error_message):
    """
    Handle errors by playing a notification sound and pausing the script.

    Args:
        error_message (str): The error message to display.
    """
    print(f"❗ Error: {error_message}")
    play_notification_sound()
    input("Press Enter to continue...")


def setup_driver():
    """Set up and configure the Chrome WebDriver."""
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": script_download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "pageLoadStrategy": "none",  # Add this to prevent waiting for full page load
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-browser-side-navigation")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-notifications")

    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(60)  # Increase timeout to 60 seconds
    driver.set_script_timeout(60)  # Add script timeout
    driver.implicitly_wait(20)  # Increase implicit wait
    return driver


def click_advanced_search(driver):
    """
    Perform advanced search on the website.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.

    Raises:
        Exception: If an error occurs during the advanced search process.
    """
    try:
        # Wait for the page to load completely
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.XPATH, "//legend[contains(text(), 'Advanced Search')]")
            )
        )

        # Scroll to the Advanced Search option
        advanced_search = driver.find_element(
            By.XPATH, "//legend[contains(text(), 'Advanced Search')]"
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", advanced_search)

        # Wait for the element to be clickable and click it
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//legend[contains(text(), 'Advanced Search')]")
            )
        )
        advanced_search.click()

        # Wait for the status dropdown to be present
        status_dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "bidSearchForm:status"))
        )

        # Select the "Bids/Solicitations" status
        status_dropdown.find_element(By.XPATH, "//option[@value='2BS']").click()

        # Wait for the search button to be clickable and click it
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "bidSearchForm:buttonSearch"))
        )
        search_button.click()

        # Wait for the search results to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//table[@role='grid']"))
        )

    except TimeoutException:
        print("Timeout occurred while waiting for elements to load.")
    except ElementClickInterceptedException:
        print(
            "Element click was intercepted. The page might not be in the expected state."
        )
    except Exception as e:
        print(f"Error in advanced search: {str(e)}")
        raise


def extract_bid_links(driver, max_links=50):
    """
    Extract bid links from the search results page.
    """
    bid_links = []
    page_number = 1
    retry_count = 0
    max_retries = 3

    while len(bid_links) < max_links and retry_count < max_retries:
        try:
            log_message(f"📄 Processing Page {page_number}")

            # Wait for page load
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//table[@role='grid']"))
            )

            # Give time for dynamic content to load
            time.sleep(3)

            # Find all bid rows
            rows = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (
                        By.XPATH,
                        "//tr[@class='ui-widget-content ui-datatable-even' or @class='ui-widget-content ui-datatable-odd']",
                    )
                )
            )

            for row in rows:
                try:
                    # Wait for and extract the bid link
                    link = (
                        WebDriverWait(row, 10)
                        .until(EC.presence_of_element_located((By.XPATH, ".//td[1]/a")))
                        .get_attribute("href")
                    )

                    if link not in bid_links:  # Avoid duplicates
                        bid_links.append(link)
                        log_message(f"🔗 Found bid link: {link}")

                    if len(bid_links) >= max_links:
                        return bid_links

                except StaleElementReferenceException:
                    continue

            # Try to find next page button
            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            "//a[contains(@class, 'ui-paginator-next') and not(contains(@class, 'ui-state-disabled'))]",
                        )
                    )
                )

                # Check if next button is actually clickable
                if "ui-state-disabled" not in next_button.get_attribute("class"):
                    # Scroll into view
                    driver.execute_script(
                        "arguments[0].scrollIntoView(true);", next_button
                    )
                    time.sleep(1)

                    # Try JavaScript click first
                    try:
                        driver.execute_script("arguments[0].click();", next_button)
                    except:
                        ActionChains(driver).move_to_element(
                            next_button
                        ).click().perform()

                    page_number += 1
                    time.sleep(3)  # Wait for page load
                    retry_count = 0  # Reset retry count on successful navigation
                else:
                    log_message("📄 Reached last page")
                    break
            except:
                retry_count += 1
                log_message(f"⚠️ Navigation retry {retry_count}/{max_retries}")
                if retry_count >= max_retries:
                    log_message("❌ Max navigation retries reached")
                    break
                time.sleep(2)

        except Exception as e:
            log_message(f"❌ Error processing page {page_number}: {str(e)}")
            retry_count += 1
            if retry_count >= max_retries:
                break
            time.sleep(2)

    return bid_links


def extract_bid_details(driver, bid_link):
    """
    Extract details from a specific bid page.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        bid_link (str): The URL of the bid detail page.

    Returns:
        dict: A dictionary containing the extracted bid details.
    """
    log_message(f"\n📄 Processing bid page: {bid_link}")
    driver.get(bid_link)

    bid_details = {}

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//table[@role='presentation']"))
        )

        # Define fields to extract and their corresponding keys
        fields_to_extract = [
            ("Available Date : ", "Posted Date"),
            ("Bid Opening Date:", "Response Date"),
            ("Bid Number:", "Solicitation Number"),
            ("Description:", "Solicitation Title"),
            ("Organization:", "Agency"),
            ("Bulletin Desc:", "Description"),
            ("Pre Bid Conference:", "Additional Summary, if any"),
            ("Ship-to Address:", "Contracting Office Address"),
        ]

        # Extract each field
        for field, key in fields_to_extract:
            try:
                if field == "Available Date : ":
                    value = driver.find_element(
                        By.XPATH,
                        f"//td[contains(text(), 'Available Date')]/following-sibling::td",
                    ).text.strip()
                else:
                    value = driver.find_element(
                        By.XPATH,
                        f"//td[contains(text(), '{field}')]/following-sibling::td",
                    ).text.strip()
                bid_details[key] = value
            except NoSuchElementException:
                bid_details[key] = ""

        # Extract UNSPSC code and category
        try:
            unspsc_element = driver.find_element(
                By.XPATH,
                "//td[contains(text(), 'U N S P S C Code:')]/following-sibling::td",
            )
            unspsc_code = unspsc_element.find_element(By.TAG_NAME, "u").text.strip()
            unspsc_text = unspsc_element.text.strip()
            bid_details["Category"] = f"{unspsc_code} - {unspsc_text}"
        except NoSuchElementException:
            bid_details["Category"] = ""

        # Set default values for some fields
        bid_details["Notice Type"] = ""
        bid_details["Contact Information"] = ""
        bid_details["Bid Detail Page URL"] = bid_link

        # Extract file and form attachments
        file_attachments = driver.find_elements(
            By.XPATH,
            "//td[contains(text(), 'File Attachments:')]/following-sibling::td//a",
        )
        form_attachments = driver.find_elements(
            By.XPATH,
            "//td[contains(text(), 'Form Attachments:')]/following-sibling::td//a",
        )
        all_attachments = [
            att.get_attribute("text") for att in file_attachments + form_attachments
        ]
        bid_details["Attachments"] = ", ".join(all_attachments)

        # Update date format for Posted Date and Response Date
        for date_field in ["Posted Date", "Response Date"]:
            if date_field in bid_details and bid_details[date_field] != "":
                try:
                    date_obj = datetime.strptime(
                        bid_details[date_field].strip(), "%m/%d/%Y %I:%M:%S %p"
                    )
                    bid_details[date_field] = date_obj.strftime("%Y-%m-%d")
                except ValueError:
                    log_message(
                        f"⚠️ Could not parse {date_field} for bid {bid_details.get('Solicitation Number', 'Unknown')}"
                    )

        # After extracting all details, print the key information
        log_message("\n📋 Bid Details:")
        log_message(f"🔢 Bid Number: {bid_details.get('Solicitation Number', 'N/A')}")
        log_message(f"📅 Posted Date: {bid_details.get('Posted Date', 'N/A')}")
        log_message(f"⏰ Response Date: {bid_details.get('Response Date', 'N/A')}")
        log_message(f"🏢 Agency: {bid_details.get('Agency', 'N/A')}")
        log_message(f"📝 Title: {bid_details.get('Solicitation Title', 'N/A')}")

        if bid_details.get("Attachments"):
            log_message(f"📎 Available Attachments: {bid_details['Attachments']}")

    except Exception as e:
        log_message(f"❌ Error extracting bid details: {str(e)}")
        return None

    return bid_details


def download_attachments(driver, bid_number):
    """Download attachments for a specific bid using ActionChains."""
    log_message(f"\n📥 Starting attachment downloads for bid {bid_number}")
    attachments_downloaded = False
    bid_folder = os.path.join(main_folder, bid_number.replace("/", "_"))

    attachments = driver.find_elements(
        By.XPATH,
        "//a[contains(@href, 'javascript:downloadFile') or contains(@href, 'javascript:downloadForm')]",
    )

    if not attachments:
        log_message("ℹ️ No attachments found for this bid")
        return attachments_downloaded, ""

    os.makedirs(bid_folder, exist_ok=True)
    log_message(f"📁 Created folder for attachments: {bid_folder}")

    downloaded_files = []
    total_attachments = len(attachments)

    log_message(f"📎 Found {total_attachments} attachments to download")

    for index, attachment in enumerate(attachments, start=1):
        file_name = attachment.text
        log_message(f"\n⬇️ [{index}/{total_attachments}] Downloading: {file_name}")

        try:
            before = os.listdir(script_download_folder)
            log_message(f"🔄 Initiating download...")

            ActionChains(driver).move_to_element(attachment).click().perform()

            # Increase wait time for downloads
            wait_time = max(60, min(120, 20 * len(attachments)))

            # Wait for the file to be downloaded
            downloaded_file = None
            for _ in range(wait_time):
                time.sleep(1)
                after = os.listdir(script_download_folder)
                new_files = [f for f in after if f not in before]
                if new_files:
                    downloaded_file = new_files[0]
                    if not downloaded_file.endswith(".crdownload"):
                        break

            if not downloaded_file:
                log_message(f"⚠️ Download timeout for {file_name}")
                continue

            log_message(f"✅ Download completed: {downloaded_file}")

            # Move file to bid folder
            source_path = os.path.join(script_download_folder, downloaded_file)
            destination_path = os.path.join(bid_folder, downloaded_file)

            log_message(f"📦 Moving file to bid folder...")
            for attempt in range(5):
                try:
                    safe_move(source_path, destination_path)
                    attachments_downloaded = True
                    downloaded_files.append(downloaded_file)
                    log_message(f"✅ Successfully moved file to: {destination_path}")
                    break
                except PermissionError:
                    log_message(f"⚠️ Move attempt {attempt + 1}/5 failed, retrying...")
                    time.sleep(1)
            else:
                log_message(f"❌ Failed to move {downloaded_file} after 5 attempts")

        except Exception as e:
            log_message(f"❌ Error downloading attachment: {str(e)}")

    log_message(f"\n📊 Download Summary for bid {bid_number}:")
    log_message(f"Total files attempted: {total_attachments}")
    log_message(f"Successfully downloaded: {len(downloaded_files)}")
    if downloaded_files:
        log_message("Downloaded files:")
        for file in downloaded_files:
            log_message(f"  - {file}")

    return attachments_downloaded, ", ".join(downloaded_files)


def update_excel(bid_data, sl_no):
    """Update the Excel file with new bid data."""
    try:
        bid_data["SL No"] = sl_no
        excel_file = os.path.join(main_folder, f"{script_name}.xlsx")

        log_message(f"\n💾 Saving bid to Excel:")
        log_message(f"📊 File: {script_name}.xlsx")
        log_message(f"🔢 Bid Number: {bid_data.get('Solicitation Number')}")

        if os.path.exists(excel_file):
            existing_df = pd.read_excel(excel_file)
            updated_df = pd.concat(
                [existing_df, pd.DataFrame([bid_data])], ignore_index=True
            )
            log_message(f"✅ Added to existing Excel file")
        else:
            updated_df = pd.DataFrame([bid_data])
            log_message(f"✅ Created new Excel file")

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
            "Additional Summary, if any",
            "Contracting Office Address",
            "Contact Information",
            "Bid Detail Page URL",
            "Attachments",
        ]

        updated_df = updated_df.reindex(columns=column_order)
        for col in column_order:
            if col not in updated_df.columns:
                updated_df[col] = ""

        updated_df.to_excel(excel_file, index=False)
        log_message(f"✅ Successfully saved bid data to {excel_file}")

    except Exception as e:
        log_message(f"❌ Error updating Excel file: {str(e)}")
        raise


def move_remaining_files():
    """
    Move any remaining files in the main folder to their respective bid folders.
    """
    for file in os.listdir(main_folder):
        if file.endswith(".xlsx"):
            continue
        for bid_folder in os.listdir(main_folder):
            if os.path.isdir(os.path.join(main_folder, bid_folder)) and file.startswith(
                bid_folder
            ):
                source = os.path.join(main_folder, file)
                destination = os.path.join(main_folder, bid_folder, file)
                safe_move(source, destination)
                log_message(f"Moved {file} to {bid_folder}")
                break


def parse_arguments():
    """
    Parse command-line arguments.

    Returns:
        argparse.Namespace: Parsed arguments.
    """
    parser = argparse.ArgumentParser(description="COMMBUYS bid scraper")
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Number of days to look back for bids (default: 2)",
    )
    parser.add_argument(
        "--start-date",
        type=str,
        help="Start date in YYYY-MM-DD format (overrides --days if provided)",
    )
    parser.add_argument(
        "--end-date",
        type=str,
        help="End date in YYYY-MM-DD format (defaults to today if start-date is provided)",
    )
    return parser.parse_args()


def get_date_range(args):
    """
    Calculate the date range for bid filtering based on command line arguments.
    Returns start_date and end_date inclusive of the full days specified.
    """
    today = datetime.now().date()

    if args.start_date:
        try:
            start_date = datetime.strptime(args.start_date, "%Y-%m-%d").date()
            end_date = today
            if args.end_date:
                end_date = datetime.strptime(args.end_date, "%Y-%m-%d").date()
            return start_date, end_date
        except ValueError:
            log_message("⚠️ Invalid date format. Using default date range.")

    # Default behavior using --days
    end_date = today
    start_date = end_date - timedelta(days=args.days)  # Changed from (days-1) to days

    log_message(f"📅 Date range calculated: {start_date} to {end_date}")
    return start_date, end_date


def should_process_bid(posted_date_str, start_date, end_date):
    """
    Determine if a bid should be processed based on its posted date and the date range.
    """
    try:
        posted_date = datetime.strptime(posted_date_str, "%Y-%m-%d").date()
        return start_date <= posted_date <= end_date
    except ValueError:
        log_message(f"⚠️ Could not parse date: {posted_date_str}")
        return False


def clean_up_script_folder():
    """
    Clean up the script-specific download folder by moving all files to their respective bid folders.
    """
    for item in os.listdir(script_download_folder):
        item_path = os.path.join(script_download_folder, item)
        if os.path.isfile(item_path):
            # This is a file that wasn't moved to a bid folder
            log_message(f"⚠️ Found unmoved file: {item}")
            # You can decide what to do with these files (e.g., move to a 'misc' folder or delete)


def signal_handler(signum, frame):
    """Handle Ctrl+C by stopping all processing and closing drivers."""
    log_message("\n\n🛑 Ctrl+C detected. Initiating graceful shutdown...")

    # Set the stop event to signal all threads to stop
    stop_processing.set()

    # Close all active browser instances
    log_message("🌐 Closing all browser instances...")
    for driver in active_drivers:
        try:
            driver.quit()
            log_message("✅ Closed a browser instance")
        except Exception as e:
            log_message(f"⚠️ Error closing browser: {str(e)}")

    # Shutdown thread pool if it exists
    if thread_pool:
        log_message("🧵 Shutting down thread pool...")
        thread_pool.shutdown(wait=False)

    log_message("👋 Shutdown complete. Exiting...")
    sys.exit(0)


def process_site(url, args):
    """Process a single procurement site."""
    site_info = get_site_info(url)
    log_message(f"\n{'='*80}")
    log_message(f"🌐 Starting process for {site_info['name']} ({url})")
    log_message(f"{'='*80}")

    driver = setup_driver()
    active_drivers.append(driver)

    try:
        driver.get(url)
        log_message(f"✅ Successfully loaded {site_info['name']}")

        if stop_processing.is_set():
            log_message(f"🛑 Stopping processing of {site_info['name']}")
            return

        # Perform advanced search
        click_advanced_search(driver)
        log_message(f"✅ Completed advanced search setup for {site_info['name']}")

        processed_bids = 0
        sl_no = 1
        cutoff_date = datetime.now() - timedelta(days=args.days)

        # Extract bid links
        bid_links = extract_bid_links(driver, max_links=50)
        total_bids = len(bid_links)
        log_message(f"\n📊 Found {total_bids} bids for {site_info['name']}")

        # Process each bid
        for index, link in enumerate(bid_links, start=1):
            if stop_processing.is_set():
                log_message(f"🛑 Stopping bid processing for {site_info['name']}")
                break

            log_message(f"\n{'='*50}")
            log_message(f"🔍 Processing {site_info['name']} bid {index}/{total_bids}")

            bid_details = extract_bid_details(driver, link)
            if bid_details is None or "Solicitation Number" not in bid_details:
                continue

            if "Posted Date" in bid_details and bid_details["Posted Date"] != "N/A":
                try:
                    posted_date = datetime.strptime(
                        bid_details["Posted Date"].strip(),
                        "%Y-%m-%d",
                    )
                    if posted_date > cutoff_date:
                        # Download attachments
                        attachments_downloaded, downloaded_files = download_attachments(
                            driver,
                            f"{site_info['prefix']}_{bid_details['Solicitation Number']}",
                        )
                        if attachments_downloaded:
                            bid_details["Attachments"] = downloaded_files
                        # Update Excel file
                        update_excel(bid_details, sl_no)
                        processed_bids += 1
                        sl_no += 1
                    else:
                        log_message(
                            f"⏭️ Skipping bid {bid_details['Solicitation Number']} - Posted Date: {posted_date}"
                        )
                except ValueError:
                    log_message(
                        f"⚠️ Could not parse date for bid {bid_details['Solicitation Number']}"
                    )
            else:
                log_message(
                    f"⚠️ No Posted Date found for bid {bid_details.get('Solicitation Number', 'Unknown')}"
                )

        log_message(f"\n📊 {site_info['name']} Summary:")
        log_message(f"Total bids found: {total_bids}")
        log_message(f"Bids processed: {processed_bids}")

    except Exception as e:
        log_message(f"❌ Error processing {site_info['name']}: {str(e)}")
        raise
    finally:
        try:
            driver.quit()
            active_drivers.remove(driver)
        except:
            pass


def process_site_with_retry(url, args, max_retries=3):
    """Process a site with retry mechanism."""
    site_info = get_site_info(url)
    retry_count = 0

    while retry_count < max_retries:
        try:
            return process_site(url, args)
        except Exception as e:
            retry_count += 1
            if retry_count >= max_retries:
                log_message(
                    f"❌ Failed to process {site_info['name']} after {max_retries} attempts"
                )
                raise
            log_message(f"⚠️ Retry {retry_count}/{max_retries} for {site_info['name']}")
            time.sleep(5 * retry_count)  # Exponential backoff


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


def save_to_cache(bid_link, bid_details):
    """Save bid details to cache."""
    # Create cache directory if it doesn't exist
    cache_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "cache")
    os.makedirs(cache_dir, exist_ok=True)

    cache_file = os.path.join(
        cache_dir, f"{os.path.splitext(os.path.basename(__file__))[0]}_cache.json"
    )

    cache = load_cache()

    # Add/update the bid in cache
    cache[bid_link] = {
        "posted_date": bid_details.get("Posted Date", ""),
        "last_checked": datetime.now().strftime("%Y-%m-%d"),
    }

    # Remove bids older than 3 months
    three_months_ago = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
    cache = {
        k: v for k, v in cache.items() if v.get("posted_date", "") >= three_months_ago
    }

    # Save updated cache
    with open(cache_file, "w") as f:
        json.dump(cache, f, indent=2)


def should_process_bid_link(bid_link):
    """Check if bid should be processed based on cache."""
    cache = load_cache()
    if bid_link in cache:
        bid_data = cache[bid_link]
        log_message(
            f"⏭️ Skipping previously processed bid from {bid_data['posted_date']}"
        )
        return False
    return True


def process_bid(bid_details, site_info, cutoff_date, sl_no, driver):
    """Process a single bid."""
    if bid_details is None or "Solicitation Number" not in bid_details:
        return False

    bid_link = bid_details.get("Bid Detail Page URL")
    if not should_process_bid_link(bid_link):
        return False

    if "Posted Date" in bid_details and bid_details["Posted Date"] != "N/A":
        try:
            posted_date = datetime.strptime(
                bid_details["Posted Date"].strip(),
                "%Y-%m-%d",
            )
            if posted_date.date() >= cutoff_date.date():
                log_message(
                    f"✅ Processing bid from {posted_date.strftime('%Y-%m-%d')}"
                )

                # Download attachments
                attachments_downloaded, downloaded_files = download_attachments(
                    driver,
                    f"{site_info['prefix']}_{bid_details['Solicitation Number']}",
                )
                if attachments_downloaded:
                    bid_details["Attachments"] = downloaded_files
                    log_message(f"✅ Downloaded attachments: {downloaded_files}")

                # Update Excel file
                update_excel(bid_details, sl_no)
                log_message(
                    f"✅ Saved bid {bid_details['Solicitation Number']} to Excel"
                )

                # Save to cache
                save_to_cache(bid_link, bid_details)

                return True
            else:
                log_message(
                    f"⏭️ Skipping bid {bid_details['Solicitation Number']} - Posted Date: {posted_date.strftime('%Y-%m-%d')} is before cutoff date {cutoff_date.strftime('%Y-%m-%d')}"
                )
        except ValueError:
            log_message(
                f"⚠️ Could not parse date for bid {bid_details['Solicitation Number']}"
            )
    else:
        log_message(
            f"⚠️ No Posted Date found for bid {bid_details.get('Solicitation Number', 'Unknown')}"
        )

    return False


def log_message(message):
    """Thread-safe logging function."""
    with log_lock:
        print(message, flush=True)


def load_page_with_retry(driver, url, max_retries=3):
    """Load a page with retry mechanism."""
    for attempt in range(max_retries):
        try:
            driver.get(url)
            # Wait for document ready state
            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            return True
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            log_message(f"⚠️ Page load attempt {attempt + 1} failed, retrying...")
            time.sleep(5 * (attempt + 1))
            try:
                driver.execute_script("window.stop();")
            except:
                pass
    return False


def main():
    """Main function to execute the bid extraction process for multiple sites."""
    global base_folder, main_folder, script_download_folder, script_name

    # Setup folders
    base_folder, main_folder, script_download_folder = setup_folders()
    script_name = os.path.splitext(os.path.basename(__file__))[0]

    args = parse_arguments()
    start_date, end_date = get_date_range(args)

    log_message(f"🚀 Multi-Site Bids Extraction Started")
    log_message(f"📅 Looking for bids from {start_date} to {end_date}")
    log_message(f"📁 Saving data in folder: {main_folder}")
    log_message(f"📁 Temporary download folder: {script_download_folder}")

    # Keep all URLs
    urls = [
        "https://www.commbuys.com/bso/view/search/external/advancedSearchBid.xhtml?openBids=true",
        "https://gvibuy.buyspeed.com/bso/view/search/external/advancedSearchBid.xhtml",
        "https://longbeachbuys.buyspeed.com/bso/view/search/external/advancedSearchBid.xhtml?openBids=true",
        "https://procure.portlandoregon.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true",
        "https://knoxbuys.buyspeed.com/bso/view/search/external/advancedSearchBid.xhtml?openBids=true",
        "https://oregon-uat.buyspeed.com/bso/view/search/external/advancedSearchBid.xhtml?openBids=true",
        "https://www.njstart.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true",
        "https://epro.sbcounty.gov/bso/view/search/external/advancedSearchBid.xhtml",
    ]

    sl_no = 1  # Global counter for all bids across all sites

    for url in urls:
        driver = None
        try:
            driver = setup_driver()
            driver.get(url)
            log_message(f"✅ Successfully loaded site: {url}")

            # Add delay for page stabilization
            time.sleep(10)

            try:
                # Perform advanced search
                click_advanced_search(driver)
                log_message(f"✅ Completed advanced search")

                # Extract bid links
                bid_links = extract_bid_links(driver, max_links=50)
                total_bids = len(bid_links)
                log_message(f"\n📊 Found {total_bids} bids")

                # Process each bid
                for index, link in enumerate(bid_links, start=1):
                    try:
                        log_message(f"\n🔍 Processing bid {index}/{total_bids}")
                        log_message(f"🔗 URL: {link}")

                        # Check cache before processing
                        if not should_process_bid_link(link):
                            continue

                        bid_details = extract_bid_details(driver, link)
                        if (
                            bid_details is None
                            or "Solicitation Number" not in bid_details
                        ):
                            continue

                        if (
                            "Posted Date" in bid_details
                            and bid_details["Posted Date"] != "N/A"
                        ):
                            if should_process_bid(
                                bid_details["Posted Date"], start_date, end_date
                            ):
                                log_message(
                                    f"✅ Bid qualifies for processing - Posted Date: {bid_details['Posted Date']}"
                                )

                                # Check for attachments before downloading
                                attachments = driver.find_elements(
                                    By.XPATH,
                                    "//a[contains(@href, 'javascript:downloadFile') or contains(@href, 'javascript:downloadForm')]",
                                )

                                if attachments:
                                    attachment_names = [att.text for att in attachments]
                                    log_message(
                                        f"📎 Found {len(attachments)} attachments:"
                                    )
                                    for att in attachment_names:
                                        log_message(f"   - {att}")

                                    # Download attachments
                                    attachments_downloaded, downloaded_files = (
                                        download_attachments(
                                            driver, bid_details["Solicitation Number"]
                                        )
                                    )
                                    if attachments_downloaded:
                                        bid_details["Attachments"] = downloaded_files
                                        log_message(
                                            f"✅ Successfully downloaded attachments"
                                        )
                                else:
                                    log_message("ℹ️ No attachments found for this bid")
                                    bid_details["Attachments"] = ""

                                # Update Excel file
                                log_message("💾 Saving bid to Excel...")
                                update_excel(bid_details, sl_no)
                                log_message(
                                    f"✅ Successfully saved bid {bid_details['Solicitation Number']} to Excel"
                                )

                                # Save to cache after successful processing
                                bid_details["Bid Detail Page URL"] = link
                                save_to_cache(link, bid_details)

                                sl_no += 1
                            else:
                                log_message(
                                    f"⏭️ Skipping bid - Posted Date: {bid_details['Posted Date']} (outside date range {start_date} to {end_date})"
                                )

                    except Exception as e:
                        log_message(f"❌ Error processing bid: {str(e)}")
                        continue

            except Exception as e:
                log_message(f"❌ Error in search process: {str(e)}")

        except Exception as e:
            log_message(f"❌ Error processing site: {str(e)}")
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass

    # Final cleanup
    try:
        shutil.rmtree(script_download_folder)
        log_message("✅ Removed temporary download folder")
    except Exception as e:
        log_message(f"⚠️ Error removing temporary folder: {str(e)}")

    log_message(f"\n💾 Final results saved to {script_name}.xlsx")

    # Mark folder as completed
    try:
        completed_folder = os.path.join(base_folder, f"{script_name}_COMPLETED")
        os.rename(main_folder, completed_folder)
        main_folder = completed_folder  # Update the reference
        log_message(f"✅ Renamed folder to indicate completion: {completed_folder}")
    except Exception as e:
        log_message(f"⚠️ Error renaming folder: {str(e)}")

    log_message("\n Multi-Site Bids Extraction Completed")
    play_notification_sound()


if __name__ == "__main__":
    main()
