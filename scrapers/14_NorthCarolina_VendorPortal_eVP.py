import logging
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import functools

# Override the print function to always flush
print = functools.partial(print, flush=True)

import argparse
import os

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import time
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent
from selenium_stealth import stealth
import shutil
import random
from selenium.webdriver.common.keys import Keys
import winsound
from selenium.webdriver.common.action_chains import ActionChains
from pathlib import Path
import re
from utils.utils import safe_move, play_notification_sound
import tempfile

# Set up logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Initialize argument parser
parser = argparse.ArgumentParser(description="North Carolina Vendor Portal Scraper")
parser.add_argument(
    "--days",
    type=int,
    default=2,
    help="Number of days to look back for bid posted dates (default: 2, maximum: 30)",
)
args = parser.parse_args()
days = args.days if args.days <= 30 else 30

# Update the folder structure and file path definitions
SCRIPT_NAME = "14_NorthCarolina_VendorPortal_eVP"
yesterday = datetime.now() - timedelta(days=1)
main_folder = os.path.join(os.getcwd(), yesterday.strftime("%Y-%m-%d"))
os.makedirs(main_folder, exist_ok=True)

# Create the IN_PROGRESS folder - this will be our working directory
working_folder = os.path.join(main_folder, f"{SCRIPT_NAME}_IN_PROGRESS")
os.makedirs(working_folder, exist_ok=True)

# Set the temporary download folder as a subfolder of working_folder
TEMP_DOWNLOAD_FOLDER = os.path.join(working_folder, SCRIPT_NAME)
os.makedirs(TEMP_DOWNLOAD_FOLDER, exist_ok=True)

# Update the excel_filename to be directly in the working_folder
excel_filename = os.path.join(working_folder, f"{SCRIPT_NAME}.xlsx")

# Configure Selenium WebDriver with anti-bot measures
options = Options()
options.headless = False  # Change to True for headless mode
ua = UserAgent()
options.add_argument(f"user-agent={ua.random}")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-extensions")
options.add_argument("--no-sandbox")
options.add_argument("--disable-gpu")  # Disables GPU acceleration
options.add_argument("--disable-software-rasterizer")  # Use CPU rendering
options.add_argument("--disable-dev-shm-usage")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--page-load-strategy=normal")
options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": TEMP_DOWNLOAD_FOLDER,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    },
)

# Update the WebDriver initialization
driver = webdriver.Chrome(options=options)
driver.set_page_load_timeout(30)

# Apply selenium-stealth
stealth(
    driver,
    languages=["en-US", "en"],
    vendor="Google Inc.",
    platform="Win32",
    webgl_vendor="Intel Inc.",
    renderer="Intel Iris OpenGL Engine",
    fix_hairline=True,
)

# Add this after other global variables
DOWNLOADS_FOLDER = str(Path.home() / "Downloads")


def play_notification_sound(message):
    logger.info(message)
    play_notification_sound(message)


def pause_script(message):
    play_notification_sound(message)
    input("Press Enter to continue...")


def find_file_in_downloads(file_name):
    downloads_dir = Path(TEMP_DOWNLOAD_FOLDER)
    logger.info(f"Searching for {file_name} in {downloads_dir}")

    # Remove file size information, trim spaces, and replace multiple spaces with single space
    cleaned_name = re.sub(r"\s*\([^)]*\)\s*$", "", file_name).strip()
    cleaned_name = re.sub(r"\s+", " ", cleaned_name).lower()

    for file in downloads_dir.glob("*"):
        cleaned_file_name = re.sub(r"\s*\([^)]*\)\s*$", "", file.name).strip()
        cleaned_file_name = re.sub(r"\s+", " ", cleaned_file_name).lower()

        if cleaned_file_name == cleaned_name:
            logger.info(f"Found exact matching file: {file.name}")
            return str(file)
        elif cleaned_file_name.startswith(cleaned_name):
            logger.info(f"Found partial matching file: {file.name}")
            return str(file)

    logger.warning(f"No matching file found for {file_name}")
    return None


def wait_for_download_completion(file_path, timeout=300):  # 5 minutes timeout
    start_time = time.time()
    while time.time() - start_time < timeout:
        if file_path.endswith(".crdownload"):
            actual_file = file_path[:-11]  # Remove '.crdownload' extension
            if os.path.exists(actual_file):
                logger.info(f"Download completed: {actual_file}")
                return actual_file
        elif os.path.exists(file_path):
            logger.info(f"File already exists: {file_path}")
            return file_path

        time.sleep(5)  # Check every 5 seconds

    logger.warning(f"Download did not complete within {timeout} seconds: {file_path}")
    return None


def download_attachment(attachment_element):
    try:
        actions = ActionChains(driver)
        actions.move_to_element(attachment_element).click().perform()

        file_name = attachment_element.text.split(" (")[0].strip()
        logger.info(f"Attempting to download: {file_name}")

        max_wait_time = 600  # 10 minutes total
        check_interval = 10
        total_wait_time = 0

        while total_wait_time < max_wait_time:
            time.sleep(check_interval)
            total_wait_time += check_interval

            downloads_source = find_file_in_downloads(file_name)
            if downloads_source:
                if downloads_source.endswith(".crdownload"):
                    logger.info(f"Found partially downloaded file: {downloads_source}")
                    completed_file = wait_for_download_completion(downloads_source)
                    if completed_file:
                        actual_file_name = os.path.basename(completed_file)
                        logger.info(f"Download completed: {actual_file_name}")
                        return actual_file_name, completed_file
                else:
                    actual_file_name = os.path.basename(downloads_source)
                    logger.info(
                        f"Found matching file in Downloads folder: {actual_file_name}"
                    )
                    return actual_file_name, downloads_source

            logger.info(
                f"Attachment not found. Retrying... (Time elapsed: {total_wait_time} seconds)"
            )

        raise Exception(f"Download timeout after {max_wait_time} seconds")

    except Exception as e:
        logger.error(f"Failed to download attachment: {e}", exc_info=True)
        play_notification_sound(f"Failed to download {file_name}")
        pause_script("Paused due to download error.")
        return None, None


def move_attachments(downloaded_attachments, bid_folder):
    moved_attachments = []
    for file_name, source_path in downloaded_attachments:
        if source_path and os.path.exists(source_path):
            destination = os.path.join(bid_folder, file_name)
            try:
                safe_move(source_path, destination)
                moved_attachments.append(file_name)
                logger.info(f"Moved {file_name} to {bid_folder} successfully.")
            except Exception as move_error:
                logger.error(
                    f"Failed to move file {file_name}: {move_error}", exc_info=True
                )
        else:
            logger.warning(f"Source file not found: {source_path}")
    return moved_attachments


def move_remaining():
    logger.info(f"Moving remaining attachments from {TEMP_DOWNLOAD_FOLDER}")
    for file in os.listdir(TEMP_DOWNLOAD_FOLDER):
        if file.endswith((".pdf", ".docx", ".xlsx", ".zip")):
            for bid_folder in os.listdir(working_folder):
                bid_path = os.path.join(working_folder, bid_folder)
                if os.path.isdir(bid_path) and bid_folder != SCRIPT_NAME:
                    destination = os.path.join(bid_path, file)
                    source = os.path.join(TEMP_DOWNLOAD_FOLDER, file)
                    try:
                        safe_move(source, destination)
                        logger.info(f"Moved {file} to {bid_path}.")
                        break
                    except Exception as e:
                        logger.error(f"Failed to move {file}: {str(e)}")
    logger.info("Finished moving remaining attachments.")


def human_like_interaction():
    logger.info("Performing human-like interaction")
    body = driver.find_element(By.TAG_NAME, "body")
    for _ in range(random.randint(1, 3)):
        body.send_keys(Keys.ARROW_DOWN)
        time.sleep(random.uniform(0.5, 1.5))
    driver.execute_script("window.scrollBy(0, window.innerHeight * 0.5);")
    time.sleep(random.uniform(1, 3))


def parse_date(date_string):
    """Parse date string and return in YYYY-MM-DD format."""
    try:
        # Try parsing as M/D/YYYY h:mm A
        return datetime.strptime(date_string, "%m/%d/%Y %I:%M %p").strftime("%Y-%m-%d")
    except ValueError:
        try:
            # Try parsing as M/D/YYYY
            return datetime.strptime(date_string, "%m/%d/%Y").strftime("%Y-%m-%d")
        except ValueError:
            try:
                # Try parsing as YYYY-MM-DD
                return datetime.strptime(date_string, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                # If all fail, return original string
                return date_string


def get_bid_links(date_threshold, count=30):
    all_bid_links = []
    current_page = 1

    while len(all_bid_links) < count and current_page <= 3:
        logger.info(f"Scanning page {current_page} for bid links")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )

        bids = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        for bid in bids:
            try:
                posted_date_element = bid.find_element(
                    By.CSS_SELECTOR, "td:nth-child(5) time"
                )
                posted_date_str = posted_date_element.get_attribute("datetime")
                posted_date = datetime.strptime(posted_date_str, "%Y-%m-%d")
                if posted_date >= date_threshold:
                    link = bid.find_element(
                        By.CSS_SELECTOR, "td:nth-child(1) a"
                    ).get_attribute("href")
                    all_bid_links.append((link, posted_date_str))
                    logger.info(f"Added bid link: {link} (Posted: {posted_date_str})")
                    if len(all_bid_links) >= count:
                        return all_bid_links
            except NoSuchElementException:
                logger.warning(
                    "Failed to extract date or link for a bid", exc_info=True
                )
                continue

        if current_page < 3:
            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a.entity-pager-next-link")
                    )
                )
                next_button.click()
                current_page += 1
                logger.info(f"Moving to page {current_page}")
                time.sleep(random.uniform(2, 4))
                human_like_interaction()
            except (NoSuchElementException, TimeoutException):
                logger.info("No more pages to scan")
                break
        else:
            break

    return all_bid_links


def process_bid_links(bid_links):
    logger.info(f"Processing {len(bid_links)} bid links")
    for idx, (link, posted_date) in enumerate(bid_links, start=1):
        logger.info(f"Processing bid {idx}/{len(bid_links)}: {link}")
        max_retries = 3
        for attempt in range(max_retries):
            try:
                driver.get(link)
                # Increase the wait time for the page to load
                WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.ID, "evp_solicitationnbr"))
                )

                human_like_interaction()

                solicitation_number = driver.find_element(
                    By.ID, "evp_solicitationnbr"
                ).get_attribute("value")
                response_date_str = driver.find_element(
                    By.ID, "evp_opendate_datepicker_description"
                ).get_attribute("value")
                response_date = parse_date(response_date_str)
                solicitation_title = driver.find_element(
                    By.CSS_SELECTOR, "div.page-header h1"
                ).text
                agency = driver.find_element(
                    By.ID, "owningbusinessunit_name"
                ).get_attribute("value")
                category = driver.find_element(
                    By.ID, "evp_commcode_name"
                ).get_attribute("value")
                description = driver.find_element(By.ID, "evp_description").text
                additional_summary = driver.find_element(
                    By.ID, "evp_mandatoryconf_datepicker_description"
                ).get_attribute("value")
                contact_info = driver.find_element(
                    By.ID, "evp_specinstr"
                ).get_attribute("value")

                # Create the new additional_info field with line breaks
                additional_info_parts = []
                if additional_summary:
                    additional_info_parts.append(
                        f"Mandatory Conference/Site Visit:\n{additional_summary}"
                    )
                if contact_info:
                    additional_info_parts.append(
                        f"Special Instructions:\n{contact_info}"
                    )
                additional_info = "\n\n".join(
                    additional_info_parts
                )  # Add double line break between sections

                # Combine additional_summary with additional_info
                if additional_info:
                    additional_summary = (
                        f"{additional_summary}\n\n{additional_info}"
                        if additional_summary
                        else additional_info
                    )

                bid_detail_url = link

                logger.info(f"Extracted details for bid {solicitation_number}")

                attachment_xpath = "/html/body/form/div[4]/div[1]/div[3]/div/div[2]/div[2]/div/div/fieldset/table/tbody/tr[7]/td[1]/div[2]/div/div/div[5]/div/div/div[2]/div[2]/div/a"
                attachment_elements = driver.find_elements(By.XPATH, attachment_xpath)
                logger.info(f"Found {len(attachment_elements)} attachments")

                # Create bid folder inside working_folder instead of main_folder
                bid_folder = os.path.join(working_folder, solicitation_number)
                os.makedirs(bid_folder, exist_ok=True)

                downloaded_attachments = []
                for attachment in attachment_elements:
                    attachment_name, source_path = download_attachment(attachment)
                    if attachment_name and source_path:
                        downloaded_attachments.append((attachment_name, source_path))
                    time.sleep(random.uniform(2, 4))

                moved_attachments = move_attachments(downloaded_attachments, bid_folder)
                attachments_str = ", ".join(moved_attachments)

                bid_detail = {
                    "SL No": idx,
                    "Posted Date": parse_date(posted_date),
                    "Response Date": response_date,
                    "Notice Type": "",
                    "Solicitation Number": solicitation_number,
                    "Solicitation Title": solicitation_title,
                    "Agency": agency,
                    "Category": category,
                    "Description": description,
                    "Additional Summary": additional_summary,
                    "Contracting Office Address": "",
                    "Contact Information": "",  # Set to empty string
                    "Bid Detail Page URL": bid_detail_url,
                    "Attachments": attachments_str,
                }

                df = pd.DataFrame([bid_detail])
                if not os.path.exists(excel_filename):
                    df.to_excel(excel_filename, index=False, engine="openpyxl")
                else:
                    with pd.ExcelWriter(
                        excel_filename,
                        engine="openpyxl",
                        mode="a",
                        if_sheet_exists="overlay",
                    ) as writer:
                        df.to_excel(
                            writer,
                            index=False,
                            header=False,
                            startrow=writer.sheets["Sheet1"].max_row,
                        )

                logger.info(
                    f"✅ Bid {solicitation_number} successfully processed and saved to Excel."
                )

                break  # If successful, break out of the retry loop
            except TimeoutException:
                if attempt < max_retries - 1:
                    logger.warning(
                        f"Timeout on attempt {attempt + 1} for bid {link}. Retrying..."
                    )
                    time.sleep(random.uniform(5, 10))  # Wait before retrying
                else:
                    logger.error(
                        f"Failed to load bid {link} after {max_retries} attempts"
                    )
                    break  # Move to the next bid
            except Exception as e:
                logger.error(f"Error processing bid {link}: {e}", exc_info=True)
                play_notification_sound(f"Error processing bid {link}")
                pause_script("Paused due to error.")
                break  # Move to the next bid


def update_attachments_column(excel_path, working_folder):
    df = pd.read_excel(excel_path)
    for index, row in df.iterrows():
        solicitation_number = row["Solicitation Number"]
        bid_folder = os.path.join(working_folder, solicitation_number)
        if os.path.exists(bid_folder):
            attachments = [
                f
                for f in os.listdir(bid_folder)
                if os.path.isfile(os.path.join(bid_folder, f))
            ]
            df.at[index, "Attachments"] = ", ".join(attachments)
    df.to_excel(excel_path, index=False, engine="openpyxl")
    logger.info("✅ Attachments column updated based on actual files in bid folders.")


def cleanup_temp_folder():
    try:
        for item in os.listdir(TEMP_DOWNLOAD_FOLDER):
            item_path = os.path.join(TEMP_DOWNLOAD_FOLDER, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
        logger.info(f"Cleaned up temporary download folder: {TEMP_DOWNLOAD_FOLDER}")
    except Exception as e:
        logger.error(f"Failed to clean up temporary download folder: {e}")


def main():
    try:
        logger.info("🟢 Login Successful - Bids Extraction Started")

        driver.set_page_load_timeout(60)  # Increase timeout to 60 seconds

        max_retries = 3
        for attempt in range(max_retries):
            try:
                driver.get("https://evp.nc.gov/solicitations")
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "#EntityListControl")
                    )
                )
                logger.info("Main page loaded successfully")
                break
            except Exception as e:
                logger.error(f"Attempt {attempt + 1} failed: {e}", exc_info=True)
                if attempt == max_retries - 1:
                    raise
                time.sleep(5)

        human_like_interaction()

        date_threshold = datetime.now().replace(
            hour=0, minute=0, second=0, microsecond=0
        ) - timedelta(days=days)
        logger.info(f"Date threshold set to: {date_threshold.strftime('%Y-%m-%d')}")

        while True:
            bid_links = get_bid_links(date_threshold, count=30)
            if not bid_links:
                logger.info("No more bids found within the date range")
                break

            process_bid_links(bid_links)
            move_remaining()
            update_attachments_column(excel_filename, working_folder)

            logger.info("Processed 30 bids. Moving to the next batch...")

            # Move to the next set of 3 pages
            for _ in range(3):
                try:
                    next_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, "a.entity-pager-next-link")
                        )
                    )
                    next_button.click()
                    logger.info("Moving to next page")
                    time.sleep(random.uniform(2, 4))
                    human_like_interaction()
                except (NoSuchElementException, TimeoutException):
                    logger.info("No more pages to scan")
                    return  # Exit the function if there are no more pages

        logger.info("🎉 All Bids and Attachments Extraction Successfully Completed.")

        # Rename folder to indicate completion
        completed_folder = os.path.join(main_folder, f"{SCRIPT_NAME}_COMPLETED")
        try:
            if os.path.exists(TEMP_DOWNLOAD_FOLDER):
                os.rename(TEMP_DOWNLOAD_FOLDER, completed_folder)
                logger.info(f"Renamed folder to {completed_folder}")
        except Exception as e:
            logger.error(f"Failed to rename folder: {e}")

        play_notification_sound("Scraping process completed successfully.")

    except Exception as e:
        logger.error(f"An error occurred: {e}", exc_info=True)
        play_notification_sound("An error occurred during execution.")
        pause_script("Paused due to unexpected error.")
    finally:
        cleanup_temp_folder()
        driver.quit()


if __name__ == "__main__":
    main()
