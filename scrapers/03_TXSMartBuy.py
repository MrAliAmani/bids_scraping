import sys
import io
import functools
import json
from datetime import datetime, timedelta

# Override the print function to always flush
print = functools.partial(print, flush=True)

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import os

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.utils import safe_move, play_notification_sound
import time
import random
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import requests
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import shutil
import winsound  # For playing notification sounds
import sys

# Get the script name without creating a directory
script_name = os.path.splitext(os.path.basename(__file__))[0]

# Create a folder named after yesterday's date
yesterday = datetime.now() - timedelta(days=1)
main_folder = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), yesterday.strftime("%Y-%m-%d")
)
os.makedirs(main_folder, exist_ok=True)

# Create the script-specific folder with IN_PROGRESS suffix
script_folder = os.path.join(main_folder, f"{script_name}_IN_PROGRESS")
os.makedirs(script_folder, exist_ok=True)

# Create temporary download folder inside script folder
progress_folder = os.path.join(script_folder, script_name)
os.makedirs(progress_folder, exist_ok=True)


def setup_driver():
    """
    Set up and configure the Chrome WebDriver.

    Returns:
        webdriver.Chrome: Configured Chrome WebDriver instance.
    """
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")  # Disables GPU acceleration
    chrome_options.add_argument("--disable-software-rasterizer")  # Use CPU rendering
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("🌐 WebDriver setup complete")
    return driver


def random_sleep(min_seconds, max_seconds):
    """
    Sleep for a random duration between min_seconds and max_seconds.

    Args:
        min_seconds (float): Minimum sleep duration in seconds.
        max_seconds (float): Maximum sleep duration in seconds.
    """
    time.sleep(random.uniform(min_seconds, max_seconds))


def apply_filters(driver):
    """
    Apply filters on the search page.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.

    Raises:
        Exception: If an error occurs while applying filters.
    """
    print("🔍 Applying filters...")
    driver.get("https://www.txsmartbuy.gov/esbd?page=1&status=1")
    random_sleep(2, 4)

    try:
        # Set start date (2 days ago)
        start_date = (datetime.now() - timedelta(days=2)).strftime("%m/%d/%Y")
        start_date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "startDate"))
        )
        start_date_input.clear()
        start_date_input.send_keys(start_date)

        # Set end date (today)
        end_date = datetime.now().strftime("%m/%d/%Y")
        end_date_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "endDate"))
        )
        end_date_input.clear()
        end_date_input.send_keys(end_date)

        # Select "Posted" status
        status_select = Select(
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "status"))
            )
        )
        status_select.select_by_value("1")

        # Click search button
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.esbd-button[type='submit']")
            )
        )
        search_button.click()

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "esbd-result-title"))
        )

        print(f"📅 Start Date: {start_date}")
        print(f"📅 End Date: {end_date}")
        print("✅ Filters applied successfully")
    except Exception as e:
        print(f"Error applying filters: {str(e)}")
        driver.save_screenshot("apply_filters_error.png")
        raise


def extract_bid_links(driver, batch_size=30):
    """
    Extract bid links from the search results page.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        batch_size (int): Number of bids to extract in this batch.

    Returns:
        list: A list of extracted bid links.
    """
    print(f"🔗 Extracting bid links (batch size: {batch_size})...")
    all_links = []
    page = 1

    while len(all_links) < batch_size:
        print(f"📄 Current Page Number: {page}")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "esbd-result-title"))
        )

        soup = BeautifulSoup(driver.page_source, "html.parser")
        bid_links = soup.find_all("div", class_="esbd-result-title")

        for link in bid_links:
            all_links.append(link.find("a")["href"])
            if len(all_links) >= batch_size:
                break

        if len(all_links) >= batch_size:
            break

        next_page = soup.find("a", id="Next")
        if not next_page:
            break

        next_page_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "Next"))
        )
        driver.execute_script("arguments[0].click();", next_page_link)

        page += 1
        random_sleep(2, 4)

    print(f"✅ Extracted {len(all_links)} bid links")
    return all_links


def extract_bid_details(driver, bid_link, index, total_bids, max_retries=3):
    """
    Extract details from a specific bid page.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        bid_link (str): The URL of the bid details page.
        index (int): The index of the current bid.
        total_bids (int): The total number of bids to process.
        max_retries (int): Maximum number of retries for loading the page.

    Returns:
        tuple: A tuple containing the bid details, a boolean indicating if an Available Date was found,
               and the Available Date if found.
    """
    print(
        f"📄 Processing bid {index}/{total_bids} - URL: https://www.txsmartbuy.gov{bid_link}"
    )

    for attempt in range(max_retries):
        try:
            driver.get("https://www.txsmartbuy.gov" + bid_link)
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "esbd-result-cell"))
            )
            break
        except TimeoutException:
            if attempt < max_retries - 1:
                print(f"Timeout waiting for page to load for bid {index}. Retrying...")
                driver.refresh()
                random_sleep(2, 4)
            else:
                print(
                    f"Timeout waiting for page to load for bid {index} after {max_retries} attempts"
                )
                return None, False, None

    soup = BeautifulSoup(driver.page_source, "html.parser")

    details = {
        "Posted Date": "",
        "Response Date": "",
        "Notice Type": "",
        "Solicitation Number": "",
        "Solicitation Title": "",
        "Agency": "",
        "Category": "",
        "Description": "",
        "Additional Summary": "",
        "Contracting Office Address": "",
        "Contact Information": "",
        "Bid Detail Page URL": driver.current_url,
        "Attachments": [],
    }

    # Extract details
    for div in soup.find_all("div", class_="esbd-result-cell"):
        label = div.find("strong")
        if label:
            key = label.text.strip().rstrip(":")
            value = div.find("p").text.strip() if div.find("p") else ""

            if key == "Solicitation Posting Date":
                details["Posted Date"] = value
                print(f"   Posted Date: {value}")
            elif key == "Response Due Date":
                details["Response Date"] = value
                print(f"   Response Date: {value}")
            elif key == "Solicitation ID":
                details["Solicitation Number"] = value
                print(f"   Solicitation Number: {value}")
            elif key == "Agency/Texas SmartBuy Member Number":
                details["Agency"] = value
                print(f"   Agency: {value}")
            elif key == "Class/Item Code":
                details["Category"] = value
                print(f"   Category: {value}")
            elif key in ["Contact Name", "Contact Email", "Contact Number"]:
                details["Contact Information"] += f"{key}: {value}\n"
                print(f"   {key}: {value}")

    # Extract title
    title_div = soup.find("div", class_="esbd-result-title")
    if title_div:
        details["Solicitation Title"] = title_div.text.strip()
        print(f"   Title: {details['Solicitation Title']}")

    # Extract Description using multiple methods
    description = ""
    try:
        description_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    "#content > div > div > div.esbd-container > div.esbd-result-body-columns > div:nth-child(3) > p:nth-child(3) > span",
                )
            )
        )
        description = description_element.text.strip()
    except TimeoutException:
        print(f"Description not found with CSS selector for bid {index}")
        try:
            description_element = driver.find_element(
                By.XPATH,
                "//strong[contains(text(), 'Solicitation Description')]/following-sibling::p",
            )
            description = description_element.text.strip()
        except NoSuchElementException:
            print(f"Description not found with XPath for bid {index}")

    details["Description"] = description
    print(
        f"   Description: {description[:100]}..."
    )  # Print first 100 characters of description

    # Extract attachments
    attachments_div = soup.find_all("div", class_="esbd-attachment-row-content")
    details["Attachments"] = []
    details["_attachment_urls"] = []
    if attachments_div:
        for attachment_row in attachments_div:
            attachment_link = attachment_row.find(
                "a", attrs={"data-action": "downloadURL"}
            )
            if attachment_link:
                details["Attachments"].append(attachment_link.text.strip())
                details["_attachment_urls"].append(
                    "https://www.txsmartbuy.gov" + attachment_link["data-href"]
                )

    print(f"   Found {len(details['Attachments'])} attachments")

    # Check for Available Date
    available_date_div = soup.find(
        "div",
        class_="esbd-result-cell",
        string=lambda text: "Available Date" in text if text else False,
    )
    if available_date_div:
        available_date = (
            available_date_div.find("p").text.strip()
            if available_date_div.find("p")
            else ""
        )
        print(f"   Available Date found: {available_date}")
        return details, True, available_date

    print("📊 Bid details extraction complete")
    return details, False, None


def update_excel(data, excel_path):
    """
    Update the Excel file with new bid data.

    Args:
        data (dict): The bid details to be added to the Excel file.
        excel_path (str): The path to the Excel file.
    """
    print(f"📊 Updating Excel file: {excel_path}")
    # Convert dates to YYYY-MM-DD format
    for date_field in ["Posted Date", "Response Date"]:
        if data[date_field]:
            try:
                data[date_field] = datetime.strptime(
                    data[date_field], "%m/%d/%Y"
                ).strftime("%Y-%m-%d")
            except ValueError:
                data[date_field] = ""

    # Remove the _attachment_urls before saving to Excel
    data_for_excel = {k: v for k, v in data.items() if k != "_attachment_urls"}

    # Convert Attachments list to a string without an extra comma at the end
    if data_for_excel["Attachments"]:
        data_for_excel["Attachments"] = ", ".join(data_for_excel["Attachments"])
    else:
        data_for_excel["Attachments"] = ""

    # Define the correct column order
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

    if os.path.exists(excel_path):
        existing_df = pd.read_excel(excel_path)
        new_df = pd.DataFrame([data_for_excel])
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        updated_df = pd.DataFrame([data_for_excel])

    # Reset the SL No to ensure it starts from 1 and increments sequentially
    updated_df["SL No"] = range(1, len(updated_df) + 1)

    # Reorder columns
    updated_df = updated_df.reindex(columns=column_order)

    # Fill missing columns with empty strings
    for col in column_order:
        if col not in updated_df.columns:
            updated_df[col] = ""

    updated_df.to_excel(excel_path, index=False)
    print(f"✅ Excel file updated: {excel_path}")


def download_attachments(details, folder_name):
    """
    Download all attachments for a specific bid and move them to the related folder.

    Args:
        details (dict): The bid details containing attachment information.
        folder_name (str): The name of the folder to save the attachments.
    """
    print(f"📎 Downloading attachments for bid {details['Solicitation Number']}")
    if not details["Attachments"]:
        print(f"No attachments found for bid {details['Solicitation Number']}")
        return

    # Create bid folder directly under main_folder
    bid_folder = os.path.join(script_folder, folder_name)
    os.makedirs(bid_folder, exist_ok=True)

    total_attachments = len(details["Attachments"])
    print(
        f"Downloading {total_attachments} attachments for bid {details['Solicitation Number']}"
    )

    for index, (file_name, file_url) in enumerate(
        zip(details["Attachments"], details["_attachment_urls"]), start=1
    ):
        print(f"Downloading attachment {index}/{total_attachments}:")
        print(f"   File Name: {file_name}")
        print(f"   URL: {file_url}")
        print(f"   Solicitation Number: {details['Solicitation Number']}")
        print(f"   Posted Date: {details['Posted Date']}")
        try:
            response = requests.get(file_url)
            if response.status_code == 200:
                # Save the file temporarily in the progress folder
                temp_file_path = os.path.join(progress_folder, file_name)
                with open(temp_file_path, "wb") as file:
                    file.write(response.content)

                # Move the file to the bid-specific folder
                final_file_path = os.path.join(bid_folder, file_name)
                safe_move(temp_file_path, final_file_path)

                print(f"✅ Successfully downloaded and moved: {file_name}")
            else:
                print(f"❌ Failed to download {file_name}: HTTP {response.status_code}")
        except Exception as e:
            print(f"❌ Error downloading {file_name}: {str(e)}")

    print(f"✅ All attachments downloaded for bid {details['Solicitation Number']}")


def move_remaining_files():
    """
    Move any remaining files in the progress folder to their respective bid folders.
    """
    print("🚚 Moving remaining files...")
    try:
        if not os.path.exists(progress_folder):
            print(
                f"No remaining files to move - folder does not exist: {progress_folder}"
            )
            return

        files_moved = False
        for file_name in os.listdir(progress_folder):
            file_path = os.path.join(progress_folder, file_name)
            if os.path.isfile(file_path):
                try:
                    bid_number = file_name.split("_")[0]
                    bid_folder = os.path.join(script_folder, bid_number)
                    os.makedirs(bid_folder, exist_ok=True)
                    safe_move(file_path, os.path.join(bid_folder, file_name))
                    print(f"Moved remaining file {file_name} to {bid_folder}")
                    files_moved = True
                except Exception as e:
                    print(f"Warning: Could not move file {file_name}: {str(e)}")

        if not files_moved:
            print("No files needed to be moved")
        else:
            print("✅ All remaining files moved")
    except Exception as e:
        print(f"Warning: Error checking for remaining files: {str(e)}")


def handle_error(error_message):
    """
    Handle errors by playing a notification sound and pausing the script.

    Args:
        error_message (str): The error message to display.
    """
    print(f"❌ Error: {error_message}")
    play_notification_sound()
    input("Press Enter to continue...")


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

    # Add/update the bid in cache
    cache[bid_url] = {
        "posted_date": bid_details.get("Posted Date", ""),
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


def scrape_texas_smartbuy(batch_size=30):
    """
    Main function to execute the Texas SmartBuy bid extraction process.

    This function initializes the WebDriver, applies filters, extracts bid links,
    processes bid details, updates the Excel file, and downloads attachments.

    Args:
        batch_size (int): Number of bids to process in each batch.
    """
    print(f"🚀 Starting Texas SmartBuy scraping (batch size: {batch_size})")
    print(f"📁 Saving data in folder: {script_folder}")

    driver = setup_driver()
    print("Headless browser initialized")

    try:
        # Apply filters to the search page
        apply_filters(driver)

        excel_path = os.path.join(script_folder, f"{script_name}.xlsx")
        total_bids_processed = 0

        while True:
            print(f"\n📃 Processing new batch of bids...")
            bid_links = extract_bid_links(driver, batch_size=batch_size)

            if not bid_links:
                print("No more bids available. Ending extraction.")
                break

            for index, link in enumerate(bid_links, start=1):
                print(
                    f"\n🔍 Processing bid {total_bids_processed + index}/{len(bid_links)} in current batch"
                )
                try:
                    print(f"Extracting details for bid link: {link}")

                    # Add cache check here
                    bid_url = "https://www.txsmartbuy.gov" + link
                    if not should_process_bid_link(bid_url):
                        continue

                    details, has_available_date, available_date = extract_bid_details(
                        driver, link, total_bids_processed + index, len(bid_links)
                    )

                    if details is None:
                        print(
                            f"⏭️ Skipping bid {total_bids_processed + index} due to loading failure"
                        )
                        continue

                    if has_available_date:
                        print(
                            f"🛑 Found 'Available Date': {available_date}. Stopping extraction."
                        )
                        return

                    print(f"✅ Scraping bid: {details['Solicitation Number']}")

                    update_excel(details, excel_path)

                    folder_name = details["Solicitation Number"]
                    download_attachments(details, folder_name)

                    # Save to cache after successful processing
                    save_to_cache(bid_url, details)

                    total_bids_processed += 1
                    print(f"✅ Bid {total_bids_processed} processed successfully")
                except Exception as e:
                    print(
                        f"❌ Error processing bid {total_bids_processed + index}: {str(e)}"
                    )
                    handle_error(
                        f"Error processing bid {total_bids_processed + index}: {str(e)}"
                    )

            print(f"\n✅ Batch complete. Total bids processed: {total_bids_processed}")

            # Check if there are more pages
            next_page = driver.find_elements(By.ID, "Next")
            if not next_page:
                print("No more pages available. Ending extraction.")
                break

            # Navigate to the next page
            next_page_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "Next"))
            )
            driver.execute_script("arguments[0].click();", next_page_link)
            random_sleep(2, 4)

        if total_bids_processed > 0:
            print(
                f"💾 Final bid results saved to {script_name}.xlsx in {script_folder}"
            )
        else:
            print("❌ No bids were processed. No Excel file created.")

        print("🎉 All Bids and Attachments Extraction Successfully Completed")
    except Exception as e:
        handle_error(f"Error during scraping: {str(e)}")
        driver.save_screenshot("scraping_error.png")
        print("Screenshot saved as scraping_error.png")
    finally:
        driver.quit()
        print("Browser closed")
        move_remaining_files()
        print("Remaining files moved to respective bid folders")
        print("🏁 Scraping process finished")


def main():
    print("🏁 Starting main script execution")
    try:
        # Download all relevant bids in batches of 30
        scrape_texas_smartbuy(batch_size=30)

        # Move any remaining files
        move_remaining_files()

        # Clean up temporary download folder before renaming
        try:
            if os.path.exists(progress_folder):
                shutil.rmtree(progress_folder)
                print(f"✅ Removed temporary download folder: {progress_folder}")
        except Exception as e:
            print(f"⚠️ Error removing temporary folder: {str(e)}")

        # Rename script folder to indicate completion
        completed_folder = script_folder.replace("_IN_PROGRESS", "_COMPLETED")
        if os.path.exists(script_folder):
            os.rename(script_folder, completed_folder)
            print(f"✅ Folder renamed to indicate completion: {completed_folder}")

        # Success message
        print("✅ Script executed successfully. All relevant bids have been downloaded.")
        play_notification_sound()

    except Exception as e:
        handle_error(f"An error occurred: {str(e)}")

    print("🎉 Script execution complete")


if __name__ == "__main__":
    main()
