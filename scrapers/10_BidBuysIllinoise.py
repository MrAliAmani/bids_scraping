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
    NoSuchElementException,
    TimeoutException,
    StaleElementReferenceException,
    WebDriverException,
)
import shutil
import re
from selenium.webdriver.common.action_chains import ActionChains
import threading
from selenium.webdriver.support.ui import Select
import signal
import sys
import winsound  # For Windows notification sounds
import argparse
from utils.utils import safe_move, play_notification_sound
import json

# Script name following the new convention
script_name = "10_BidBuy_Illinois"

# Create a folder named after yesterday's date
yesterday = datetime.now() - timedelta(days=1)
date_folder = yesterday.strftime("%Y-%m-%d")

# Create the script folder path with IN_PROGRESS suffix
script_folder = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    date_folder,
    f"{script_name}_IN_PROGRESS",
)

# Create the temporary download folder inside the script folder
main_folder = os.path.join(script_folder, script_name)

# Create all necessary directories
os.makedirs(script_folder, exist_ok=True)
os.makedirs(main_folder, exist_ok=True)

# Global variables for bid processing
total_bids = 0
current_bid = 0


def parse_arguments():
    parser = argparse.ArgumentParser(description="BidBuy Illinois Scraper")
    parser.add_argument(
        "--days", type=int, default=2, help="Number of days to scrape (default: 2)"
    )
    return parser.parse_args()


def setup_driver():
    """Set up and configure the Chrome WebDriver."""
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": main_folder,
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
    return webdriver.Chrome(options=chrome_options)


def click_advanced_search(driver):
    """
    Perform advanced search on the BidBuy website.

    This function clicks the advanced search button, sets the status to "Sent",
    and initiates the search.
    """
    try:
        # Wait for the page to load completely
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//body"))
        )

        # Scroll to the top of the page
        driver.execute_script("window.scrollTo(0, 0);")

        # Click Advanced Search button
        advanced_search = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[9]/div/div[1]/div[1]/div/fieldset/legend")
            )
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", advanced_search)
        driver.execute_script("arguments[0].click();", advanced_search)

        # Wait for the status dropdown to be visible and interactable
        status_dropdown = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "bidSearchForm:status"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", status_dropdown)
        select = Select(status_dropdown)
        select.select_by_value("2BS")

        # Click Search button
        search_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/div[9]/div/div[1]/div[1]/div/fieldset/div/div[2]/div/form/div[9]/div/button[1]",
                )
            )
        )
        driver.execute_script("arguments[0].click();", search_button)

        # Wait for search results to load
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//table[@role='grid']"))
        )

        print("Advanced search completed successfully")
    except Exception as e:
        print(f"Error in advanced search: {str(e)}")
        raise


def extract_bid_links(driver, max_links=25):
    """
    Extract bid links from the search results page.

    Args:
        driver (WebDriver): The Selenium WebDriver instance.
        max_links (int): Maximum number of links to extract per page.

    Returns:
        list: A list of extracted bid links.
    """
    bid_links = []

    # Wait for the table to be present
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//table[@role='grid']"))
    )

    rows = driver.find_elements(
        By.XPATH,
        "//tr[@class='ui-widget-content ui-datatable-even' or @class='ui-widget-content ui-datatable-odd']",
    )

    for row in rows:
        try:
            bid_link = row.find_element(By.XPATH, ".//td[1]/a").get_attribute("href")
            bid_links.append(bid_link)
            if len(bid_links) >= max_links:
                break
        except StaleElementReferenceException:
            print("Encountered a stale element. Retrying...")
            break
        except Exception as e:
            print(f"Error processing row: {str(e)}")
            continue

    return bid_links


from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


def extract_bid_details(driver, bid_link):
    """
    Extract details from a specific bid page.

    Args:
        driver (WebDriver): The Selenium WebDriver instance.
        bid_link (str): The URL of the bid detail page.

    Returns:
        dict: A dictionary containing the extracted bid details.
    """
    driver.get(bid_link)
    bid_details = {}

    try:
        # Wait for the page to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//table[@role='presentation']"))
        )

        # Helper function to safely extract text from an element
        def safe_extract(xpath, key):
            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                bid_details[key] = element.text.strip()
            except (TimeoutException, NoSuchElementException):
                print(f"Could not find element for {key}")
                bid_details[key] = ""

        # Extract Available Date (Posted Date)
        safe_extract(
            "//td[contains(text(), 'Available Date')]/following-sibling::td",
            "Posted Date",
        )
        if bid_details["Posted Date"]:
            date_match = re.search(r"\d{2}/\d{2}/\d{4}", bid_details["Posted Date"])
            if date_match:
                bid_details["Posted Date"] = datetime.strptime(
                    date_match.group(), "%m/%d/%Y"
                ).strftime("%Y-%m-%d")
            else:
                bid_details["Posted Date"] = ""

        # Extract Bid Opening Date (Response Date)
        safe_extract(
            "//td[contains(text(), 'Bid Opening Date')]/following-sibling::td",
            "Response Date",
        )
        if bid_details["Response Date"]:
            date_match = re.search(r"\d{2}/\d{2}/\d{4}", bid_details["Response Date"])
            if date_match:
                bid_details["Response Date"] = datetime.strptime(
                    date_match.group(), "%m/%d/%Y"
                ).strftime("%Y-%m-%d")
            else:
                bid_details["Response Date"] = ""

        # Extract Bid Number
        safe_extract(
            "//td[contains(text(), 'Bid Number')]/following-sibling::td",
            "Solicitation Number",
        )

        # Extract Description
        safe_extract(
            "//td[contains(text(), 'Description')]/following-sibling::td",
            "Solicitation Title",
        )

        # Extract Organization
        safe_extract(
            "//td[contains(text(), 'Organization')]/following-sibling::td", "Agency"
        )

        # Extract Bulletin Description
        safe_extract(
            "//td[contains(text(), 'Bulletin Desc')]/following-sibling::td",
            "Description",
        )

        # Extract Ship-to Address
        safe_extract(
            "//td[contains(text(), 'Ship-to Address')]/following-sibling::td",
            "Contracting Office Address",
        )

        # Extract NIGP codes and descriptions
        nigp_codes = []
        table_number = 2
        while True:
            try:
                code_xpath = f"/html/body/form/table/tbody/tr/td/table/tbody/tr[4]/td/table[{table_number}]/tbody/tr/td/table/tbody/tr[2]/td/table[1]/tbody/tr/td[2]/u"
                desc_xpath = f"/html/body/form/table/tbody/tr/td/table/tbody/tr[4]/td/table[{table_number}]/tbody/tr/td/table/tbody/tr[2]/td/table[1]/tbody/tr/td[2]"

                code_element = driver.find_element(By.XPATH, code_xpath)
                desc_element = driver.find_element(By.XPATH, desc_xpath)

                code = code_element.text.strip()
                description = desc_element.text.replace(code, "").strip()

                nigp_codes.append(f"{code} - {description}")
                table_number += 1
            except NoSuchElementException:
                # No more NIGP codes found, exit the loop
                break
            except Exception as e:
                print(f"Error extracting NIGP code: {str(e)}")
                break

        bid_details["Category"] = ", ".join(nigp_codes)

        bid_details["Notice Type"] = ""
        bid_details["Contact Information"] = ""
        bid_details["Bid Detail Page URL"] = bid_link
        bid_details["Additional Summary, if any"] = ""

        # Extract file attachments
        file_attachments = driver.find_elements(
            By.XPATH,
            "//td[contains(text(), 'File Attachments')]/following-sibling::td//a",
        )
        bid_details["Attachments"] = ", ".join([att.text for att in file_attachments])

    except Exception as e:
        print(f"Error extracting bid details: {str(e)}")
        return None

    return bid_details


def move_file(source, destination):
    """Move a file from source to destination."""
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            safe_move(source, destination)
            return True
        except PermissionError:
            if attempt < max_attempts - 1:
                time.sleep(2)  # Wait for 2 seconds before retrying
            else:
                print(f"Failed to move file after {max_attempts} attempts: {source}")
                return False


def download_attachments(driver, bid_number):
    """
    Download attachments for a specific bid.

    Args:
        driver (WebDriver): The Selenium WebDriver instance.
        bid_number (str): The bid number used to create a subfolder.

    Returns:
        bool: True if any attachments were downloaded successfully, False otherwise.
    """
    attachments_downloaded = False
    bid_folder = os.path.join(
        script_folder,  # Changed from os.path.dirname(main_folder)
        bid_number.replace("/", "_"),
    )
    os.makedirs(bid_folder, exist_ok=True)

    attachments = driver.find_elements(
        By.XPATH, "//a[contains(@href, 'javascript:downloadFile')]"
    )

    if not attachments:
        print(f"No attachments found for bid {bid_number}")
        return attachments_downloaded

    print(f"Found {len(attachments)} attachments for bid {bid_number}")

    for index, attachment in enumerate(attachments, start=1):
        file_name = attachment.text
        print(f"Downloading attachment {index}/{len(attachments)}: {file_name}")
        try:
            before = os.listdir(main_folder)
            attachment.click()

            # Wait for up to 5 minutes (300 seconds) in 30-second intervals
            max_wait_time = 300
            interval = 30
            downloaded_file = None

            for _ in range(0, max_wait_time, interval):
                time.sleep(interval)
                after = os.listdir(main_folder)
                new_files = [f for f in after if f not in before]

                if new_files:
                    downloaded_file = new_files[0]
                    source_path = os.path.join(main_folder, downloaded_file)

                    # Check if the file has finished downloading
                    if not downloaded_file.endswith(".crdownload") and os.path.exists(
                        source_path
                    ):
                        break

            if downloaded_file:
                source_path = os.path.join(main_folder, downloaded_file)
                destination_path = os.path.join(bid_folder, file_name)

                # Retry moving the file
                move_success = False
                for attempt in range(5):
                    try:
                        safe_move(source_path, destination_path)
                        print(f"Successfully moved {file_name} to {bid_folder}")
                        attachments_downloaded = True
                        move_success = True
                        break
                    except FileNotFoundError:
                        if attempt < 4:
                            time.sleep(2)
                        else:
                            print(f"Failed to move {file_name} after 5 attempts")
                    except Exception as e:
                        if attempt < 4:
                            time.sleep(2)
                        else:
                            print(f"Error moving {file_name}: {str(e)}")

                if not move_success:
                    print(f"Failed to move {file_name}. It will be moved later.")
            else:
                print(f"⚠️ Download did not complete for: {file_name} within 5 minutes")
        except Exception as e:
            print(f"⚠️ Error downloading attachment for bid {bid_number}: {str(e)}")

    if attachments_downloaded:
        print(f"✅ Successfully downloaded attachments for bid {bid_number}")
    else:
        print(f"❌ Failed to download any attachments for bid {bid_number}")

    return attachments_downloaded


def update_excel(bid_data, sl_no):
    """Update the Excel file with new bid data."""
    try:
        bid_data["SL No"] = sl_no
        excel_file = os.path.join(
            script_folder,  # Changed from os.path.dirname(main_folder)
            f"{script_name}.xlsx",
        )

        for date_field in ["Posted Date", "Response Date"]:
            if bid_data[date_field]:
                try:
                    bid_data[date_field] = datetime.strptime(
                        bid_data[date_field], "%Y-%m-%d"
                    ).strftime("%Y-%m-%d")
                except ValueError:
                    bid_data[date_field] = ""

        if os.path.exists(excel_file):
            existing_df = pd.read_excel(excel_file)
            updated_df = pd.concat(
                [existing_df, pd.DataFrame([bid_data])], ignore_index=True
            )
        else:
            updated_df = pd.DataFrame([bid_data])

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

        # Remove extra comma after the last attachment name
        updated_df["Attachments"] = (
            updated_df["Attachments"].str.strip().str.rstrip(",")
        )

        updated_df.to_excel(excel_file, index=False)
        print(f"Updated {script_name}.xlsx with bid {bid_data['Solicitation Number']}")

        # Debug information
        print(f"Excel file path: {excel_file}")
        print(f"Excel file exists: {os.path.exists(excel_file)}")
        print(f"Excel file size: {os.path.getsize(excel_file)} bytes")
        print(f"Number of rows in Excel: {len(updated_df)}")

    except Exception as e:
        print(f"Error updating Excel file: {str(e)}")
        play_notification_sound()
        input("Press Enter to continue...")


def signal_handler(sig, frame):
    print("You pressed Ctrl+C!")
    print("Moving remaining files before exiting...")
    move_remaining_files()
    sys.exit(0)


def move_remaining_files():
    """Move any remaining files in the download folder to appropriate bid folders."""
    for file in os.listdir(main_folder):
        file_path = os.path.join(main_folder, file)
        if os.path.isfile(file_path):
            if file.endswith(".crdownload"):
                print(f"Removing incomplete download: {file}")
                os.remove(file_path)
            else:
                moved = False
                date_folder_path = os.path.dirname(main_folder)
                for bid_folder in os.listdir(date_folder_path):
                    bid_folder_path = os.path.join(date_folder_path, bid_folder)
                    if os.path.isdir(bid_folder_path) and not bid_folder.endswith(
                        "_IN_PROGRESS"
                    ):
                        try:
                            safe_move(file_path, bid_folder_path)
                            print(f"Moved {file} to {bid_folder}")
                            moved = True
                            break
                        except:
                            pass
                if not moved:
                    print(f"Failed to move {file}")


def mark_completion():
    """Rename the folder to mark completion."""
    try:
        # Clean up temporary download folder first
        if os.path.exists(main_folder):
            try:
                shutil.rmtree(main_folder, ignore_errors=True)
                print(f"✅ Removed temporary download folder: {main_folder}")
            except Exception as e:
                print(f"Error removing temporary download folder: {e}")

        # Then rename the script folder to COMPLETED
        completed_folder = script_folder.replace("_IN_PROGRESS", "_COMPLETED")
        if os.path.exists(script_folder):
            os.rename(script_folder, completed_folder)
            print(f"✅ Marked folder as completed: {completed_folder}")
    except Exception as e:
        print(f"❌ Error marking completion: {str(e)}")


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
        print(f"⏭️ Skipping previously processed bid from {bid_data['posted_date']}")
        return False
    return True


def main(days_to_scrape):
    """Main function to execute the BidBuy bid extraction process."""
    global total_bids, current_bid
    print(
        f"🚀 Bids Extraction Started (Scraping bids from the last {days_to_scrape} days)"
    )

    signal.signal(signal.SIGINT, signal_handler)

    max_retries = 3
    for attempt in range(max_retries):
        driver = None
        try:
            driver = setup_driver()
            driver.get(
                "https://www.bidbuy.illinois.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true"
            )

            click_advanced_search(driver)

            # Open a new tab for bid details
            driver.execute_script("window.open('');")
            search_window = driver.window_handles[0]
            detail_window = driver.window_handles[-1]

            processed_bids = 0
            sl_no = 1
            cutoff_date = datetime.now() - timedelta(days=days_to_scrape)
            page_number = 1
            total_bids_processed = 0

            while True:
                print(f"Current Page Number: {page_number}")
                driver.switch_to.window(search_window)
                bid_links = extract_bid_links(driver, max_links=25)
                print(f"📊 Total bid links extracted: {len(bid_links)}")

                if not bid_links:
                    print("No more bids found. Exiting.")
                    break

                total_bids += len(bid_links)
                new_bids_processed = 0

                for link in bid_links:
                    current_bid += 1
                    print(f"Processing bid {current_bid} out of {total_bids}")
                    max_bid_retries = 3
                    for bid_attempt in range(max_bid_retries):
                        try:
                            # Switch to the detail window
                            driver.switch_to.window(detail_window)

                            # Add cache check here
                            if not should_process_bid_link(link):
                                break

                            driver.get(link)

                            # Wait for the page to load completely
                            WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//body"))
                            )

                            bid_details = extract_bid_details(driver, link)
                            if bid_details is None:
                                print(f"Failed to extract details for bid: {link}")
                                break

                            if (
                                "Posted Date" in bid_details
                                and bid_details["Posted Date"]
                            ):
                                try:
                                    posted_date = datetime.strptime(
                                        bid_details["Posted Date"], "%Y-%m-%d"
                                    )
                                    if posted_date > cutoff_date:
                                        attachments_downloaded = download_attachments(
                                            driver, bid_details["Solicitation Number"]
                                        )
                                        update_excel(bid_details, sl_no)

                                        # Save to cache after successful processing
                                        save_to_cache(link, bid_details)

                                        processed_bids += 1
                                        new_bids_processed += 1
                                        sl_no += 1
                                        print(
                                            f"Processed bid: {bid_details['Solicitation Number']}"
                                        )
                                    else:
                                        print(
                                            f"Skipping bid {bid_details['Solicitation Number']} - older than {days_to_scrape} days"
                                        )
                                except ValueError:
                                    print(
                                        f"⚠️ Could not parse date for bid {bid_details['Solicitation Number']}"
                                    )
                            else:
                                print(
                                    f"⚠️ No Posted Date found for bid {bid_details.get('Solicitation Number', 'Unknown')}"
                                )

                            break  # Break the retry loop if successful
                        except WebDriverException as e:
                            print(
                                f"WebDriverException occurred (Attempt {bid_attempt + 1}/{max_bid_retries}): {str(e)}"
                            )
                            if bid_attempt < max_bid_retries - 1:
                                print("Attempting to recover...")
                                try:
                                    driver.quit()
                                    driver = setup_driver()
                                    driver.get(
                                        "https://www.bidbuy.illinois.gov/bso/view/search/external/advancedSearchBid.xhtml?openBids=true"
                                    )
                                    click_advanced_search(driver)
                                    driver.execute_script("window.open('');")
                                    search_window = driver.window_handles[0]
                                    detail_window = driver.window_handles[-1]
                                    time.sleep(5)  # Wait for the browser to stabilize
                                except Exception as recovery_error:
                                    print(f"Failed to recover: {str(recovery_error)}")
                            else:
                                print(
                                    "Max retries reached for this bid. Moving to the next."
                                )
                        except Exception as e:
                            print(f"Error processing bid: {str(e)}")
                            play_notification_sound()
                            input("Press Enter to continue...")
                            break

                    # Add a small delay between processing bids to avoid overwhelming the server
                    time.sleep(2)

                total_bids_processed += new_bids_processed
                if new_bids_processed == 0:
                    print("No new bids found on this page. Exiting.")
                    break

                # Move to the next page
                driver.switch_to.window(search_window)
                try:
                    next_page = WebDriverWait(driver, 30).until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                "/html/body/div[9]/div/div[1]/div[1]/div/div/div/div[2]/div/form/span/div/div[1]/a[3]",
                            )
                        )
                    )
                    next_page.click()
                    page_number += 1
                    WebDriverWait(driver, 30).until(EC.staleness_of(next_page))
                except (NoSuchElementException, TimeoutException):
                    print("No more pages available. Exiting.")
                    break

            print(f"✅ Total bids processed: {total_bids_processed}")

            if total_bids_processed > 0:
                excel_file = os.path.join(main_folder, f"{script_name}.xlsx")
                print(f"💾 Final bid results saved to {excel_file}")
            else:
                print("❌ No bids were processed. No Excel file created.")

            # Move remaining files
            move_remaining_files()

            print("🎉 All Bids and Attachments Extraction Successfully Completed")
            play_notification_sound()  # Success notification
            break  # Exit the retry loop if successful

            # After successful completion, mark the folder as completed
            mark_completion()

        except Exception as e:
            print(
                f"❗ An error occurred (Attempt {attempt + 1}/{max_retries}): {str(e)}"
            )
            play_notification_sound()  # Error notification
            if attempt < max_retries - 1:
                input("Press Enter to retry...")
            else:
                print("Max retries reached. Exiting.")
                move_remaining_files()  # Ensure remaining files are moved even if max retries are reached

        finally:
            if driver:
                driver.quit()


if __name__ == "__main__":
    args = parse_arguments()
    main(args.days)
