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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
import shutil
import threading
import queue
import argparse
from utils.utils import safe_move, play_notification_sound


def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Scrape bids from MyFloridaMarketPlace."
    )
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Number of days to look back for bids. Default is 2 days.",
    )
    return parser.parse_args()


def setup_driver(download_dir):
    """Set up and return a configured Chrome WebDriver."""
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )
    return webdriver.Chrome(options=chrome_options)


def wait_and_click(driver, by, value, wait_time=60, retries=3):
    """Wait for an element to be clickable and then click it, with retries."""
    for attempt in range(retries):
        try:
            # Wait for document ready state with longer timeout
            WebDriverWait(driver, wait_time).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            
            # Wait for any jQuery animations with longer timeout
            WebDriverWait(driver, wait_time).until(
                lambda d: d.execute_script("return jQuery.active == 0") if driver.execute_script("return typeof jQuery !== 'undefined'") else True
            )
            
            # Wait for network requests to complete
            WebDriverWait(driver, wait_time).until(
                lambda d: d.execute_script(
                    "return window.performance.getEntriesByType('resource').filter(r => !r.responseEnd).length == 0"
                )
            )
            
            # Wait for element and click with longer timeout
            element = WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((by, value))
            )
            
            # Scroll element into view
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            time.sleep(28)  # Longer pause after scroll
            
            # Try to click with both Selenium and JavaScript
            try:
                element.click()
            except:
                driver.execute_script("arguments[0].click();", element)
                time.sleep(7)  # Wait after JavaScript click
                
            return element
            
        except Exception as e:
            if attempt < retries - 1:
                print(f"Attempt {attempt + 1} failed. Retrying...")
                time.sleep(15)  # Longer wait between retries
                # Try refreshing the page
                driver.refresh()
                WebDriverWait(driver, wait_time).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
            else:
                print(f"Failed to click element after {retries} attempts: {by}, {value}")
                raise e


def is_within_date_range(date_str, days):
    """Check if the given date is within the specified range from today."""
    today = datetime.now().date()
    bid_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    return (today - bid_date).days <= days


def extract_bids(driver, days):
    """Extract bid information from the current page."""
    bids = []
    rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
    for row in rows:
        columns = row.find_elements(By.TAG_NAME, "td")
        if len(columns) >= 8:
            bid_number = columns[1].find_element(By.TAG_NAME, "a").text
            bid_url = f"https://vendor.myfloridamarketplace.com/search/bids/detail/{bid_number.split('-')[1]}"
            start_date = datetime.strptime(columns[6].text, "%m/%d/%Y").strftime(
                "%Y-%m-%d"
            )
            end_date = datetime.strptime(columns[7].text, "%m/%d/%Y").strftime(
                "%Y-%m-%d"
            )
            title = columns[0].text
            agency = columns[4].text
            if is_within_date_range(start_date, days):
                bids.append((bid_number, bid_url, start_date, end_date, title, agency))
    return bids


def extract_bid_details(driver, bid_url):
    """Extract detailed information for a single bid."""
    try:
        driver.get(bid_url)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        details = {
            "Notice Type": "",
            "Solicitation Number": "",
            "Solicitation Title": "",
            "Agency": "",
            "Category": "",
            "Description": "",
            "Contracting Office Address": "",
            "Bid Detail Page URL": bid_url,
            "Attachments": "",
        }

        # Extract Notice Type
        notice_type_selectors = [".mat-headline.f-sm", ".mat-headline", "h1", "h2"]
        for selector in notice_type_selectors:
            try:
                notice_type = (
                    WebDriverWait(driver, 10)
                    .until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    .text
                )
                if notice_type and notice_type != "Loading...":
                    details["Notice Type"] = notice_type
                    break
            except:
                continue

        # Extract Solicitation Number
        try:
            solicitation_number = (
                WebDriverWait(driver, 10)
                .until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//div[contains(text(), 'Advertisement Number:')]")
                    )
                )
                .text
            )
            details["Solicitation Number"] = solicitation_number.split(": ")[1].strip()
        except:
            pass

        # Extract other details with error handling
        try:
            details["Solicitation Title"] = driver.find_element(
                By.CLASS_NAME, "mat-headline"
            ).text
        except NoSuchElementException:
            pass

        try:
            details["Agency"] = driver.find_element(By.CLASS_NAME, "f-sm").text
        except NoSuchElementException:
            pass

        try:
            details["Description"] = driver.find_element(By.ID, "mainSection").text
        except NoSuchElementException:
            pass

        # Extract Category (Commodity Codes)
        try:
            commodity_codes_table = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "mfmp-commodity-codes-list table tbody")
                )
            )
            rows = commodity_codes_table.find_elements(By.TAG_NAME, "tr")
            categories = []
            for row in rows:
                code = row.find_element(By.CSS_SELECTOR, "td:nth-child(1)").text.strip()
                description = row.find_element(
                    By.CSS_SELECTOR, "td:nth-child(2)"
                ).text.strip()
                categories.append(f"{code} - {description}")
            details["Category"] = "; ".join(categories)
        except Exception as e:
            print(f"Error extracting category: {str(e)}")

        # Extract Contracting Office Address
        try:
            contact_info = []
            xpaths = [
                "/html/body/mfmp-root/mfmp-bids-view/main/div/mfmp-bid-detail/section[4]/div[1]/span[2]",
                "/html/body/mfmp-root/mfmp-bids-view/main/div/mfmp-bid-detail/section[4]/div[2]/span[2]",
                "/html/body/mfmp-root/mfmp-bids-view/main/div/mfmp-bid-detail/section[4]/div[3]/span[2]",
                "/html/body/mfmp-root/mfmp-bids-view/main/div/mfmp-bid-detail/section[4]/div[3]/span[2]/span[1]",
                "/html/body/mfmp-root/mfmp-bids-view/main/div/mfmp-bid-detail/section[4]/div[3]/span[2]/span[2]",
                "/html/body/mfmp-root/mfmp-bids-view/main/div/mfmp-bid-detail/section[4]/div[3]",
                "/html/body/mfmp-root/mfmp-bids-view/main/div/mfmp-bid-detail/section[4]/div[4]/span[2]/a",
            ]

            for xpath in xpaths:
                try:
                    element = driver.find_element(By.XPATH, xpath)
                    if element.tag_name == "a":
                        contact_info.append(element.get_attribute("href"))
                    else:
                        contact_info.append(element.text.strip())
                except NoSuchElementException:
                    pass

            details["Contracting Office Address"] = " ".join(filter(None, contact_info))
        except Exception as e:
            print(f"Error extracting Contracting Office Address: {str(e)}")

        # Extract Attachments
        try:
            attachment_links = driver.find_elements(By.CSS_SELECTOR, "a.document-link")
            attachments = [link.text.strip() for link in attachment_links]
            details["Attachments"] = ", ".join(attachments)
        except:
            pass

        return details
    except Exception as e:
        print(f"Error extracting details for bid URL {bid_url}: {str(e)}")
        return None


def is_download_completed(file_path, timeout=300):
    """Check if the file has finished downloading."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        if (
            os.path.exists(file_path)
            and not file_path.endswith(".tmp")
            and not file_path.endswith(".crdownload")
        ):
            return True
        time.sleep(1)
    return False


def wait_for_downloads_to_complete(downloads_folder, timeout=600):
    """Wait for all downloads in the folder to complete."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        incomplete_downloads = [
            f
            for f in os.listdir(downloads_folder)
            if f.endswith(".crdownload") or f.endswith(".tmp")
        ]
        if not incomplete_downloads:
            return True
        time.sleep(1)
    print(f"Timeout reached: Downloads did not complete within {timeout} seconds.")
    return False


def download_file(driver, link, file_name, downloads_folder, bid_folder):
    """Download a single file and move it to the bid folder."""
    try:
        initial_files = set(os.listdir(downloads_folder))

        # Get the original file name from the link text
        original_filename = link.text.strip()

        # Scroll element into view and click using ActionChains
        driver.execute_script("arguments[0].scrollIntoView(true);", link)
        time.sleep(1)

        actions = ActionChains(driver)
        actions.move_to_element(link)
        actions.click()
        actions.perform()
        time.sleep(2)  # Wait for download to initiate

        # Wait for download to start
        try:
            WebDriverWait(driver, 30).until(
                lambda x: len(set(os.listdir(downloads_folder)) - initial_files) > 0
            )
        except TimeoutException:
            print(f"Timeout waiting for download to start: {original_filename}")
            return

        # Wait for download to complete
        if wait_for_downloads_to_complete(downloads_folder):
            new_files = set(os.listdir(downloads_folder)) - initial_files
            if new_files:
                downloaded_file = new_files.pop()
                source_path = os.path.join(downloads_folder, downloaded_file)

                # Get file extension from downloaded file
                _, ext = os.path.splitext(downloaded_file)
                if not ext:
                    ext = ".pdf"  # Default to .pdf if no extension

                # Create valid filename with extension
                base_name = "".join(
                    c for c in original_filename if c.isalnum() or c in (" ", "-", "_")
                )
                valid_filename = f"{base_name}{ext}"
                target_path = os.path.join(bid_folder, valid_filename)

                # Ensure unique filename
                counter = 1
                while os.path.exists(target_path):
                    target_path = os.path.join(
                        bid_folder, f"{base_name}_{counter}{ext}"
                    )
                    counter += 1

                try:
                    safe_move(source_path, target_path)
                    print(f"Successfully moved file: {valid_filename}")
                except Exception as e:
                    print(f"Error moving file {valid_filename}: {str(e)}")
            else:
                print(f"No new file found for: {original_filename}")
        else:
            print(f"Download did not complete in time: {original_filename}")

    except Exception as e:
        print(f"Error downloading file {original_filename}: {str(e)}")


def download_attachments(driver, bid_number, main_folder, downloads_folder):
    """Download attachments for a bid using threads."""
    try:
        # Wait for attachment links to be present
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.document-link"))
        )

        attachment_links = driver.find_elements(
            By.CSS_SELECTOR, "tbody tr td:first-child a.document-link"
        )

        if not attachment_links:
            print(f"No attachments found for bid {bid_number}")
            return

        # Create bid folder in the script's folder (not in main_folder)
        script_name = os.path.splitext(os.path.basename(__file__))[0]
        script_folder = os.path.join(main_folder, f"{script_name}_IN_PROGRESS")
        bid_folder = os.path.join(script_folder, bid_number)

        if not os.path.exists(bid_folder):
            os.makedirs(bid_folder)

        total_attachments = len(attachment_links)
        print(f"Found {total_attachments} attachments for bid {bid_number}")

        download_threads = []
        for index, link in enumerate(attachment_links, start=1):
            try:
                file_name = link.text.strip()
                if not file_name:
                    file_name = f"attachment_{index}.pdf"
                print(f"Downloading {file_name} ({index}/{total_attachments})")

                thread = threading.Thread(
                    target=download_file,
                    args=(driver, link, file_name, downloads_folder, bid_folder),
                )
                thread.start()
                download_threads.append(thread)
                time.sleep(1)  # Add small delay between starting downloads
            except Exception as e:
                print(f"Error starting download for attachment {index}: {str(e)}")

        # Wait for all downloads to complete
        for thread in download_threads:
            thread.join(timeout=300)  # 5 minute timeout per thread

        print(f"Completed downloading attachments for bid {bid_number}")

    except Exception as e:
        pass
        # print(f"Error in download_attachments for bid {bid_number}: {str(e)}")


def main():
    """Main function to execute the bid extraction process for multiple sites."""
    args = parse_arguments()
    days = args.days
    print(f"Bids Extraction Started for the last {days} days")

    # Create yesterday's date folder
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    main_folder = os.path.join(project_root, yesterday)
    os.makedirs(main_folder, exist_ok=True)

    # Create script-specific folder with IN_PROGRESS suffix
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    script_folder_name = f"{script_name}_IN_PROGRESS"
    script_folder = os.path.join(main_folder, script_folder_name)

    # Remove existing script folder if it exists
    if os.path.exists(script_folder):
        shutil.rmtree(script_folder)
    os.makedirs(script_folder)

    # Create temporary downloads folder inside script folder
    downloads_folder = os.path.join(script_folder, "downloads")
    os.makedirs(downloads_folder)

    # Initialize driver with downloads folder
    driver = setup_driver(downloads_folder)

    try:
        driver.get("https://vendor.myfloridamarketplace.com/search/bids")

        # Wait for the page to load completely
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        # Set maximum results to 50
        wait_and_click(driver, By.ID, "mat-select-0")
        wait_and_click(driver, By.XPATH, "//span[contains(text(), ' 50 ')]")

        # Set Ad Type filters
        wait_and_click(
            driver, By.XPATH, "//mat-panel-title[contains(text(), 'Ad Type')]"
        )
        ad_types = [
            "Grant Opportunities",
            "Invitation to Bid",
            "Request for Proposals",
            "Request for Information",
            "Request for Statement of Qualifications",
        ]
        for ad_type in ad_types:
            wait_and_click(
                driver,
                By.XPATH,
                f"//div[contains(text(), '{ad_type}')]/preceding-sibling::mat-pseudo-checkbox",
            )

        # Set Ad Status to OPEN
        wait_and_click(
            driver, By.XPATH, "//mat-panel-title[contains(text(), 'Ad Status')]"
        )
        wait_and_click(
            driver,
            By.XPATH,
            "//div[contains(text(), ' OPEN ')]/preceding-sibling::mat-pseudo-checkbox",
        )

        # Click Search
        wait_and_click(driver, By.XPATH, "//span[contains(text(), ' Search ')]")

        # Wait for results to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr"))
        )

        all_bids = []
        page_number = 1
        while True:
            print(f"Current Page Number: {page_number}")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr"))
            )
            bids = extract_bids(driver, days)
            all_bids.extend(bids)

            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "button.mat-paginator-navigation-next")
                    )
                )
                if "disabled" in next_button.get_attribute("class"):
                    print("Reached the last page.")
                    break

                # Scroll the button into view
                driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                time.sleep(1)  # Short pause to allow any animations to complete

                # Use ActionChains to move to and click the button
                actions = ActionChains(driver)
                actions.move_to_element(next_button).click().perform()

                # Wait for the page to update
                WebDriverWait(driver, 10).until(
                    EC.staleness_of(driver.find_element(By.CSS_SELECTOR, "tbody tr"))
                )
                page_number += 1
                time.sleep(2)  # Additional wait to ensure page has loaded
            except Exception as e:
                print(f"Error navigating to next page: {str(e)}")
                break

        print("Bids Extraction Completed")
        print(f"Total bids found: {len(all_bids)}")

        # Create Excel file in the script folder (not directly in yesterday's folder)
        excel_file = os.path.join(script_folder, f"{script_name}.xlsx")

        # Create DataFrame to store bid details
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
                "Additional Summary",
                "Contracting Office Address",
                "Contact Information",
                "Bid Detail Page URL",
                "Attachments",  # New column
            ]
        )

        for i, (bid_number, bid_url, start_date, end_date, title, agency) in enumerate(
            all_bids, start=1
        ):
            print(f"Processing bid {i}/{len(all_bids)}: {bid_number}")
            details = extract_bid_details(driver, bid_url)
            if details:
                details["SL No"] = i
                details["Posted Date"] = start_date
                details["Response Date"] = end_date
                details["Additional Summary"] = ""
                details["Contact Information"] = ""
                details["Solicitation Title"] = title
                details["Agency"] = agency

                # Update DataFrame with new bid details
                df = pd.concat([df, pd.DataFrame([details])], ignore_index=True)

                # Format date columns
                date_columns = ["Posted Date", "Response Date"]
                for col in date_columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime("%Y-%m-%d")

                # Create bid-specific folder inside script folder
                bid_folder = os.path.join(script_folder, bid_number)
                os.makedirs(bid_folder, exist_ok=True)

                # Download attachments and move them to the bid-specific folder
                download_attachments(driver, bid_number, main_folder, downloads_folder)

                # Wait for all downloads to complete
                if wait_for_downloads_to_complete(downloads_folder):
                    # Update Attachments column based on actual downloaded files
                    attachments = []
                    script_folder = os.path.join(
                        main_folder, f"{script_name}_IN_PROGRESS"
                    )
                    bid_folder = os.path.join(script_folder, bid_number)

                    if os.path.exists(bid_folder):
                        attachments = [
                            f
                            for f in os.listdir(bid_folder)
                            if os.path.isfile(os.path.join(bid_folder, f))
                        ]

                    # Update the Attachments column for this bid
                    df.loc[
                        df["Solicitation Number"] == details["Solicitation Number"],
                        "Attachments",
                    ] = ", ".join(attachments)

                    # Save updated DataFrame to Excel file directly in the script folder
                    df.to_excel(excel_file, index=False)
                    print(f"Excel file updated: {excel_file}")

                    print(f"Attachments downloaded and moved to: {bid_folder}")
                else:
                    print(f"Warning: Not all downloads completed for bid {bid_number}")

        if not df.empty:
            print(f"Final Excel file saved: {excel_file}")
        else:
            print("No bid details were successfully extracted.")

        # Remove the file moving section and directly rename the folder
        completed_folder_name = f"{script_name}_COMPLETED"
        completed_folder = os.path.join(main_folder, completed_folder_name)
        try:
            # Clean up temporary download folder before renaming
            try:
                if os.path.exists(downloads_folder):
                    shutil.rmtree(downloads_folder)
                    print(f"✅ Removed temporary download folder: {downloads_folder}")
            except Exception as e:
                print(f"⚠️ Error removing temporary folder: {str(e)}")

            # Remove completed folder if it already exists
            if os.path.exists(completed_folder):
                shutil.rmtree(completed_folder)
            # Rename the in-progress folder to completed
            os.rename(script_folder, completed_folder)
            print(f"Marked folder as completed: {completed_folder}")
        except Exception as e:
            print(f"Error marking folder as completed: {str(e)}")

        print("All Bids and Attachments Extraction Successfully Completed")
        play_notification_sound()

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        import traceback
        print(traceback.format_exc())
        play_notification_sound()
        input("Press Enter to continue...")
    finally:
        driver.quit()
        # Clean up downloads folder
        shutil.rmtree(downloads_folder, ignore_errors=True)
        print(f"Cleaned up downloads folder: {downloads_folder}")


if __name__ == "__main__":
    main()
