import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import functools

# Override the print function to always flush
print = functools.partial(print, flush=True)

import os
import shutil

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import time
import shutil
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    NoSuchWindowException,
    StaleElementReferenceException,
    WebDriverException,
)
import requests
import threading
import re
import urllib3
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
from selenium.webdriver.common.action_chains import ActionChains
import urllib.parse
from utils.utils import safe_move, play_notification_sound

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Update the folder structure
yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
project_folder = os.path.dirname(
    os.path.dirname(__file__)
)  # Get the main project folder
main_folder = os.path.join(project_folder, yesterday)
os.makedirs(main_folder, exist_ok=True)

# Create a script-specific folder for temporary downloads and mark as in progress
script_name = os.path.splitext(os.path.basename(__file__))[0]
script_folder_name = f"{script_name}_IN_PROGRESS"
script_folder = os.path.join(main_folder, script_folder_name)
os.makedirs(script_folder, exist_ok=True)

# Create a temporary download folder within the script folder
script_download_folder = os.path.join(script_folder, script_name)
os.makedirs(script_download_folder, exist_ok=True)


def setup_driver():
    """
    Set up and configure the Chrome WebDriver.

    Returns:
        webdriver.Chrome: Configured Chrome WebDriver instance.
    """
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--ignore-ssl-errors=yes")
    chrome_options.add_argument("--ignore-certificate-errors")
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")  # Disables GPU acceleration
    chrome_options.add_argument("--disable-software-rasterizer")  # Use CPU rendering
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    )

    prefs = {
        "download.default_directory": script_download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    return webdriver.Chrome(options=chrome_options)


def set_search_criteria(driver):
    """
    Set the search criteria on the Georgia Procurement Registry website.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
    """
    print("Bids Extraction Started")

    # Select "Event Status = OPEN"
    status_dropdown = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "eventStatus"))
    )
    status_dropdown.find_element(By.XPATH, "//option[@value='OPEN']").click()

    # Select "Event Date Range = Event Start Date"
    date_range_dropdown = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "dateRangeType"))
    )
    date_range_dropdown.find_element(
        By.XPATH, "//option[@value='EVENT_START_DATE_TYPE']"
    ).click()

    # Set date range for the last 2 days
    end_date = datetime.now()
    start_date = end_date - timedelta(days=2)

    start_date_input = driver.find_element(By.ID, "rangeStartDate")
    start_date_input.clear()
    start_date_input.send_keys(start_date.strftime("%m/%d/%Y"))

    end_date_input = driver.find_element(By.ID, "rangeEndDate")
    end_date_input.clear()
    end_date_input.send_keys(end_date.strftime("%m/%d/%Y"))

    # Click search button
    search_button = driver.find_element(By.ID, "eventSearchButton")
    search_button.click()


def extract_bid_links(driver):
    """
    Extract bid links from the search results page.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.

    Returns:
        list: A list of tuples containing bid links and titles.
    """
    bid_links = []
    page_number = 1
    max_retries = 3

    while True:
        print(f"Current Page Number: {page_number}")

        for attempt in range(max_retries):
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "eventSearchTable"))
                )

                # Wait for the loading overlay to disappear
                WebDriverWait(driver, 10).until(
                    EC.invisibility_of_element_located((By.ID, "loadingDiv"))
                )

                rows = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//table[@id='eventSearchTable']/tbody/tr")
                    )
                )

                for row in rows:
                    link = (
                        WebDriverWait(row, 10)
                        .until(EC.presence_of_element_located((By.XPATH, ".//td[2]/a")))
                        .get_attribute("href")
                    )
                    title = (
                        WebDriverWait(row, 10)
                        .until(EC.presence_of_element_located((By.XPATH, ".//td[3]")))
                        .text
                    )
                    bid_links.append((link, title))

                break  # If successful, break the retry loop
            except (StaleElementReferenceException, TimeoutException) as e:
                if attempt < max_retries - 1:
                    print(f"Attempt {attempt + 1} failed. Retrying...")
                    time.sleep(2)
                else:
                    print(f"Failed to extract data after {max_retries} attempts.")
                    raise e

        try:
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "eventSearchTable_next"))
            )
            if "disabled" in (next_button.get_attribute("class") or ""):
                print("No more pages available.")
                break
            next_button.click()
            page_number += 1
            time.sleep(2)
        except (TimeoutException, StaleElementReferenceException):
            print("No more pages available.")
            break

    print("Bids with Posting Range within 2 Days Successfully Extracted")
    return bid_links


def extract_bid_details(driver, bid_link, bid_title):
    """
    Extract details from a specific bid page.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        bid_link (str): The URL of the bid detail page.
        bid_title (str): The title of the bid.

    Returns:
        dict: A dictionary containing the extracted bid details.
    """
    driver.get(bid_link)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "details-title"))
    )

    bid_details = {"Solicitation Title": bid_title, "Bid Detail Page URL": bid_link}

    # Extract details from the table
    table = driver.find_element(By.CLASS_NAME, "table.mobile-accordion")
    rows = table.find_elements(By.CLASS_NAME, "tr")
    for row in rows:
        cells = row.find_elements(By.CLASS_NAME, "td")
        for cell in cells:
            header = cell.get_attribute("data-header")
            if header == "Event ID:":
                bid_details["Solicitation Number"] = cell.text
            elif header == "Event Type:":
                bid_details["Notice Type"] = cell.text

    # Extract other details
    details_mapping = {
        "Start Date:": "Posted Date",
        "End Date:": "Response Date",
    }

    for key, value in details_mapping.items():
        try:
            element = driver.find_element(By.XPATH, f"//*[contains(text(), '{key}')]")
            date_text = (
                element.find_element(By.XPATH, "..").text.replace(f"{key} ", "").strip()
            )
            # Extract only the date part
            date_only = date_text.split("@")[0].strip()

            # Try different date formats
            date_formats = ["%m/%d/%Y", "%b %d, %Y", "%B %d, %Y"]
            parsed_date = None
            for date_format in date_formats:
                try:
                    parsed_date = datetime.strptime(date_only, date_format)
                    break
                except ValueError:
                    continue

            if parsed_date:
                formatted_date = parsed_date.strftime("%Y-%m-%d")
                bid_details[value] = formatted_date
            else:
                print(f"Unable to parse date: {date_only}")
                bid_details[value] = date_only
        except NoSuchElementException:
            bid_details[value] = ""

    # Extract Contact Information
    try:
        contact_info = driver.find_element(
            By.XPATH, "//h3[contains(text(), 'Buyer Contact:')]/following-sibling::p"
        ).text
        bid_details["Contact Information"] = contact_info.strip()
    except NoSuchElementException:
        bid_details["Contact Information"] = ""

    # Extract Agency
    try:
        agency = driver.find_element(By.XPATH, "//h4[contains(@class, 'mt-3')]").text
        bid_details["Agency"] = agency
    except NoSuchElementException:
        bid_details["Agency"] = ""

    # Extract Category (NIGP Code)
    try:
        nigp_code = driver.find_element(
            By.XPATH, "//table[contains(@class, 'nigp_table')]//tr[2]/td[1]"
        ).text
        nigp_desc = driver.find_element(
            By.XPATH, "//table[contains(@class, 'nigp_table')]//tr[2]/td[2]"
        ).text
        bid_details["Category"] = f"{nigp_code} - {nigp_desc}"
    except NoSuchElementException:
        bid_details["Category"] = ""

    # Extract Description
    try:
        description = driver.find_element(
            By.XPATH, "//p[contains(@class, 'paragraph description')]"
        ).text
        bid_details["Description"] = description
    except NoSuchElementException:
        bid_details["Description"] = ""

    # Extract Attachments
    try:
        attachments = driver.find_elements(By.XPATH, "//div[@class='document']//a")
        attachment_urls = [a.get_attribute("href") for a in attachments]
        bid_details["Attachments"] = ", ".join(attachment_urls)
    except NoSuchElementException:
        bid_details["Attachments"] = ""

    return bid_details


def requests_retry_session(
    retries=3, backoff_factor=0.3, status_forcelist=(500, 502, 504)
):
    session = requests.Session()
    retry = Retry(
        total=retries,
        read=retries,
        connect=retries,
        backoff_factor=backoff_factor,
        status_forcelist=status_forcelist,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


MAX_THREADS = 3  # Limit the number of concurrent threads
download_semaphore = threading.Semaphore(MAX_THREADS)


def clean_filename(filename):
    """
    Clean the filename by removing the 'yyyymm_supportingdocs_' prefix.

    Args:
        filename (str): The original filename.

    Returns:
        str: The cleaned filename.
    """
    parts = filename.split("_supportingdocs_")
    if len(parts) > 1:
        return "_".join(parts[1:])
    return filename


def download_file(driver, link, folder):
    with download_semaphore:
        try:
            original_file_name = link.text.strip()
            if not original_file_name:
                original_file_name = (
                    link.get_attribute("href").split("attachmentId=")[-1].split("&")[0]
                )

            print(f"Attempting to download file: {original_file_name}")

            # Clean the filename before downloading
            clean_file_name = clean_filename(original_file_name)

            # Set the download preferences to use the clean filename
            driver.execute_cdp_cmd(
                "Page.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": script_download_folder,
                    "eventsEnabled": True,
                },
            )

            # Use JavaScript to click the link and specify the filename
            driver.execute_script(
                """
                var link = arguments[0];
                var fileName = arguments[1];
                var clickEvent = document.createEvent('MouseEvents');
                clickEvent.initEvent('click', true, true);
                link.dispatchEvent(clickEvent);
                link.setAttribute('download', fileName);
            """,
                link,
                clean_file_name,
            )

            print(f"Clicked download link for {clean_file_name}")

            # Wait for the download to complete
            max_wait_time = 30
            wait_time = 0
            while wait_time < max_wait_time:
                if os.path.exists(
                    os.path.join(script_download_folder, clean_file_name)
                ):
                    print(f"Downloaded: {clean_file_name}")
                    # Move the file from script_download_folder to the bid folder
                    source_path = os.path.join(script_download_folder, clean_file_name)
                    destination_path = os.path.join(folder, clean_file_name)
                    safe_move(source_path, destination_path)
                    return clean_file_name
                time.sleep(1)
                wait_time += 1

            print(f"Download failed for: {clean_file_name}")
            return None
        except Exception as e:
            print(f"Error downloading file: {str(e)}")
            return None


def download_direct_links(driver, links, folder):
    downloaded_files = []
    threads = []
    for link in links:
        thread = threading.Thread(
            target=lambda: downloaded_files.append(download_file(driver, link, folder))
        )
        threads.append(thread)
        thread.start()
        time.sleep(1)  # Small delay between starting threads

    for thread in threads:
        thread.join()

    return [f for f in downloaded_files if f]


def download_from_buttons(driver, buttons, folder):
    downloaded_files = []
    threads = []
    for i, button in enumerate(buttons):
        thread = threading.Thread(
            target=lambda: downloaded_files.append(
                download_button_file(driver, button, i, folder)
            )
        )
        threads.append(thread)
        thread.start()
        time.sleep(1)  # Small delay between starting threads

    for thread in threads:
        thread.join()

    return [f for f in downloaded_files if f]


def download_button_file(driver, button, index, folder):
    with download_semaphore:
        try:
            file_name_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//span[@id='PV_ATTACH_WRK_ATTACHUSERFILE${index}']")
                )
            )
            file_name = file_name_element.text.strip()
            if not file_name:
                file_name = f"attachment_{index}"

            print(f"Attempting to download file: {file_name}")

            # Use ActionChains to click the button
            actions = ActionChains(driver)
            actions.move_to_element(button).click().perform()
            print(f"Clicked download button for {file_name}")

            # Wait for the download to complete
            time.sleep(20)  # Increased wait time

            # Move the downloaded file to the bid folder
            source_files = [
                f
                for f in os.listdir(main_folder)
                if f.lower().startswith(file_name.lower())
            ]
            if source_files:
                source_path = os.path.join(main_folder, source_files[0])
                _, extension = os.path.splitext(source_files[0])
                if not extension:
                    extension = ".bin"  # Default extension if none is present
                destination_file = f"{file_name}{extension}"
                destination_path = os.path.join(folder, destination_file)
                safe_move(source_path, destination_path)
                print(f"Downloaded and moved: {destination_file}")
                return destination_file
            else:
                print(f"Download failed for: {file_name}")
                return None
        except Exception as e:
            print(f"Error downloading file at index {index}: {str(e)}")
            return None


def download_attachments(driver, event_id):
    """
    Download attachments for a specific bid or get the event link.

    Args:
        driver (webdriver.Chrome): The WebDriver instance.
        event_id (str): The Event ID (Solicitation Number) of the bid.

    Returns:
        list: List of downloaded file names, or URL of the event link if present.
    """
    try:
        print(f"Attempting to download attachments for event ID: {event_id}")

        # Click on the Documents tab
        documents_tab = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "tab-documents"))
        )
        documents_tab.click()
        print("Clicked on Documents tab")

        # Wait for the content to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "collapse-documents"))
        )

        # Check for the event link
        event_link_element = driver.find_elements(
            By.XPATH,
            "//div[@id='collapse-documents']//a[contains(@href, 'bids.sciquest.com')]",
        )
        if event_link_element:
            event_url = event_link_element[0].get_attribute("href")
            print(f"Found event link: {event_url}")
            return [event_url]

        # Check if the document section is empty
        document_content = driver.find_element(By.ID, "collapse-documents").text.strip()
        if document_content == "Documents":
            print("No attachments found.")
            return []

        # Check for direct download links
        direct_links = driver.find_elements(By.XPATH, "//a[@class='DetattachDownLink']")
        if direct_links:
            print(f"Found {len(direct_links)} direct download links")

            # Only create bid folder if attachments are found
            bid_folder = os.path.join(script_folder, event_id)
            os.makedirs(bid_folder, exist_ok=True)

            downloaded_files = []
            for link in direct_links:
                try:
                    file_name = link.text.strip()
                    if not file_name:
                        file_name = (
                            link.get_attribute("href")
                            .split("attachmentId=")[-1]
                            .split("&")[0]
                        )
                    print(f"Attempting to download file: {file_name}")

                    # Use JavaScript to click the link
                    driver.execute_script("arguments[0].click();", link)
                    time.sleep(5)  # Wait for download to start
                except Exception as e:
                    print(f"Error clicking download link: {str(e)}")

            # Wait for downloads to complete and move files
            time.sleep(10)
            for file in os.listdir(script_download_folder):
                if (
                    os.path.isfile(os.path.join(script_download_folder, file))
                    and file != "07_StateOfGeorgia.xlsx"
                ):
                    source_path = os.path.join(script_download_folder, file)
                    clean_file_name = clean_filename(file)
                    destination_path = os.path.join(bid_folder, clean_file_name)
                    safe_move(source_path, destination_path)
                    downloaded_files.append(clean_file_name)

            return downloaded_files
        else:
            print("No attachments or links found.")
            return []

    except Exception as e:
        print(f"Error downloading attachments: {str(e)}")
        print(f"Current URL: {driver.current_url}")
        return []
    finally:
        print("Finished processing attachments")


def rename_bid_attachments(bid_folder):
    """
    Rename attachment files for a specific bid folder.

    Args:
        bid_folder (str): Path to the bid folder.

    Returns:
        list: List of renamed attachment filenames.
    """
    renamed_files = []
    for file in os.listdir(bid_folder):
        if "_supportingdocs_" in file:
            old_path = os.path.join(bid_folder, file)
            new_file_name = clean_filename(file)
            new_path = os.path.join(bid_folder, new_file_name)
            if old_path != new_path:
                try:
                    os.rename(old_path, new_path)
                    print(f"Renamed: {file} to {new_file_name}")
                    renamed_files.append(new_file_name)
                except Exception as e:
                    print(f"Error renaming {file}: {str(e)}")
                    renamed_files.append(file)
            else:
                renamed_files.append(file)
        else:
            renamed_files.append(file)
    return renamed_files


def update_excel(bid_data):
    """
    Update the Excel file with new bid data.

    Args:
        bid_data (dict): The bid details to be added to the Excel file.
    """
    # Save Excel file in the script_folder
    excel_file = os.path.join(script_folder, f"{script_name}.xlsx")

    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
    else:
        df = pd.DataFrame()

    # Ensure the Attachments field contains the actual filenames
    if "Attachments" in bid_data and isinstance(bid_data["Attachments"], list):
        bid_data["Attachments"] = ", ".join(bid_data["Attachments"])

    new_row = pd.DataFrame([bid_data])
    df = pd.concat([df, new_row], ignore_index=True)

    # Ensure all required columns are present
    required_columns = [
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

    for col in required_columns:
        if col not in df.columns:
            df[col] = ""

    df = df[required_columns]  # Reorder columns

    # Ensure Solicitation Title and Attachments are not empty
    df["Solicitation Title"] = df["Solicitation Title"].fillna("")
    df["Attachments"] = df["Attachments"].fillna("")

    df.to_excel(excel_file, index=False)
    print(f"Updated Excel file: {excel_file}")


# Add this new function to rename existing files in bid folders
def rename_existing_attachments(main_folder):
    """
    Rename existing attachment files in all bid folders.
    """
    for root, dirs, files in os.walk(main_folder):
        for file in files:
            if "_supportingdocs_" in file:
                old_path = os.path.join(root, file)
                new_file_name = clean_filename(file)
                new_path = os.path.join(root, new_file_name)
                if old_path != new_path:
                    try:
                        os.rename(old_path, new_path)
                        print(f"Renamed: {file} to {new_file_name}")
                    except Exception as e:
                        print(f"Error renaming {file}: {str(e)}")


def main():
    """
    Main function to execute the Georgia Procurement Registry bid extraction process.
    """
    driver = setup_driver()
    try:
        driver.get("https://ssl.doas.state.ga.us/gpr/index")

        set_search_criteria(driver)
        bid_links = extract_bid_links(driver)

        for index, (link, title) in enumerate(bid_links, start=1):
            try:
                print(f"Processing bid {index} of {len(bid_links)}: {title}")

                bid_details = extract_bid_details(driver, link, title)
                bid_details["SL No"] = index

                # Download attachments without creating folder initially
                attachments_result = download_attachments(
                    driver, bid_details["Solicitation Number"]
                )
                if attachments_result:
                    if isinstance(attachments_result[0], str) and attachments_result[
                        0
                    ].startswith("http"):
                        bid_details["Attachments"] = attachments_result[0]
                    else:
                        bid_details["Attachments"] = ", ".join(attachments_result)
                else:
                    bid_details["Attachments"] = ""

                update_excel(bid_details)

                time.sleep(5)
            except NoSuchWindowException:
                print(
                    f"Error: Browser window closed unexpectedly. Restarting browser for bid {index}"
                )
                play_notification_sound()
                input("Press Enter to continue...")
                driver.quit()
                driver = setup_driver()
                driver.get("https://ssl.doas.state.ga.us/gpr/index")
                continue
            except Exception as e:
                print(f"Error processing bid {index}: {str(e)}")
                play_notification_sound()
                input("Press Enter to continue...")
                continue

        print("All Bids and Attachments Extraction Successfully Completed")

        # Clean up temporary download folder before renaming
        try:
            if os.path.exists(script_download_folder):
                shutil.rmtree(script_download_folder)
                print(f"✅ Removed temporary download folder: {script_download_folder}")
        except Exception as e:
            print(f"⚠️ Error removing temporary folder: {str(e)}")

        # Rename folder to mark as completed
        completed_folder_name = f"{script_name}_COMPLETED"
        completed_folder = os.path.join(main_folder, completed_folder_name)
        if os.path.exists(completed_folder):
            shutil.rmtree(completed_folder)
        os.rename(script_folder, completed_folder)

        play_notification_sound()
    except Exception as e:
        print(f"An error occurred during execution: {str(e)}")
        play_notification_sound()
        input("Press Enter to continue...")
    finally:
        driver.quit()
        # Clean up the script-specific download folder
        shutil.rmtree(script_download_folder, ignore_errors=True)


if __name__ == "__main__":
    main()
