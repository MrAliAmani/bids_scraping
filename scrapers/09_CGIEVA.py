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
import random
from datetime import datetime, timedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    ElementNotInteractableException,
    StaleElementReferenceException,
)
from selenium.webdriver.common.action_chains import ActionChains
from fake_useragent import UserAgent
import pickle
import argparse
from utils.utils import safe_move, play_notification_sound
from selenium_stealth import stealth
import json


# Add this near the top of the script, after imports
def parse_arguments():
    parser = argparse.ArgumentParser(description="CGI eVA Bid Scraper")
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Number of days to search for bids (default: 2)",
    )
    return parser.parse_args()


# Update near the top of the script, after imports and before function definitions
script_name = "09_CGIEVA"
yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
yesterday_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), yesterday)
script_folder = os.path.join(
    yesterday_folder, f"{script_name}_IN_PROGRESS"
)  # Main folder with progress suffix
temp_download_folder = os.path.join(
    script_folder, script_name
)  # Temporary download folder without suffix

# Create necessary folders
os.makedirs(yesterday_folder, exist_ok=True)
os.makedirs(script_folder, exist_ok=True)
os.makedirs(temp_download_folder, exist_ok=True)


def setup_driver(download_directory):
    """Set up and configure the Chrome WebDriver with anti-bot measures."""
    ua = UserAgent()
    user_agent = ua.random
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-gpu")  # Disables GPU acceleration
    chrome_options.add_argument("--disable-software-rasterizer")  # Use CPU rendering
    chrome_options.add_argument(f"user-agent={user_agent}")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    # Add these options to prevent disconnection
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-session-crashed-bubble")
    chrome_options.add_argument("--disable-infobars")

    prefs = {
        "download.default_directory": download_directory,  # This will now point to temp_download_folder
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_settings.popups": 0,
        "power.keep_awake": True,  # Prevent system sleep
    }
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_options)

    # Remove the stealth code if you don't want to use selenium-stealth
    # If you want to keep it, make sure to install: pip install selenium-stealth
    try:
        stealth(
            driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
        )
    except NameError:
        print("Stealth mode not available - continuing without it")

    # Load cookies if available
    if os.path.exists("cookies.pkl"):
        cookies = pickle.load(open("cookies.pkl", "rb"))
        for cookie in cookies:
            driver.add_cookie(cookie)

    return driver


def human_like_mouse_move(driver, element):
    """Simulate human-like mouse movement to an element."""
    action = ActionChains(driver)
    action.move_to_element_with_offset(element, 5, 5)
    action.move_by_offset(10, 5)
    action.move_by_offset(-5, -5)
    action.perform()


def random_scroll(driver):
    """Perform random scrolling on the page."""
    scroll_amount = random.randint(100, 500)
    driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
    time.sleep(random.uniform(0.5, 1.5))


def is_visible(driver, element):
    """Check if an element is visible on the page."""
    return driver.execute_script(
        "var elem = arguments[0],                 "
        "  box = elem.getBoundingClientRect(),    "
        "  cx = box.left + box.width / 2,         "
        "  cy = box.top + box.height / 2,         "
        "  e = document.elementFromPoint(cx, cy); "
        "for (; e; e = e.parentElement) {         "
        "  if (e === elem)                        "
        "    return true;                         "
        "}                                        "
        "return false;                            ",
        element,
    )


def perform_advanced_search(driver):
    """Perform advanced search on the eVA website with anti-bot measures."""
    try:
        # Wait for the 'Open' status filter to be clickable
        open_status = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//li[contains(@class, 'facet-item-type-status') and contains(text(), 'Open')]",
                )
            )
        )

        # Check if the element is visible before interacting
        if is_visible(driver, open_status):
            human_like_mouse_move(driver, open_status)
            open_status.click()

        random_scroll(driver)
        time.sleep(random.uniform(8, 12))  # Randomized wait time
    except Exception as e:
        print(f"Error in advanced search: {str(e)}")


def extract_bid_links(driver, max_links=30):
    """Extract bid links from the search results page."""
    bid_links = []
    while len(bid_links) < max_links:
        try:
            # Wait for bid elements to be present on the page
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//div[contains(@class, 'card-body')]")
                )
            )
            bid_elements = driver.find_elements(
                By.XPATH, "//div[contains(@class, 'card-body')]"
            )
            for element in bid_elements:
                if len(bid_links) < max_links:
                    view_opportunity = element.find_element(
                        By.XPATH,
                        ".//span[contains(@class, 'btn-outline-primary') and contains(text(), 'View Opportunity')]",
                    )
                    bid_links.append(view_opportunity)
                else:
                    break

            if len(bid_links) < max_links:
                # Scroll to load more results
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(random.uniform(4, 6))
            else:
                break
        except Exception as e:
            print(f"Error extracting bid links: {str(e)}")
            break
    return bid_links


def extract_bid_details(driver, bid_element):
    """Extract details from a specific bid page."""
    bid_details = {
        "Notice Type": "",
        "Agency": "",
        "Category": "",
        "Additional Summary, if any": "",
        "Contracting Office Address": "",
        "Contact Information": "",
        "Posted Date": "",
        "Response Date": "",
    }

    def try_alternate_method(xpath):
        """Try alternate methods if primary method fails."""
        try:
            # Try JavaScript method
            element = driver.execute_script(
                f"return document.evaluate('{xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;"
            )
            if element:
                return element.textContent.strip()
        except:
            pass

        try:
            # Try explicit wait method
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            return element.text.strip()
        except:
            pass

        return ""

    try:
        # Primary click method
        try:
            driver.execute_script("arguments[0].scrollIntoView(true);", bid_element)
            time.sleep(random.uniform(1, 3))
            ActionChains(driver).move_to_element(bid_element).click().perform()
        except (StaleElementReferenceException, ElementNotInteractableException):
            # Fallback click method
            try:
                driver.execute_script("arguments[0].click();", bid_element)
            except:
                print("Failed to click bid element")
                return bid_details

        # Keep existing wait logic
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//li[contains(@class, 'list-group-item')]")
            )
        )

        # Try primary method for each field first, then fallback if needed
        try:
            posted_date = (
                WebDriverWait(driver, 10)
                .until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//li[contains(text(), 'Issued:')]/font")
                    )
                )
                .text.split()[0]
            )
        except:
            posted_date = try_alternate_method(
                "//li[contains(text(), 'Issued:')]/font"
            ).split()[0]

        if posted_date:
            try:
                bid_details["Posted Date"] = datetime.strptime(
                    posted_date, "%m/%d/%y"
                ).strftime("%Y-%m-%d")
            except:
                print(f"Invalid posted date format: {posted_date}")

        # Extract other fields with primary/fallback pattern
        field_mappings = {
            "Notice Type": "//li[contains(text(), 'Opportunity Type:')]/font",
            "Solicitation Number": "//h4[contains(@class, 'headerDocIdtext')]",
            "Solicitation Title": "//li[contains(@class, 'list-group-item docTitle')]",
            "Agency": "//li[contains(text(), 'Issued By:')]/font",
            "Category": "//li[contains(text(), 'Category:')]/font",
            "Description": "//li[contains(text(), 'Summary:')]/font",
            "Contracting Office Address": "//li[contains(text(), 'Work Location:')]/font",
            "Contact Information": "//li[contains(text(), 'Buyer:')]/font",
        }

        for field, xpath in field_mappings.items():
            try:
                # Try primary method first
                bid_details[field] = driver.find_element(By.XPATH, xpath).text.strip()
            except:
                # Only try alternate method if primary fails
                value = try_alternate_method(xpath)
                if value:
                    bid_details[field] = value

        bid_details["Bid Detail Page URL"] = driver.current_url

        # Try to find Pre-Bid Conference information
        try:
            pre_bid = driver.find_element(
                By.XPATH, "//li[contains(text(), 'Pre-Bid Conference:')]/font"
            ).text
            bid_details["Additional Summary, if any"] = pre_bid
        except NoSuchElementException:
            pass

    except Exception as e:
        print(f"Error extracting bid details: {str(e)}")

    return bid_details


def handle_captcha(driver):
    """Handle the captcha challenge."""
    try:
        print("Waiting for CAPTCHA to appear...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[@title='reCAPTCHA']"))
        )
        print("CAPTCHA detected. Please solve it manually.")
        play_notification_sound()  # Play notification sound
        input("Press Enter after solving the CAPTCHA...")
        time.sleep(random.uniform(5, 8))  # Wait a bit after solving
    except Exception as e:
        print(f"Error handling captcha: {str(e)}")


def download_attachments(driver, bid_number):
    """Download attachments for a specific bid."""
    attachments = []
    bid_folder = os.path.join(script_folder, bid_number)
    os.makedirs(bid_folder, exist_ok=True)

    try:
        attachment_table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "DataTables_Table_0"))
        )
        rows = attachment_table.find_elements(By.TAG_NAME, "tr")
        total_attachments = len(rows) - 1  # Subtract 1 for header row

        if total_attachments == 0:
            print(f"No attachments found for bid {bid_number}")
            return ""

        print(f"Total attachments found: {total_attachments}")

        for index, row in enumerate(rows[1:], start=1):  # Skip header row
            cells = row.find_elements(By.TAG_NAME, "td")
            if cells:
                attachment_link = cells[0].find_element(By.TAG_NAME, "a")
                attachment_name = attachment_link.text
                attachments.append(attachment_name)

                print(
                    f"Downloading attachment {index}/{total_attachments}: {attachment_name}"
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", attachment_link
                )
                time.sleep(random.uniform(2, 4))
                attachment_link.click()
                time.sleep(random.uniform(5, 7))

                try:
                    # First try to find and click the download button directly
                    try:
                        download_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "attachDownloadBtn"))
                        )
                        time.sleep(random.uniform(2, 4))
                        download_button.click()
                        print(f"Successfully initiated download for: {attachment_name}")
                        time.sleep(random.uniform(5, 7))
                        continue  # Skip captcha handling if download button was clickable
                    except TimeoutException:
                        pass  # If download button not immediately clickable, proceed to captcha check

                    # Check for CAPTCHA iframe
                    captcha_iframe = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//iframe[@title='reCAPTCHA']")
                        )
                    )

                    # CAPTCHA found - check if it's checkbox or image
                    driver.switch_to.frame(captcha_iframe)

                    try:
                        # Try to click the checkbox CAPTCHA
                        checkbox = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, "recaptcha-anchor"))
                        )
                        checkbox.click()
                        time.sleep(random.uniform(2, 4))

                        # Switch back to main content
                        driver.switch_to.default_content()

                        # Try to click download button after checkbox CAPTCHA
                        download_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "attachDownloadBtn"))
                        )
                        download_button.click()
                        print(
                            f"Successfully initiated download after checkbox CAPTCHA for: {attachment_name}"
                        )
                        time.sleep(random.uniform(5, 7))

                    except (TimeoutException, ElementNotInteractableException):
                        # If checkbox click fails, it's probably an image CAPTCHA
                        driver.switch_to.default_content()
                        print(
                            "Image-based CAPTCHA detected. Manual intervention required."
                        )
                        play_notification_sound()
                        input("Press Enter after solving the CAPTCHA...")

                        # Click download button after manual CAPTCHA solution
                        download_button = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.ID, "attachDownloadBtn"))
                        )
                        download_button.click()
                        print(
                            f"Successfully initiated download after manual CAPTCHA for: {attachment_name}"
                        )
                        time.sleep(random.uniform(5, 7))

                except TimeoutException:
                    # No CAPTCHA found - try direct download
                    download_button = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.ID, "attachDownloadBtn"))
                    )
                    download_button.click()
                    print(
                        f"Successfully initiated direct download for: {attachment_name}"
                    )
                    time.sleep(random.uniform(5, 7))

                except Exception as e:
                    print(
                        f"Failed to download attachment: {attachment_name}. Error: {str(e)}"
                    )

        print(f"All attachments for bid {bid_number} processed.")
        return ", ".join(attachments)

    except TimeoutException:
        print(f"No attachments found for bid {bid_number}")
        return ""
    except Exception as e:
        print(f"Error processing attachments for bid {bid_number}: {str(e)}")
        return ""


def update_excel(bid_data):
    """Update the Excel file with new bid data."""
    excel_file = os.path.join(
        script_folder, "09_CGIEVA.xlsx"
    )  # Excel file goes in script_folder
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
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

    # Check if the Solicitation Number already exists
    existing_row = df[df["Solicitation Number"] == bid_data["Solicitation Number"]]

    if not existing_row.empty:
        # Update the existing row
        index = existing_row.index[0]
        for key, value in bid_data.items():
            df.at[index, key] = value
    else:
        # Add a new row
        df = pd.concat([df, pd.DataFrame([bid_data])], ignore_index=True)

    # Update SL No to be sequential starting from 1
    df["SL No"] = range(1, len(df) + 1)

    # Update Attachments column
    solicitation_number = bid_data["Solicitation Number"]
    attachments = get_attachment_filenames(solicitation_number)
    df.loc[df["Solicitation Number"] == solicitation_number, "Attachments"] = (
        attachments
    )

    df.to_excel(excel_file, index=False)
    print(f"Updated 09_CGIEVA.xlsx with bid {bid_data['Solicitation Number']}")


def get_attachment_filenames(bid_number):
    """Get the filenames of attachments in the bid folder."""
    bid_folder = os.path.join(script_folder, bid_number)
    if os.path.exists(bid_folder):
        files = [
            file
            for file in os.listdir(bid_folder)
            if os.path.isfile(os.path.join(bid_folder, file))
        ]
        return ", ".join(files)
    return ""


def update_attachments_in_excel():
    """Update the Attachments column in the Excel file after all files have been moved."""
    excel_file = os.path.join(script_folder, "09_CGIEVA.xlsx")
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        for index, row in df.iterrows():
            solicitation_number = row["Solicitation Number"]
            attachments = get_attachment_filenames(solicitation_number)
            if attachments:
                df.at[index, "Attachments"] = attachments
        df.to_excel(excel_file, index=False)
        print("Updated Attachments column in 09_CGIEVA.xlsx")


def move_remaining_files(folder, bid_number):
    """Move any leftover attachments to their respective bid folders."""
    bid_folder = os.path.join(
        script_folder, bid_number
    )  # Bid folders go in script_folder
    if not os.path.exists(bid_folder):
        os.makedirs(bid_folder)
    for file in os.listdir(folder):
        file_path = os.path.join(folder, file)
        if os.path.isfile(file_path):
            safe_move(file_path, os.path.join(bid_folder, file))
            print(f"Moved {file} to {bid_number}")

    # Remove bid folder if it's empty
    if os.path.exists(bid_folder) and not os.listdir(bid_folder):
        os.rmdir(bid_folder)
        print(f"Removed empty folder for bid {bid_number}")


def cleanup_script_folder():
    """Remove the script-specific folder and any temporary files."""
    if os.path.exists(script_folder):
        # Clean up any temporary download files
        if os.path.exists(temp_download_folder):
            for item in os.listdir(temp_download_folder):
                item_path = os.path.join(temp_download_folder, item)
                try:
                    if os.path.isfile(item_path):
                        os.remove(item_path)
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                except Exception as e:
                    print(f"Error removing {item_path}: {e}")

            try:
                os.rmdir(temp_download_folder)
            except Exception as e:
                print(f"Error removing temporary download folder: {e}")

        # Rename folder to indicate completion
        completed_folder = os.path.join(yesterday_folder, f"{script_name}_COMPLETED")
        if os.path.exists(completed_folder):
            shutil.rmtree(completed_folder)
        os.rename(script_folder, completed_folder)
        print(f"Renamed folder to indicate completion: {completed_folder}")


def keep_session_alive(driver):
    """Perform actions to keep the session alive."""
    try:
        # Scroll slightly to simulate activity
        driver.execute_script("window.scrollBy(0, 1);")
        driver.execute_script("window.scrollBy(0, -1);")

        # Move mouse to prevent screen saver
        action = ActionChains(driver)
        action.move_by_offset(1, 1).perform()
        action.move_by_offset(-1, -1).perform()

    except Exception as e:
        print(f"Error in keep_session_alive: {e}")


def process_bid(driver, bid_url):
    """Process a single bid with session keeping."""
    try:
        # Start a background thread to keep the session alive
        import threading

        stop_flag = threading.Event()

        def session_keeper():
            while not stop_flag.is_set():
                keep_session_alive(driver)
                time.sleep(60)  # Run every minute

        keeper_thread = threading.Thread(target=session_keeper)
        keeper_thread.daemon = True
        keeper_thread.start()

        # Your existing bid processing code here
        # ...

        # Stop the session keeper when done
        stop_flag.set()
        keeper_thread.join(timeout=1)

    except Exception as e:
        print(f"Error processing bid: {e}")


def retry_on_stale_element(func, max_retries=3, delay=2):
    """Retry function on StaleElementReferenceException."""
    for attempt in range(max_retries):
        try:
            return func()
        except StaleElementReferenceException:
            if attempt == max_retries - 1:
                raise
            time.sleep(delay)
        except Exception as e:
            print(f"Error in retry_on_stale_element: {e}")
            raise


def reconnect_if_needed(driver):
    """Check connection and reconnect if necessary."""
    try:
        # Test if driver is responsive
        driver.current_url
        return driver
    except:
        print("WebDriver disconnected. Attempting to reconnect...")
        try:
            driver.quit()
        except:
            pass
        return setup_driver()


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


def main():
    """Main function to execute the eVA bid extraction process."""
    print("🚀 Bids Extraction Started")

    # Parse command-line arguments
    args = parse_arguments()
    days_to_search = args.days
    print(f"Searching for bids posted in the last {days_to_search} days")

    driver = setup_driver(
        temp_download_folder
    )  # Use temp_download_folder for initial downloads
    driver.get("https://mvendor.cgieva.com/Vendor/public/AllOpportunities.jsp")

    try:
        perform_advanced_search(driver)
        bid_links = extract_bid_links(driver)
        total_bids = len(bid_links)
        print(f"📊 Total bid links extracted: {total_bids}")

        cutoff_date = datetime.now() - timedelta(days=days_to_search)

        for index, bid_element in enumerate(bid_links, start=1):
            try:
                print(f"Processing bid {index}/{total_bids}...")

                # Add random scrolling and delay before clicking
                random_scroll(driver)
                time.sleep(random.uniform(2, 4))

                # Click the bid element first
                bid_details = extract_bid_details(driver, bid_element)

                # Now get the URL after the page has loaded
                bid_url = driver.current_url

                # Add cache check here
                if not should_process_bid_link(bid_url):
                    driver.back()
                    time.sleep(random.uniform(10, 15))
                    continue

                if (
                    bid_details
                    and "Posted Date" in bid_details
                    and bid_details["Posted Date"]
                ):
                    posted_date = datetime.strptime(
                        bid_details["Posted Date"], "%Y-%m-%d"
                    )
                    if posted_date < cutoff_date:
                        print(
                            f"Bid {index}/{total_bids} is older than {days_to_search} days. Skipping..."
                        )
                        driver.back()
                        time.sleep(random.uniform(10, 15))
                        continue

                    bid_number = bid_details.get("Solicitation Number", "")
                    bid_folder = os.path.join(script_folder, bid_number)
                    os.makedirs(bid_folder, exist_ok=True)
                    print(f"Downloading attachments for bid {bid_number}")

                    # Update download directory for this specific bid
                    driver.command_executor._commands["send_command"] = (
                        "POST",
                        "/session/$sessionId/chromium/send_command",
                    )
                    params = {
                        "cmd": "Page.setDownloadBehavior",
                        "params": {"behavior": "allow", "downloadPath": bid_folder},
                    }
                    driver.execute("send_command", params)

                    attachments = download_attachments(driver, bid_number)
                    if attachments:
                        bid_details["Attachments"] = attachments
                    else:
                        bid_details["Attachments"] = "No attachments"

                    bid_details["SL No"] = str(index)
                    bid_details["Bid Detail Page URL"] = (
                        bid_url  # Add the URL to bid details
                    )

                    move_remaining_files(script_folder, bid_number)

                    time.sleep(random.uniform(3, 5))

                    update_excel(bid_details)

                    # Save to cache after successful processing
                    save_to_cache(bid_url, bid_details)

                    print(f"Navigating back to the main page...")
                    driver.back()
                    time.sleep(random.uniform(10, 15))
                else:
                    print(
                        f"Bid {index}/{total_bids} has missing or invalid data. Skipping..."
                    )
                    driver.back()
                    time.sleep(random.uniform(10, 15))
            except ElementNotInteractableException:
                print(f"Interaction error for bid {index}/{total_bids}. Skipping...")
                play_notification_sound()
                input("Press Enter to continue...")
                driver.back()
                time.sleep(random.uniform(10, 15))
            except StaleElementReferenceException:
                print(
                    f"Stale element reference for bid {index}/{total_bids}. Skipping..."
                )
                play_notification_sound()
                input("Press Enter to continue...")
                driver.back()
                time.sleep(random.uniform(10, 15))
            except Exception as e:
                print(
                    f"Error processing bid {index}/{total_bids}: {str(e)}. Skipping..."
                )
                play_notification_sound()
                input("Press Enter to continue...")
                driver.back()
                time.sleep(random.uniform(10, 15))

        print("🎉 All Bids and Attachments Extraction Successfully Completed")
        play_notification_sound()  # Play notification sound on completion

    except Exception as e:
        print(f"❗ An error occurred: {str(e)}")
        play_notification_sound()  # Play notification sound on error
        input("Press Enter to continue...")

    finally:
        # Save cookies before quitting
        pickle.dump(driver.get_cookies(), open("cookies.pkl", "wb"))
        driver.quit()
        update_attachments_in_excel()
        cleanup_script_folder()  # Clean up the script-specific folder


if __name__ == "__main__":
    main()
