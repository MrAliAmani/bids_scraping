import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import functools

# Override the print function to always flush
print = functools.partial(print, flush=True)

import os

# Add the parent directory to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import sys
import time
import argparse
from datetime import datetime, timedelta
import random
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium_stealth import stealth
import pandas as pd
import winsound
from fake_useragent import UserAgent
from selenium.webdriver.support.ui import Select
from functools import wraps
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import re
import pyautogui
from utils.utils import play_notification_sound, safe_move
from fuzzywuzzy import fuzz
import shutil

# Configuration
DEFAULT_DAYS = 2

# Add this constant at the beginning of the script
DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

# Add this constant at the top of the file with other constants
CUSTOMER_URLS = [
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=31",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=36",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=37",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=38",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=41",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=46",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=48",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=51",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=52",
    "https://webprocure.proactiscloud.com/wp-web-public/en/#/bidboard/search?searchterm=*&customerid=53",
]


def retry(exceptions, tries=4, delay=3, backoff=2):
    def deco_retry(f):
        @wraps(f)
        def f_retry(*args, **kwargs):
            mtries, mdelay = tries, delay
            while mtries > 1:
                try:
                    return f(*args, **kwargs)
                except exceptions as e:
                    print(f"{str(e)}, Retrying in {mdelay} seconds...")
                    time.sleep(mdelay)
                    mtries -= 1
                    mdelay *= backoff
            return f(*args, **kwargs)

        return f_retry

    return deco_retry


def parse_arguments():
    parser = argparse.ArgumentParser(description="Scrape State of Connecticut BidBoard")
    parser.add_argument(
        "--days",
        type=int,
        default=DEFAULT_DAYS,
        help="Number of days for bid date range",
    )
    return parser.parse_args()


def create_directory_structure():
    """Create the required directory structure for the script"""
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    script_name = os.path.splitext(os.path.basename(__file__))[0]

    # Create main folder with IN_PROGRESS suffix in yesterday's folder
    main_dir = os.path.join(os.getcwd(), yesterday, f"{script_name}_IN_PROGRESS")
    os.makedirs(main_dir, exist_ok=True)

    # Create temporary download folder inside main folder
    script_download_dir = os.path.join(main_dir, script_name)
    os.makedirs(script_download_dir, exist_ok=True)

    print(f"\n📂 Created directory structure:")
    print(f"Main directory: {main_dir}")
    print(f"Download directory: {script_download_dir}")

    return main_dir, script_download_dir


def init_driver(download_dir):
    chrome_options = Options()

    # Enhanced anti-bot detection settings
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")

    # Add window size randomization
    window_sizes = [(1366, 768), (1920, 1080), (1536, 864), (1440, 900)]
    window_size = random.choice(window_sizes)
    chrome_options.add_argument(f"--window-size={window_size[0]},{window_size[1]}")

    # Additional anti-detection arguments
    chrome_options.add_argument("--disable-automation")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-logging")
    chrome_options.add_argument("--no-default-browser-check")
    chrome_options.add_argument("--no-first-run")
    chrome_options.add_argument("--allow-running-insecure-content")

    # Random timezone
    timezones = [
        "America/New_York",
        "America/Chicago",
        "America/Denver",
        "America/Los_Angeles",
    ]
    chrome_options.add_argument(f"--timezone={random.choice(timezones)}")

    # Enhanced experimental options
    chrome_options.add_experimental_option("useAutomationExtension", False)
    chrome_options.add_experimental_option(
        "excludeSwitches", ["enable-automation", "enable-logging"]
    )

    # Download preferences
    chrome_options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.notifications": 2,
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "plugins.always_open_pdf_externally": True,
            # Additional preferences to make browser appear more human-like
            "profile.default_content_settings.popups": 0,
            "profile.managed_default_content_settings.images": 1,
            "profile.default_content_setting_values.cookies": 1,
            "profile.default_content_setting_values.plugins": 1,
            "profile.default_content_setting_values.geolocation": 2,
            "profile.default_content_setting_values.media_stream": 2,
        },
    )

    # Enhanced user agent rotation
    ua = UserAgent()
    user_agents = [ua.chrome, ua.firefox, ua.edge]  # Use common browser user agents
    chrome_options.add_argument(f"user-agent={random.choice(user_agents)}")

    driver = webdriver.Chrome(options=chrome_options)

    # Enhanced stealth settings
    stealth(
        driver,
        languages=["en-US", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
        run_on_insecure_origins=True,
    )

    # Additional JavaScript to modify navigator properties
    driver.execute_script(
        """
        // Overwrite the navigator properties
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        });
        
        // Add plugins
        Object.defineProperty(navigator, 'plugins', {
            get: () => [
                {name: 'Chrome PDF Plugin'},
                {name: 'Chrome PDF Viewer'},
                {name: 'Native Client'}
            ]
        });
        
        // Add languages
        Object.defineProperty(navigator, 'languages', {
            get: () => ['en-US', 'en']
        });
        
        // Add platform
        Object.defineProperty(navigator, 'platform', {
            get: () => 'Win32'
        });
        
        // Add touch support
        Object.defineProperty(navigator, 'maxTouchPoints', {
            get: () => 0
        });
    """
    )

    return driver


def start_process(driver, url):
    """Modified to accept URL parameter and handle no results case"""
    max_retries = 3
    retry_count = 0

    while retry_count < max_retries:
        try:
            print(f"\nAttempting to load URL: {url}")
            driver.get(url)

            # Wait for initial page load
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            time.sleep(5)  # Additional wait for page stability

            # Check for error message
            error_script = """
            function checkForError() {
                const errorSelectors = [
                    'div.error-message',
                    'div.alert-error',
                    'div.unexpected-error'
                ];
                
                for (const selector of errorSelectors) {
                    const element = document.querySelector(selector);
                    if (element && element.textContent.includes('Well this was unexpected')) {
                        return true;
                    }
                }
                return false;
            }
            return checkForError();
            """

            if driver.execute_script(error_script):
                print("\n⚠️ Error: Well this was unexpected. Something went wrong.")
                retry_count += 1
                if retry_count < max_retries:
                    print(f"Retrying... (Attempt {retry_count + 1} of {max_retries})")
                    time.sleep(5)
                    continue
                else:
                    print("Max retries reached. Moving to next URL.")
                    return False

            # Check for "No search results" message
            if check_no_results(driver):
                print("No search results found, refreshing page...")
                retry_count += 1
                driver.refresh()
                time.sleep(5)
                continue

            # Wait for the status checkbox to appear
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "facet-status-2"))
            )
            print("Page loaded successfully")
            return True

        except Exception as e:
            print("\n⚠️ Error: Well this was unexpected. Something went wrong.")
            retry_count += 1
            if retry_count < max_retries:
                print(f"Retrying... (Attempt {retry_count + 1} of {max_retries})")
                time.sleep(5)
            else:
                print("Max retries reached. Moving to next URL.")
                return False

    return False


def apply_filters(driver):
    try:
        print("Applying filters...")
        # Increased initial wait before applying filters
        time.sleep(5)

        # 1. Click the "Open" status checkbox with value verification
        checkbox_script = """
        function clickCheckboxWithDelay() {
            return new Promise((resolve) => {
                setTimeout(() => {
                    const checkbox = document.querySelector('#facet-status-2');
                    if (checkbox) {
                        // Get the associated label text
                        const label = document.querySelector('label[for="facet-status-2"]');
                        const labelText = label ? label.textContent.trim().toLowerCase() : '';
                        
                        // Check both the value and label text
                        if (labelText.includes('open')) {
                            checkbox.click();
                            resolve("✅ 'Open' status checkbox clicked successfully.");
                        } else {
                            resolve("⚠️ Could not verify 'Open' status checkbox");
                        }
                    } else {
                        resolve("❌ Status checkbox not found");
                    }
                }, 3000);  // Increased delay to 3 seconds
            });
        }
        return clickCheckboxWithDelay();
        """
        checkbox_result = driver.execute_script(checkbox_script)
        print(checkbox_result)

        if "not found" in checkbox_result or "Could not verify" in checkbox_result:
            print("⚠️ Retrying filter application...")
            time.sleep(5)  # Increased retry wait time
            checkbox_result = driver.execute_script(checkbox_script)
            print(checkbox_result)

        # Increased wait time after checkbox click
        time.sleep(5)

        # 2. Click the "Apply Filter" button with verification
        apply_filter_script = """
        function clickApplyFilterIfVisible() {
            return new Promise((resolve) => {
                const button = document.querySelector('#search-facets > button');
                if (button) {
                    const style = window.getComputedStyle(button);
                    if (style.display !== 'none' && style.visibility !== 'hidden') {
                        button.click();
                        resolve("✅ 'Apply Filter' button clicked successfully.");
                    } else {
                        resolve("🔍 Filter button not visible");
                    }
                } else {
                    resolve("❌ Filter button not found");
                }
            });
        }
        return clickApplyFilterIfVisible();
        """
        apply_filter_result = driver.execute_script(apply_filter_script)
        print(apply_filter_result)

        # Wait for the page to update after applying filters
        time.sleep(5)

        # 3. Set sort order to "Start Date (Descending)"
        sort_script = """
        function setSortOrderDescending() {
            const selectWrapper = document.querySelector('.select-wrapper');
            if (!selectWrapper) return "Select wrapper not found";

            const dropdown = selectWrapper.querySelector('.select-dropdown');
            if (!dropdown) return "Dropdown not found";

            // Trigger dropdown to open
            dropdown.click();

            // Find and click the "Start Date (Descending)" option
            const options = document.querySelectorAll('.select-dropdown li');
            for (let option of options) {
                if (option.textContent.trim() === "Start Date (Descending)") {
                    option.click();
                    return "��� Sort order set to 'Start Date (Descending)'";
                }
            }

            return "❌ Option 'Start Date (Descending)' not found";
        }
        return setSortOrderDescending();
        """
        sort_result = driver.execute_script(sort_script)
        print(f"Sorting result: {sort_result}")

        if "not found" in sort_result:
            print("⚠️ Warning: Could not set sort order to 'Start Date (Descending)'")
            play_notification_sound()
            input("Please set the sort order manually, then press Enter to continue...")
            return

        # Wait for the page to update after sorting
        time.sleep(5)

        print("Filters and sorting applied successfully.")
    except Exception as e:
        print(f"Error in apply_filters: {str(e)}")
        play_notification_sound()
        input("Please apply filters manually, then press Enter to continue...")


@retry(Exception, tries=3, delay=1)
def sort_using_select(driver):
    select_element = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "select[data-select-id]"))
    )
    select = Select(select_element)
    select.select_by_visible_text("Start Date (Descending)")


def sort_using_javascript(driver):
    script = """
    const selectElement = document.querySelector('select[data-select-id]');
    if (selectElement) {
        const option = Array.from(selectElement.options).find(opt => 
            opt.text.includes("Start Date (Descending)"));
        if (option) {
            selectElement.value = option.value;
            selectElement.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
    }
    return false;
    """
    result = driver.execute_script(script)
    if not result:
        raise Exception("JavaScript sorting failed")


def sort_using_action_chains(driver):
    select_element = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "select[data-select-id]"))
    )
    ActionChains(driver).move_to_element(select_element).click().perform()
    option = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, "//option[contains(text(), 'Start Date (Descending)')]")
        )
    )
    option.click()


def extract_bids(driver, days, main_dir, script_download_dir):
    bids = []
    cutoff_date = datetime.now() - timedelta(days=days)

    try:
        print("\n📊 Starting bid extraction process...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        time.sleep(5)  # Add a short delay to ensure the page has fully loaded

        if "bidboard/search" not in driver.current_url:
            print("⚠️ Page has changed unexpectedly. Returning to bid board.")
            driver.get(driver.current_url)
            apply_filters(driver)

        print("🔍 Waiting for bid elements to load...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span.title > a"))
        )
    except Exception as e:
        print(f"❌ Error waiting for bid elements: {e}")
        return bids

    bid_data_script = """
    function getBidData() {
        const bidElements = document.querySelectorAll('span.title > a');
        if (!bidElements || bidElements.length === 0) return null;
        return Array.from(bidElements).map(element => {
            const row = element.closest('li');
            const dateElement = row.querySelector('span.text-small.teal-text.darken-4');
            const dateText = dateElement ? dateElement.textContent : '';
            const startDate = dateText.split('Start Date:')[1]?.split('End Date:')[0]?.trim() || '';
            const endDate = dateText.split('End Date:')[1]?.trim() || '';
            const agency = row.querySelector('span.text-small.grey-text')?.textContent.trim() || '';
            return {
                url: element.href,
                startDate: startDate,
                endDate: endDate,
                agency: agency,
                title: element.textContent.trim()
            };
        });
    }
    return getBidData();
    """

    while True:  # Loop to handle pagination
        try:
            print("\n📑 Extracting bid data from current page...")
            bid_data = driver.execute_script(bid_data_script)
            if bid_data is None or len(bid_data) == 0:
                print("❌ No bid elements found on the page.")
                break
            print(f"✅ Found {len(bid_data)} bids on current page")
        except Exception as e:
            print(f"❌ Error executing bid data script: {e}")
            break

        stop_extraction = False
        for index, bid in enumerate(bid_data):
            try:
                if not bid["startDate"]:
                    print(f"⚠️ Warning: No start date for bid {bid['title']}")
                    continue

                start_date = bid["startDate"]
                try:
                    parsed_start_date = datetime.strptime(start_date, "%b %d, %Y")
                    if parsed_start_date < cutoff_date:
                        print(
                            f"\n🕒 Bid start date {start_date} is before cutoff date {cutoff_date}."
                        )
                        stop_extraction = True
                        break
                except ValueError as e:
                    print(f"⚠️ Error parsing date {start_date}: {e}")
                    continue

                print(
                    f"\n🔄 Processing bid {index + 1}/{len(bid_data)}: {bid['title']}"
                )
                print(f"📅 Start Date: {start_date}")
                print(f"🏢 Agency: {bid['agency']}")

                click_script = f"""
                document.querySelectorAll('span.title > a')[{index}].click();
                """
                driver.execute_script(click_script)

                print("⏳ Waiting for bid details page to load...")
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                time.sleep(5)

                bid_details = scrape_bid_details(
                    driver,
                    main_dir,
                    start_date,
                    bid["endDate"],
                    bid["agency"],
                    script_download_dir,
                )
                if bid_details:
                    bids.append(bid_details)
                    print(
                        f"✅ Successfully scraped bid: {bid_details['Solicitation Number']}"
                    )
                    move_remaining_files_for_bid(
                        main_dir,
                        script_download_dir,
                        bid_details["Solicitation Number"],
                    )

                print("🔄 Returning to main page...")
                driver.execute_script("window.history.go(-1)")
                time.sleep(3)

                print("🔄 Re-applying filters and sorting...")
                apply_filters(driver)
                verify_sort_order(driver)

            except Exception as e:
                print(f"❌ Error processing bid: {e}")
                continue

        if stop_extraction:
            print("\n🛑 Stopping extraction due to reaching cutoff date")
            break

        print("\n🔍 Checking for next page...")
        next_page_script = """
        const nextButton = document.querySelector('a.waves-effect i.material-icons');
        return nextButton && nextButton.textContent.trim() === 'chevron_right';
        """
        has_next_page = driver.execute_script(next_page_script)
        if not has_next_page:
            print(" No more pages to scrape.")
            break

        try:
            print("➡️ Moving to next page...")
            driver.execute_script(
                """
            document.querySelector('a.waves-effect i.material-icons').click();
            """
            )
            time.sleep(5)
        except Exception as e:
            print(f"❌ Error navigating to next page: {e}")
            break

    print(f"\n📊 Total bids extracted: {len(bids)}")
    return bids


def apply_sort_order_descending(driver):
    sort_script = """
    function setSortOrderDescending() {
        const selectWrapper = document.querySelector('.select-wrapper');
        if (!selectWrapper) return "Select wrapper not found";

        const dropdown = selectWrapper.querySelector('.select-dropdown');
        if (!dropdown) return "Dropdown not found";

        // Trigger dropdown to open
        dropdown.click();

        // Find and click the "Start Date (Descending)" option
        const options = document.querySelectorAll('.select-dropdown li');
        for (let option of options) {
            if (option.textContent.trim() === "Start Date (Descending)") {
                option.click();
                return "Sort order set to 'Start Date (Descending)'";
            }
        }

        return "Option 'Start Date (Descending)' not found";
    }
    return setSortOrderDescending();
    """
    sort_result = driver.execute_script(sort_script)
    print(f"Sorting result: {sort_result}")
    time.sleep(5)  # Wait for the sorting to take effect


def normalize_filename(filename):
    # Remove leading/trailing spaces and collapse multiple spaces into one
    return re.sub(r"\s+", " ", filename.strip())


def scrape_bid_details(
    driver, main_dir, start_date, end_date, agency, script_download_dir
):
    try:
        print("\nScraping bid details:")
        start_date = datetime.strptime(start_date, "%b %d, %Y").strftime("%Y-%m-%d")
        end_date = datetime.strptime(end_date, "%b %d, %Y").strftime("%Y-%m-%d")
        print(f"Start Date: {start_date}")
        print(f"End Date: {end_date}")
        print(f"Agency: {agency}")

        script = """
        function extractBidDetails() {
            const details = {};
            
            // Solicitation Number
            const solNumElement = document.evaluate("/html/body/div[1]/app-root/app-bid-board/app-bid-board-details/div[2]/div/div[1]/div[1]/h5/u/b", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            details.solicitationNumber = solNumElement ? solNumElement.textContent.trim() : '';
            
            // Solicitation Title
            details.solicitationTitle = document.querySelector('h5.teal-text.heading b')?.textContent.trim() || '';
            
            // Categories - with validation
            const categoryRows = Array.from(document.querySelectorAll('.datatable-row-center.datatable-row-group'));
            if (categoryRows.length >= 2) {
                const categories = categoryRows.slice(0, 2).map(row => {
                    const cells = row.querySelectorAll('datatable-body-cell');
                    if (cells.length >= 2) {
                        const code = cells[0].textContent.trim();
                        const desc = cells[1].textContent.trim();
                        // Only include if both code and description are meaningful
                        if (code && desc && code !== '1' && desc !== '1') {
                            return `${code}, ${desc}`;
                        }
                    }
                    return null;
                }).filter(cat => cat !== null);
                
                details.categories = categories.length > 0 ? categories.join('; ') : '';
            } else {
                details.categories = '';
            }
            
            // Description
            const descriptionElement = document.evaluate("/html/body/div[1]/app-root/app-bid-board/app-bid-board-details/div[2]/div/div[2]/div/p", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            details.description = descriptionElement ? descriptionElement.textContent.trim() : '';
            
            // Contracting Office Address
            const addressSelector = "#webprocure_public_contract_board > app-root > app-bid-board > app-bid-board-details > div:nth-child(3) > div > div.col.m5 > div > p";
            const addressElement = document.querySelector(addressSelector);
            details.contractingOfficeAddress = addressElement ? addressElement.innerText.replace(/\\n/g, ' ') : '';
            
            // Attachments
            details.attachments = Array.from(document.querySelectorAll('a.pointer')).map(a => ({
                name: a.textContent.trim(),
                element: a
            }));
            
            return details;
        }
        return extractBidDetails();
        """
        bid_details = driver.execute_script(script)

        for key, value in bid_details.items():
            if key != "attachments":
                print(f"{key.capitalize()}: {value}")

        if not bid_details["solicitationNumber"]:
            print(
                f"Warning: Could not extract solicitation number for bid URL: {driver.current_url}"
            )
            bid_details["solicitationNumber"] = "Unknown"

        folder_name = bid_details["solicitationNumber"].replace("/", "_")

        print(f"Extracting attachments for bid: {folder_name}")
        attachments = extract_attachments(
            driver, folder_name, main_dir, script_download_dir
        )

        bid_data = {
            "Posted Date": start_date,
            "Response Date": end_date,
            "Notice Type": "",
            "Solicitation Number": bid_details["solicitationNumber"],
            "Solicitation Title": bid_details["solicitationTitle"],
            "Agency": agency,
            "Category": bid_details[
                "categories"
            ],  # This will now be empty if no valid categories
            "Description": bid_details["description"],
            "Additional Summary": "",
            "Contracting Office Address": bid_details["contractingOfficeAddress"],
            "Contact Information": "",
            "Bid Detail Page URL": driver.current_url,
            "Attachments": attachments,
        }

        print(
            f"Successfully extracted bid details: {bid_details['solicitationNumber']}"
        )

        save_to_excel([bid_data], main_dir, append=True)

        return bid_data
    except Exception as e:
        print(f"Error scraping bid details for {driver.current_url}: {e}")
        import traceback

        traceback.print_exc()
        return None


def find_matching_file(directory, target_filename, timeout=300):
    target_filename = target_filename.lower()
    start_time = time.time()

    while time.time() - start_time < timeout:
        best_match = None
        highest_ratio = 0

        for filename in os.listdir(directory):
            if filename.endswith(".crdownload"):
                continue  # Skip incomplete downloads

            if os.path.isfile(os.path.join(directory, filename)):
                ratio = fuzz.ratio(target_filename, filename.lower())
                if ratio > highest_ratio:
                    highest_ratio = ratio
                    best_match = filename

        if highest_ratio > 80:  # You can adjust this threshold
            return best_match

        time.sleep(1)  # Wait a second before checking again

    print(f"Timeout reached while waiting for file: {target_filename}")
    return None


def handle_virus_scan_failure(driver):
    try:
        # Wait for and find the "Download Anyway" button
        download_anyway_script = """
        function findDownloadAnywayButton() {
            // Look for common text patterns in buttons or links
            const elements = document.querySelectorAll('button, a');
            for (const element of elements) {
                const text = element.textContent.toLowerCase();
                if (text.includes('download anyway') || 
                    text.includes('download file') || 
                    text.includes('continue download')) {
                    element.click();
                    return true;
                }
            }
            return false;
        }
        return findDownloadAnywayButton();
        """

        result = driver.execute_script(download_anyway_script)
        if result:
            print("Successfully clicked 'Download Anyway' button")
            time.sleep(2)  # Wait for download to start
            return True
    except Exception as e:
        print(f"Error handling virus scan failure: {str(e)}")
    return False


def wait_for_download_complete(
    download_dir, filename, timeout=300
):  # 5 minutes timeout
    """Wait for file to appear and return its path"""
    print(f"\n⏳ Waiting for download: {filename}")
    start_time = time.time()

    while True:
        # Check if timeout reached
        if time.time() - start_time > timeout:
            print(f"❌ Timeout waiting for download: {filename}")
            return None

        # List all files in download directory
        files = os.listdir(download_dir)

        # Skip any incomplete downloads
        files = [
            f for f in files if not f.endswith(".crdownload") and not f.endswith(".tmp")
        ]

        # Check for matching file
        for file in files:
            if fuzz.ratio(file.lower(), filename.lower()) > 90:
                file_path = os.path.join(download_dir, file)
                print(f"✅ Found downloaded file: {file}")
                return file_path

        # Wait before next check
        time.sleep(2)


def check_if_file_exists(bid_folder, filename):
    """Check if a file with similar name exists in the bid folder"""
    if not os.path.exists(bid_folder):
        return False

    target_filename = filename.lower()
    for existing_file in os.listdir(bid_folder):
        # Compare filenames without numbers in parentheses and extensions
        existing_name = re.sub(
            r"\s*\(\d+\)", "", os.path.splitext(existing_file)[0].lower()
        )
        new_name = re.sub(
            r"\s*\(\d+\)", "", os.path.splitext(target_filename)[0].lower()
        )

        if fuzz.ratio(existing_name, new_name) > 90:  # High similarity threshold
            print(f"Similar file already exists: {existing_file}")
            return True
    return False


def wait_for_file_download_and_move(
    script_download_dir, bid_folder, filename, timeout=300
):
    """Wait for file to download and move it to bid folder"""
    start_time = time.time()
    file_moved = False

    while time.time() - start_time < timeout and not file_moved:
        # Check for incomplete downloads
        downloading_files = [
            f
            for f in os.listdir(script_download_dir)
            if f.endswith(".crdownload") or f.endswith(".tmp")
        ]
        if downloading_files:
            print("Download in progress...")
            time.sleep(2)
            continue

        # Look for matching file
        matching_file = None
        highest_ratio = 0

        for f in os.listdir(script_download_dir):
            if os.path.isfile(os.path.join(script_download_dir, f)):
                ratio = fuzz.ratio(filename.lower(), f.lower())
                if ratio > highest_ratio:
                    highest_ratio = ratio
                    matching_file = f

        if matching_file and highest_ratio > 80:
            try:
                source_path = os.path.join(script_download_dir, matching_file)
                destination_path = os.path.join(bid_folder, matching_file)

                # Ensure file is completely written
                while True:
                    initial_size = os.path.getsize(source_path)
                    time.sleep(1)
                    if os.path.getsize(source_path) == initial_size:
                        break

                safe_move(source_path, destination_path)
                print(f"Successfully moved {matching_file} to bid folder")
                return matching_file

            except Exception as e:
                print(f"Error moving file: {str(e)}")
                time.sleep(1)
        else:
            time.sleep(1)

    print(f"Timeout waiting for file: {filename}")
    return None


def extract_attachments(driver, folder_name, main_dir, script_download_dir):
    """Extract attachments for a bid"""
    try:
        print(f"Extracting attachments for solicitation: {folder_name}")

        # Create bid folder if it doesn't exist
        bid_folder = os.path.join(main_dir, folder_name)
        os.makedirs(bid_folder, exist_ok=True)

        attachment_links = WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.pointer"))
        )

        successful_downloads = []

        for link in attachment_links:
            try:
                original_filename = link.text.strip()
                if not original_filename:
                    print("Skipping attachment with empty filename")
                    continue

                filename = normalize_filename(
                    re.sub(r"^\d+\.\s*", "", original_filename)
                )
                print(f"Attempting to download: {filename}")

                # Check if similar file already exists
                if check_if_file_exists(bid_folder, filename):
                    print(
                        f"Skipping download as similar file already exists for: {filename}"
                    )
                    for existing_file in os.listdir(bid_folder):
                        if (
                            fuzz.ratio(
                                os.path.splitext(existing_file)[0].lower(),
                                os.path.splitext(filename)[0].lower(),
                            )
                            > 90
                        ):
                            successful_downloads.append(existing_file)
                            break
                    continue

                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", link)
                    ActionChains(driver).move_to_element(link).click().perform()

                    downloaded_file = handle_captcha_and_download(
                        driver, script_download_dir, filename, bid_folder
                    )
                    if downloaded_file:
                        successful_downloads.append(downloaded_file)
                        print(f"Successfully downloaded: {downloaded_file}")

                except Exception as e:
                    print(f"Error downloading {filename}: {str(e)}")

                time.sleep(2)

            except Exception as e:
                print(f"Error processing attachment: {str(e)}")
                continue

        print(f"Extracted {len(successful_downloads)} attachments for {folder_name}")
        return ", ".join(successful_downloads)

    except Exception as e:
        print(f"Error extracting attachments for {folder_name}: {e}")
        return ""


def check_for_image_captcha(driver):
    script = """
    function checkForImageCaptcha() {
        function isElementVisible(selector) {
            const element = document.querySelector(selector);
            if (!element) {
                return false;
            }

            const style = window.getComputedStyle(element);
            const isVisible = style.visibility === 'visible' &&
                              parseFloat(style.opacity) > 0 &&
                              element.offsetWidth > 0 &&
                              element.offsetHeight > 0;

            return isVisible;
        }

        const captchaSelector = 'body > div:nth-child(20)';

        if (isElementVisible(captchaSelector)) {
            return true;  // Image captcha detected
        } else {
            return false;  // No image captcha detected
        }
    }

    return checkForImageCaptcha();
    """

    try:
        is_captcha_visible = driver.execute_script(script)

        if is_captcha_visible:
            print("\n🖼️ Image verification captcha detected!")
            print("\n👉 Manual intervention required:")
            print("1. Please complete the image verification")
            print("2. Select all required images")
            print("3. Click verify when complete")
            input("\nPress Enter after completing the image verification... ")
            print("\n✅ Continuing execution...")
            return True
        else:
            print("✅ Regular captcha check completed.")
            return False
    except Exception as e:
        print(f"\n❌ Error checking for image captcha: {str(e)}")
        return False


def wait_for_download_start(driver, timeout=30):
    """Wait and verify that download has started"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # Check for active downloads using Chrome DevTools Protocol
            downloads = driver.execute_script(
                """
                return document.querySelector('downloads-manager')?.shadowRoot
                    ?.querySelector('#downloadsList')
                    ?.items?.filter(e => e.state === 'IN_PROGRESS').length || 0;
            """
            )
            if downloads > 0:
                print("✅ Download started successfully")
                return True
        except:
            pass
        time.sleep(1)

    print("❌ Download did not start within timeout period")
    return False


def handle_captcha_and_download(driver, download_dir, filename, bid_folder):
    try:
        print("\n🔒 Waiting for reCAPTCHA iframe...")
        try:
            iframe = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located(
                    (
                        By.CSS_SELECTOR,
                        "iframe[src^='https://www.google.com/recaptcha/api2/anchor']",
                    )
                )
            )

            print("Captcha detected. Attempting to solve...")
            play_notification_sound()

            captcha_iframe = driver.find_element(
                By.CSS_SELECTOR,
                "iframe[src^='https://www.google.com/recaptcha/api2/anchor']",
            )
            driver.switch_to.frame(captcha_iframe)

            try:
                checkbox = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "recaptcha-anchor"))
                )
                checkbox.click()
                print("✅ Captcha checkbox clicked automatically.")
            except:
                print("\n⚠️ Could not click captcha checkbox automatically.")
                print("\n👉 Please complete the following steps:")
                print("1. Click the reCAPTCHA checkbox")
                print("2. Complete any verification challenges if they appear")
                print(
                    "3. Click the download button that appears after solving the captcha"
                )

            driver.switch_to.default_content()

        except Exception as e:
            print("\n⚠️ Could not find or interact with captcha iframe")
            print("\n👉 Please complete the following steps:")
            print("1. Solve the captcha manually")
            print("2. Complete any verification challenges if they appear")
            print("3. Click the download button that appears after solving the captcha")

        play_notification_sound()
        input("\nAfter clicking the download button, press Enter to continue... ")

        # Wait for download to complete and get file path
        downloaded_file = wait_for_download_complete(download_dir, filename)
        if downloaded_file:
            # Move file to bid folder
            try:
                destination = os.path.join(
                    bid_folder, os.path.basename(downloaded_file)
                )
                safe_move(downloaded_file, destination)
                print(f"✅ Moved file to: {destination}")
                return os.path.basename(downloaded_file)
            except Exception as e:
                print(f"❌ Error moving file: {str(e)}")
                return None

        return None

    except Exception as e:
        print(f"\n❌ Error during captcha handling: {str(e)}")
        return None


def click_download_button(driver):
    """Click the download button using multiple selectors and methods"""
    script = """
    (function() {
        // Array of CSS selectors and XPath selectors for the Download Attachment button
        const selectors = [
            '#captchaVerificationModal > div > div > button',
            '/html/body/div[1]/app-root/app-bid-board/app-bid-board-details/div[3]/div/div/button',
            '#captchaVerificationModal > div > div > button',
            '/html/body/div[1]/app-root/app-bid-board/app-bid-board-details/div[3]/div/div/button'
        ];

        function clickButtonByCSS(selector) {
            const button = document.querySelector(selector);
            if (button) {
                button.click();
                console.log(`Clicked button using CSS selector: ${selector}`);
                return true;
            }
            console.log(`Button not found for CSS selector: ${selector}`);
            return false;
        }

        function clickButtonByXPath(xpath) {
            const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
            const button = result.singleNodeValue;
            if (button) {
                button.click();
                console.log(`Clicked button using XPath selector: ${xpath}`);
                return true;
            }
            console.log(`Button not found for XPath selector: ${xpath}`);
            return false;
        }

        function clickAllDownloadButtons() {
            let clicked = false;

            selectors.forEach(selector => {
                if (selector.startsWith('/')) {
                    const wasClicked = clickButtonByXPath(selector);
                    if (wasClicked) clicked = true;
                } else {
                    const wasClicked = clickButtonByCSS(selector);
                    if (wasClicked) clicked = true;
                }
            });

            return clicked;
        }

        return clickAllDownloadButtons();
    })();
    """

    try:
        result = driver.execute_script(script)
        if result:
            print("✅ Download button clicked successfully")
            return True
        else:
            print("⚠️ Could not find download button with any selector")
            return False
    except Exception as e:
        print(f"❌ Error clicking download button: {str(e)}")
        return False


def save_to_excel(bids, main_dir, append=False):
    """Save bids to Excel file in the main directory"""
    excel_filename = f"{os.path.splitext(os.path.basename(__file__))[0]}.xlsx"
    excel_path = os.path.join(main_dir, excel_filename)

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
        "Additional Summary",
        "Contracting Office Address",
        "Contact Information",
        "Bid Detail Page URL",
        "Attachments",
    ]

    if append and os.path.exists(excel_path):
        existing_df = pd.read_excel(excel_path)
        new_df = pd.DataFrame(bids)
        new_df["SL No"] = range(
            len(existing_df) + 1, len(existing_df) + len(new_df) + 1
        )
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        updated_df = pd.DataFrame(bids)
        updated_df["SL No"] = range(1, len(updated_df) + 1)

    updated_df = updated_df[columns]

    # Convert dates to yyyy-mm-dd format
    for date_column in ["Posted Date", "Response Date"]:
        updated_df[date_column] = pd.to_datetime(updated_df[date_column]).dt.strftime(
            "%Y-%m-%d"
        )

    updated_df.to_excel(excel_path, index=False)
    print(f"Bids saved to {excel_path}")


def move_remaining_files(main_dir, script_download_dir):
    for filename in os.listdir(script_download_dir):
        file_path = os.path.join(script_download_dir, filename)
        if os.path.isfile(file_path):
            moved = False
            for bid_folder in os.listdir(main_dir):
                folder_path = os.path.join(main_dir, bid_folder)
                if os.path.isdir(folder_path):
                    folder_compare = bid_folder.replace("_", "").lower()
                    file_compare = filename.split(".")[0].replace("_", "").lower()

                    if file_compare.startswith(
                        folder_compare
                    ) or folder_compare.startswith(file_compare):
                        try:
                            destination = os.path.join(folder_path, filename)
                            safe_move(file_path, destination)
                            print(f"Moved {filename} to {bid_folder}")
                            moved = True
                            break
                        except Exception as e:
                            print(f"Error moving {filename} to {bid_folder}: {str(e)}")

            if not moved:
                print(f"Could not find a matching folder for {filename}")


def verify_sort_order(driver):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "span.text-small.teal-text.darken-4")
            )
        )
        start_dates = driver.find_elements(
            By.CSS_SELECTOR, "span.text-small.teal-text.darken-4"
        )
        dates = []
        for date_element in start_dates:
            date_text = date_element.text
            if "Start Date:" in date_text:
                date_str = (
                    date_text.split("Start Date:")[1].split("End Date:")[0].strip()
                )
                dates.append(datetime.strptime(date_str, "%b %d, %Y"))

        is_sorted = all(dates[i] >= dates[i + 1] for i in range(len(dates) - 1))

        if is_sorted:
            print("Verification: Bids are correctly sorted by Start Date (Descending).")
        else:
            print(
                "Verification: Bids are NOT correctly sorted by Start Date (Descending)."
            )
            print("Dates found:")
            for i, date in enumerate(dates):
                print(f"  {i+1}. {date.strftime('%Y-%m-%d')}")
    except Exception as e:
        print(f"Error verifying sort order: {str(e)}")


def update_attachments_for_bid(main_dir, solicitation_number):
    # Move up one level to yesterday folder for excel file
    yesterday_dir = os.path.dirname(main_dir)
    excel_filename = "15_State_of_Conneticut_BidBoard.xlsx"
    excel_path = os.path.join(yesterday_dir, excel_filename)

    if not os.path.exists(excel_path):
        print(f"Excel file not found: {excel_path}")
        return

    df = pd.read_excel(excel_path)

    if "Attachments" not in df.columns:
        df["Attachments"] = ""
    else:
        df["Attachments"] = df["Attachments"].fillna("").astype(str)

    folder_name = solicitation_number.replace("/", "_")
    # Bid folders should be directly in yesterday folder
    bid_folder = os.path.join(yesterday_dir, folder_name)

    if os.path.exists(bid_folder):
        attachments = [
            f
            for f in os.listdir(bid_folder)
            if os.path.isfile(os.path.join(bid_folder, f))
        ]
        attachments_str = ", ".join(attachments) if attachments else ""

        df.loc[df["Solicitation Number"] == solicitation_number, "Attachments"] = (
            attachments_str
        )

        df.to_excel(excel_path, index=False)
        print(f"Attachments updated for bid {solicitation_number} in {excel_path}")
    else:
        print(f"Bid folder not found for {solicitation_number}")


def move_remaining_files_for_bid(main_dir, script_download_dir, solicitation_number):
    """Move downloaded files to bid-specific folders in the main directory"""
    folder_name = solicitation_number.replace("/", "_")
    bid_folder = os.path.join(main_dir, folder_name)

    if not os.path.exists(bid_folder):
        os.makedirs(bid_folder)

    for filename in os.listdir(script_download_dir):
        file_path = os.path.join(script_download_dir, filename)
        if os.path.isfile(file_path):
            normalized_filename = normalize_filename(filename)
            file_compare = normalized_filename.split(".")[0].replace("_", "").lower()
            folder_compare = folder_name.replace("_", "").lower()

            if file_compare.startswith(folder_compare) or folder_compare.startswith(
                file_compare
            ):
                try:
                    destination = os.path.join(bid_folder, normalized_filename)
                    safe_move(file_path, destination)
                    print(f"Moved {filename} to {folder_name}")
                except Exception as e:
                    print(f"Error moving {filename} to {folder_name}: {str(e)}")

    # Delete empty folders
    if not os.listdir(bid_folder):
        os.rmdir(bid_folder)
        print(f"Deleted empty folder: {bid_folder}")


def update_attachments_in_excel(main_dir, folder_name, attachments):
    try:
        excel_filename = "15_State_of_Conneticut_BidBoard.xlsx"
        excel_path = os.path.join(main_dir, excel_filename)

        if not os.path.exists(excel_path):
            print(f"Excel file not found: {excel_path}")
            return

        df = pd.read_excel(excel_path)

        # Convert folder_name back to solicitation number format
        solicitation_number = folder_name.replace("_", "/")

        # Update the Attachments column for this solicitation number
        attachments_str = ", ".join(attachments) if attachments else ""
        mask = df["Solicitation Number"] == solicitation_number
        if mask.any():
            df.loc[mask, "Attachments"] = attachments_str
            df.to_excel(excel_path, index=False)
            print(
                f"Updated attachments in Excel for {solicitation_number}: {attachments_str}"
            )
        else:
            print(
                f"Could not find solicitation number {solicitation_number} in Excel file"
            )

    except Exception as e:
        print(f"Error updating Excel file with attachments: {str(e)}")


def move_remaining_attachments(main_dir, script_download_dir):
    """Move any remaining files in the downloads folder to appropriate bid folders"""
    print("\nChecking for remaining attachments...")

    if not os.path.exists(script_download_dir):
        return

    # Get all bid folders
    bid_folders = [
        d
        for d in os.listdir(main_dir)
        if os.path.isdir(os.path.join(main_dir, d))
        and d != os.path.basename(script_download_dir)
    ]

    remaining_files = [
        f
        for f in os.listdir(script_download_dir)
        if os.path.isfile(os.path.join(script_download_dir, f))
        and not f.endswith(".crdownload")
    ]

    if not remaining_files:
        print("No remaining files found.")
        return

    print(f"Found {len(remaining_files)} remaining files.")

    for filename in remaining_files:
        best_match = None
        highest_ratio = 0

        # Clean filename for comparison
        clean_filename = re.sub(
            r"\s*\(\d+\)", "", os.path.splitext(filename)[0].lower().replace("_", "")
        )

        # Try to find matching bid folder
        for folder in bid_folders:
            clean_folder = folder.lower().replace("_", "")

            # Check if folder name is in filename or vice versa
            if clean_folder in clean_filename or clean_filename in clean_folder:
                ratio = fuzz.ratio(clean_folder, clean_filename)
                if ratio > highest_ratio:
                    highest_ratio = ratio
                    best_match = folder

            # Also check existing files in the folder for similarity
            folder_path = os.path.join(main_dir, folder)
            for existing_file in os.listdir(folder_path):
                clean_existing = re.sub(
                    r"\s*\(\d+\)",
                    "",
                    os.path.splitext(existing_file)[0].lower().replace("_", ""),
                )
                ratio = fuzz.ratio(clean_existing, clean_filename)
                if ratio > highest_ratio:
                    highest_ratio = ratio
                    best_match = folder

        if best_match and highest_ratio > 60:  # Adjust threshold as needed
            try:
                source = os.path.join(script_download_dir, filename)
                destination = os.path.join(main_dir, best_match, filename)
                safe_move(source, destination)
                print(f"Moved remaining file {filename} to {best_match}")

                # Update Excel with the new attachment
                excel_path = os.path.join(
                    main_dir, "15_State_of_Conneticut_BidBoard.xlsx"
                )
                if os.path.exists(excel_path):
                    df = pd.read_excel(excel_path)
                    solicitation_number = best_match.replace("_", "/")
                    mask = df["Solicitation Number"] == solicitation_number
                    if mask.any():
                        current_attachments = str(df.loc[mask, "Attachments"].iloc[0])
                        if current_attachments and current_attachments != "nan":
                            new_attachments = f"{current_attachments}, {filename}"
                        else:
                            new_attachments = filename
                        df.loc[mask, "Attachments"] = new_attachments
                        df.to_excel(excel_path, index=False)
                        print(f"Updated Excel attachments for {solicitation_number}")
            except Exception as e:
                print(f"Error moving {filename}: {str(e)}")
        else:
            print(f"Could not find matching folder for {filename}")


def check_for_error_message(driver):
    """Check if the error message is present on the page"""
    error_selectors = [
        "/html/body/div[1]/app-root/app-bid-board/app-bid-board-result/div[1]/h5/span/span/b",
        "#webprocure_public_contract_board > app-root > app-bid-board > app-bid-board-result > div.search-results > h5 > span > span",
        "//*[@id='webprocure_public_contract_board']/app-root/app-bid-board/app-bid-board-result/div[1]/h5/span/span",
        "/html/body/div[1]/app-root/app-bid-board/app-bid-board-result/div[1]/h5/span/span",
    ]

    for selector in error_selectors:
        try:
            if selector.startswith("/"):
                element = driver.find_element(By.XPATH, selector)
            else:
                element = driver.find_element(By.CSS_SELECTOR, selector)

            if (
                "Well this was unexpected" in element.text
                or "Something went wrong" in element.text
            ):
                return True
        except:
            continue
    return False


def handle_page_load(driver, url, max_attempts=3):
    """Handle page load with retries and longer waits"""
    attempts = 0
    while attempts < max_attempts:
        try:
            driver.get(url)
            # Increased initial wait time to 10 seconds
            time.sleep(10)

            if check_for_error_message(driver):
                print("\n⚠️ Error: Page didn't load properly!")
                print("Message: 'Well this was unexpected. Something went wrong.'")
                attempts += 1
                if attempts < max_attempts:
                    print(f"\nRetrying... (Attempt {attempts + 1} of {max_attempts})")
                    # Increased wait between attempts to 5 seconds
                    time.sleep(5)
                    continue
                else:
                    print("\n❌ All attempts failed. Skipping to next URL...")
                    return False

            # Additional wait after successful load
            time.sleep(5)
            return True

        except Exception as e:
            print(f"\nError loading page: {str(e)}")
            attempts += 1
            if attempts < max_attempts:
                print(f"Retrying... (Attempt {attempts + 1} of {max_attempts})")
                time.sleep(5)
            else:
                print("\n❌ All attempts failed. Skipping to next URL...")
                return False

    return False


# Add this function to check for "No search results" message
def check_no_results(driver):
    """Check if the page shows 'No search results available'"""
    try:
        script = """
        function checkNoResults() {
            const selectors = [
                "#webprocure_public_contract_board > app-root > app-bid-board > app-bid-board-result > div:nth-child(1) > h5 > span > span",
                "//*[@id='webprocure_public_contract_board']/app-root/app-bid-board/app-bid-board-result/div[1]/h5/span/span",
                "/html/body/div[1]/app-root/app-bid-board/app-bid-board-result/div[1]/h5/span/span"
            ];
            
            for (const selector of selectors) {
                let element;
                try {
                    if (selector.startsWith('/')) {
                        // XPath
                        const result = document.evaluate(selector, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                        element = result.singleNodeValue;
                    } else {
                        // CSS selector
                        element = document.querySelector(selector);
                    }
                    
                    if (element && element.textContent.trim().includes('No search results available')) {
                        return true;
                    }
                } catch (e) {
                    console.log('Error checking selector:', e);
                }
            }
            return false;
        }
        return checkNoResults();
        """
        return driver.execute_script(script)
    except Exception as e:
        print(f"Error checking for no results: {e}")
        return False


def main():
    args = parse_arguments()
    main_dir, script_download_dir = create_directory_structure()
    driver = init_driver(script_download_dir)
    all_bids = []

    try:
        # Randomize URL order
        shuffled_urls = CUSTOMER_URLS.copy()
        random.shuffle(shuffled_urls)

        for index, url in enumerate(shuffled_urls, 1):
            print(f"\n📌 Processing URL {index}/{len(CUSTOMER_URLS)}")
            print(f"🔗 URL: {url}")

            # Add random delay between URLs (5-15 seconds)
            delay = random.uniform(5, 15)
            print(f"Waiting {delay:.2f} seconds before processing next URL...")
            time.sleep(delay)

            if start_process(driver, url):
                try:
                    # Add random delay before applying filters (2-5 seconds)
                    time.sleep(random.uniform(2, 5))
                    apply_filters(driver)

                    # Add random delay before verifying sort (1-3 seconds)
                    time.sleep(random.uniform(1, 3))
                    verify_sort_order(driver)

                    bids = extract_bids(
                        driver, args.days, main_dir, script_download_dir
                    )
                    if bids:
                        all_bids.extend(bids)
                        print(
                            f"✅ Successfully extracted {len(bids)} bids from URL {index}"
                        )
                    else:
                        print(f"ℹ️ No bids found for URL {index}")
                except Exception as e:
                    print(f"❌ Error processing URL {index}: {str(e)}")
                    continue

            # Add longer random delay between URLs (10-20 seconds)
            delay = random.uniform(10, 20)
            print(f"Waiting {delay:.2f} seconds before next URL...")
            time.sleep(delay)

        if all_bids:
            save_to_excel(all_bids, main_dir)
            print(f"📂 Total bids extracted: {len(all_bids)}")
        else:
            print("ℹ️ No bids found from any URL")

    except Exception as e:
        print(f"❌ An error occurred in main execution: {str(e)}")
        import traceback

        traceback.print_exc()
    finally:
        try:
            driver.quit()
        except:
            pass

        # Clean up download directory
        try:
            if os.path.exists(script_download_dir):
                shutil.rmtree(script_download_dir)
                print("\n🗑️ Cleaned up temporary download directory")
        except Exception as e:
            print(f"\n⚠️ Error cleaning up download directory: {str(e)}")

        # Final update of attachments in Excel file
        try:
            excel_filename = f"{os.path.splitext(os.path.basename(__file__))[0]}.xlsx"
            excel_path = os.path.join(main_dir, excel_filename)
            if os.path.exists(excel_path):
                df = pd.read_excel(excel_path)
                for _, row in df.iterrows():
                    update_attachments_for_bid(main_dir, row["Solicitation Number"])
                print("\n✅ Updated attachments in Excel file")
        except Exception as e:
            print(f"\n❌ Error updating Excel attachments: {str(e)}")

        # Rename folder to indicate completion
        try:
            if os.path.exists(main_dir):
                yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
                script_name = os.path.splitext(os.path.basename(__file__))[0]
                completed_dir = os.path.join(
                    os.getcwd(), yesterday, f"{script_name}_COMPLETED"
                )

                if os.path.exists(completed_dir):
                    shutil.rmtree(completed_dir)

                os.rename(main_dir, completed_dir)
                print(f"\n✅ Renamed folder to indicate completion: {completed_dir}")
        except Exception as e:
            print(f"\n❌ Error renaming folder: {str(e)}")

    print(
        "🏁 Process completed. Attachments have been updated in the Excel file for each bid."
    )


if __name__ == "__main__":
    main()
