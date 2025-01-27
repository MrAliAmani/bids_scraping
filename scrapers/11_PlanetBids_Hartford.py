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
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
    StaleElementReferenceException,
)
from dotenv import load_dotenv
import threading
import requests
import logging
from dateutil import parser
import pytz
import argparse
import traceback
import zipfile
import glob
import shutil
from selenium.webdriver.common.keys import Keys
import winsound  # For playing notification sounds
from fake_useragent import UserAgent
from selenium_stealth import stealth
import random
import pickle
from utils.utils import safe_move, play_notification_sound

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def setup_driver():
    """Set up and return a configured Chrome WebDriver with anti-bot measures."""
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--disable-gpu")  # Disables GPU acceleration
    options.add_argument("--disable-software-rasterizer")  # Use CPU rendering
    options.add_argument("--log-level=3")  # Only show fatal errors

    # Rotate user agents
    ua = UserAgent()
    options.add_argument(f"user-agent={ua.random}")

    # Handle headless browser detection
    options.add_argument("--disable-blink-features=AutomationControlled")

    # Add these options to make the browser more like a real user
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(options=options)

    # Use selenium-stealth to make detection harder
    stealth(
        driver,
        languages=["en-US", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
    )

    return driver, options


def wait_and_click(driver, xpath, timeout=60):
    """Wait for an element to be clickable and then click it."""
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    element.click()


def random_delay(min_delay=1, max_delay=5):
    """Add a random delay to mimic human behavior."""
    time.sleep(random.uniform(min_delay, max_delay))


def mimic_human_interaction(driver):
    """Mimic human interaction with random scrolling."""
    try:
        # Scroll a random amount
        driver.execute_script(f"window.scrollTo(0, {random.randint(100, 500)});")
    except Exception as e:
        logging.warning(f"Error in mimic_human_interaction: {str(e)}")
        # If an error occurs, we'll just log it and continue without raising an exception


def login(driver):
    """Log in to the PlanetBids website and navigate to the search page."""
    try:
        # Navigate to the login page
        driver.get("https://vendors.planetbids.com/portal/22554/login")
        logging.info("Navigated to login page")
        random_delay()

        # Perform login
        email_input = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, "username-field"))
        )
        mimic_human_interaction(driver)
        email_input.clear()
        for char in os.getenv("HARTFORD_EMAIL"):
            email_input.send_keys(char)
            random_delay(0.1, 0.3)
        logging.info("Entered email")
        random_delay(0.5, 2)

        password_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#password-field"))
        )
        mimic_human_interaction(driver)
        password_input.clear()
        for char in os.getenv("HARTFORD_PASSWORD"):
            password_input.send_keys(char)
            random_delay(0.1, 0.3)
        logging.info("Entered password")
        random_delay(0.5, 2)

        submit_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//button[contains(@class, 'success-block-btn') and contains(text(), 'LOG IN')]",
                )
            )
        )
        mimic_human_interaction(driver)
        submit_button.click()
        logging.info("Clicked submit button")

        # Wait for login to complete
        WebDriverWait(driver, 30).until(
            lambda d: d.current_url
            != "https://vendors.planetbids.com/portal/22554/login"
        )

        # Attempt to navigate to the search page
        max_attempts = 5
        for attempt in range(max_attempts):
            logging.info(f"Attempt {attempt + 1} to navigate to search page")
            driver.get("https://vendors.planetbids.com/portal/22554/bo/bo-search")
            random_delay(3, 5)

            # Check if we're on the search page
            if "bo/bo-search" in driver.current_url:
                logging.info("Successfully navigated to search page")
                break
            elif "portal-home" in driver.current_url:
                logging.info("Redirected to portal home. Attempting to log in again.")
                # Attempt to log in again
                login_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//a[contains(text(), 'Log In')]")
                    )
                )
                login_button.click()
                random_delay(2, 4)
                # Re-enter credentials
                email_input = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "username-field"))
                )
                email_input.clear()
                email_input.send_keys(os.getenv("HARTFORD_EMAIL"))
                password_input = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "#password-field")
                    )
                )
                password_input.clear()
                password_input.send_keys(os.getenv("HARTFORD_PASSWORD"))
                submit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//button[contains(@class, 'success-block-btn') and contains(text(), 'LOG IN')]",
                        )
                    )
                )
                submit_button.click()
                random_delay(3, 5)
            else:
                logging.warning(
                    f"Unexpected URL after login attempt: {driver.current_url}"
                )

        # Final check
        if "bo/bo-search" not in driver.current_url:
            raise Exception("Failed to navigate to search page after multiple attempts")

        # Wait for an element that should be present on the search page
        WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, "stageId-field"))
        )

        logging.info("🟢 Login Successful and on correct page")

        # Persist cookies
        pickle.dump(driver.get_cookies(), open("cookies.pkl", "wb"))

    except TimeoutException:
        logging.error("Login failed: Timeout while waiting for an element")
        raise
    except NoSuchElementException as e:
        logging.error(f"Login failed: Element not found - {str(e)}")
        raise
    except Exception as e:
        logging.error(f"Login failed: Unexpected error - {str(e)}")
        logging.error(f"Traceback: {traceback.format_exc()}")
        raise


def set_filters(driver):
    """Set the required filters for bid search."""
    try:
        logging.info("Waiting for the page to load completely...")
        # Wait for the page to load completely
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "stageId-field"))
        )
        logging.info("Page loaded, stageId-field found")

        # Wait a bit more to ensure all elements are interactive
        time.sleep(5)

        logging.info("Setting stage as 'Bidding'...")
        # Set stage as "Bidding"
        try:
            stage_dropdown = Select(driver.find_element(By.ID, "stageId-field"))
            stage_dropdown.select_by_value("3")
            logging.info("Stage set to 'Bidding'")
        except Exception as e:
            logging.error(f"Error setting stage: {str(e)}")
            raise

        # Wait for the dropdown to update
        time.sleep(2)

        logging.info("Locating search button...")
        try:
            search_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//button[@class='search-btn' and contains(text(), 'Search')]",
                    )
                )
            )
            logging.info("Search button found")
        except Exception as e:
            logging.error(f"Error locating search button: {str(e)}")
            raise

        logging.info("Clicking search button...")
        try:
            search_button.click()
            logging.info("Search button clicked")
        except Exception as e:
            logging.error(f"Error clicking search button: {str(e)}")
            raise

        # Wait for a short time after clicking search
        time.sleep(5)

        logging.info("Waiting for results table...")
        try:
            # Wait for the results table to be present using a more specific XPath
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//div[contains(@class, 'table-overflow-container')]/table[contains(@class, 'pb-datatable data')]",
                    )
                )
            )
            logging.info("Results table found")
        except Exception as e:
            logging.error(f"Error waiting for results table: {str(e)}")
            # Log the page source if the table is not found
            logging.debug(f"Current page source:\n{driver.page_source}")
            raise

        logging.info("Filters set and search completed")
    except Exception as e:
        logging.error(f"Error setting filters: {str(e)}")
        logging.debug(f"Current page source:\n{driver.page_source}")
        raise


def extract_bid_details(driver, posted_date):
    """
    Extract bid details from the current page using JavaScript.

    Args:
        driver (WebDriver): Selenium WebDriver instance
        posted_date (str): The posted date extracted from the search page

    Returns:
        dict: Dictionary containing extracted bid details
    """
    try:
        # Check if we're on the PB page and navigate to the bid detail page if necessary
        if "#bidPBDetail" in driver.current_url:
            bid_detail_url = driver.current_url.split("#")[0]
            driver.get(bid_detail_url)
            logging.info(f"Navigated from PB page to bid detail page: {bid_detail_url}")
            # Wait for the page to load
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, "bid-detail-item-title"))
            )

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "bid-detail-item-title"))
        )

        details = {
            "Posted Date": posted_date,
            "Response Date": safe_get_text(
                driver, "/html/body/div[3]/div/div/div[2]/div[6]/div/div[2]"
            ),
            "Notice Type": safe_get_text(
                driver, "/html/body/div[3]/div/div/div[2]/div[8]/div/div[2]"
            ),
            "Solicitation Number": get_solicitation_number(driver),
            "Solicitation Title": safe_get_text(
                driver, "/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]"
            ),
            "Agency": safe_get_text(
                driver, "/html/body/div[3]/div/div/div[2]/div[15]/div/div[2]"
            ),
            "Category": safe_get_text(
                driver, "/html/body/div[3]/div/div/div[2]/div[11]/div[2]"
            ),
            "Contracting Office Address": safe_get_text(
                driver, "/html/body/div[3]/div/div/div[2]/div[14]/div[2]/div[1]"
            ),
            "Bid Detail Page URL": driver.current_url,
        }

        # Helper function to clean JavaScript output
        def clean_js_output(output):
            return "" if output in ["<!---->", "None", "null"] else output

        # Extract Contact Information
        contact_info_script = """
        const titles = document.querySelectorAll('.bid-detail-item-title');
        const contactInfos = [];
        titles.forEach(title => {
            if (title.textContent.trim() === 'Contact Info') {
                const valueDiv = title.nextElementSibling;
                if (valueDiv && valueDiv.classList.contains('bid-detail-item-value')) {
                    const contactInfo = valueDiv.innerHTML.replace(/<br\\s*\\/?>/gi, '\\n').trim();
                    contactInfos.push(contactInfo);
                }
            }
        });
        return contactInfos[0] || '';
        """
        details["Contact Information"] = clean_js_output(
            driver.execute_script(contact_info_script)
        )

        # Extract Description (Scope of Services)
        description_script = """
        const titles = document.querySelectorAll('.bid-detail-item-title');
        const scopeOfServicesList = [];
        titles.forEach(title => {
            if (title.textContent.trim() === 'Scope of Services') {
                const valueDiv = title.nextElementSibling;
                if (valueDiv && valueDiv.classList.contains('bid-detail-item-value')) {
                    const scopeOfServices = valueDiv.innerHTML.replace(/<br\\s*\\/?>/gi, '\\n').trim();
                    scopeOfServicesList.push(scopeOfServices);
                }
            }
        });
        return scopeOfServicesList[0] || '';
        """
        details["Description"] = clean_js_output(
            driver.execute_script(description_script)
        )

        # Extract Additional Summary (Other Details)
        additional_summary_script = """
        const titles = document.querySelectorAll('.bid-detail-item-title');
        let otherDetails = '';
        titles.forEach(title => {
            if (title.textContent.trim() === 'Other Details') {
                const valueDiv = title.nextElementSibling;
                if (valueDiv && valueDiv.classList.contains('bid-detail-item-value')) {
                    otherDetails = valueDiv.innerHTML.replace(/<br\\s*\\/?>/gi, '\\n').trim();
                }
            }
        });
        return otherDetails;
        """
        details["Additional Summary"] = clean_js_output(
            driver.execute_script(additional_summary_script)
        )

        # Parse Response Date
        try:
            date_str = details["Response Date"]
            if date_str != "N/A":
                date_part = date_str.split()[0]
                parsed_date = datetime.strptime(date_part, "%m/%d/%Y")
                details["Response Date"] = parsed_date.strftime("%Y-%m-%d")
            else:
                details["Response Date"] = "Unknown"
        except Exception as e:
            logging.warning(
                f"Could not parse date for Response Date: {date_str}. Error: {str(e)}"
            )
            details["Response Date"] = "Unknown"

        logging.info(f"Extracted bid details: {details}")
        return details

    except Exception as e:
        logging.error(f"Error extracting bid details: {str(e)}")
        logging.error(f"Traceback: {traceback.format_exc()}")
        return None


def get_solicitation_number(driver):
    """
    Extract the Solicitation Number from the bid detail page.
    """
    methods = [
        lambda: driver.find_element(
            By.XPATH, "/html/body/div[5]/div/div/div[1]/div/div[2]/div[1]/h3/span"
        ).text.strip(),
        lambda: driver.find_element(
            By.CSS_SELECTOR,
            "#ember204 > div > div.fixed-detail-heading-wrapper > div > div.title-and-countdown-display.fixed-heading-row > div:nth-child(1) > h3 > span",
        ).text.strip(),
        lambda: driver.execute_script(
            "return document.querySelector('#ember204 > div > div.fixed-detail-heading-wrapper > div > div.title-and-countdown-display.fixed-heading-row > div:nth-child(1) > h3 > span').textContent.trim()"
        ),
        lambda: driver.find_element(
            By.XPATH, "//h3[contains(@class, 'bid-detail-title')]/span"
        ).text.strip(),
        lambda: driver.find_element(
            By.XPATH, "/html/body/div[5]/div/div/div[2]/div[3]/div/div[2]"
        ).text.strip(),
    ]

    for method in methods:
        try:
            result = method()
            if result and result != "N/A":
                return result
        except Exception:
            continue  # Silently continue to the next method if this one fails

    logging.error("Failed to extract Solicitation Number using all available methods")
    return ""


def safe_get_text(driver, xpath, default=""):
    """Safely get text from an element, returning a default value if not found."""
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        text = element.text.strip()
        logging.info(f"Successfully extracted text for XPath {xpath}: {text}")
        return text
    except Exception as e:
        logging.warning(f"Could not find element with XPath: {xpath}. Error: {str(e)}")
        return default


def download_attachments(driver, options, bid_number):
    """Download attachments for a single bid."""
    try:
        # Wait for the page to load completely
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        # Use JavaScript to click on the Documents tab
        js_click_documents = """
        const documentsLink = document.querySelector('#detail-navigation > ul > li.bidDocs > a');
        if (documentsLink) {
            documentsLink.click();
            return true;
        } else {
            return false;
        }
        """
        documents_clicked = driver.execute_script(js_click_documents)
        if not documents_clicked:
            logging.error(
                "Failed to find the 'Documents' link with the provided selector."
            )
            return ""

        # Get the script folder path
        yesterday = datetime.now() - timedelta(days=1)
        yesterday_folder = yesterday.strftime("%Y-%m-%d")
        script_name = "11_PlanetBids_Hartford"
        script_folder = f"{script_name}_IN_PROGRESS"
        script_folder_path = os.path.join(yesterday_folder, script_folder)

        # Only create bid folder when we confirm there are attachments to download
        bid_folder = os.path.join(script_folder_path, bid_number)

        # Find the "Download All" button using its text
        try:
            download_all_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(text(), 'Download All')]")
                )
            )
        except TimeoutException:
            logging.error("'Download All' button not found.")
            return ""

        # Create the bid folder only after confirming there are attachments
        os.makedirs(bid_folder, exist_ok=True)

        # Set Chrome options to download to the bid-specific folder
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": bid_folder},
        )

        # Scroll the button into view
        driver.execute_script("arguments[0].scrollIntoView(true);", download_all_button)
        time.sleep(2)  # Wait for scroll to complete

        # Use ActionChains to click the download button
        ActionChains(driver).move_to_element(download_all_button).click().perform()
        logging.info("Clicked 'Download All' button.")
        time.sleep(5)  # Wait for potential "Become a PB" button

        # Check for "Become a PB" button and handle it
        try:
            become_pb_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(text(), 'Become a PB')]")
                )
            )
            become_pb_button.click()
            logging.info("'Become a PB' button clicked.")

            # Wait for the classification dropdown to be present in the DOM
            classification_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "classification"))
            )

            # Scroll the classification dropdown into view
            driver.execute_script(
                "arguments[0].scrollIntoView(true);", classification_dropdown
            )
            time.sleep(2)  # Wait for scroll to complete

            # Select "Other" classification using JavaScript
            js_select_classification = """
            let selectElement = document.querySelector('select[name="classification"]');
            if (selectElement) {
                selectElement.value = "4";
                selectElement.dispatchEvent(new Event('change'));
                return true;
            } else {
                return false;
            }
            """
            classification_selected = driver.execute_script(js_select_classification)
            if classification_selected:
                logging.info("Classification set to 'Other'.")
            else:
                logging.warning("Classification dropdown not found.")

            # Set pre-bid meeting to "No" if present
            try:
                pre_bid_meeting = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "pre-bid-mtg"))
                )
                if pre_bid_meeting.is_enabled():
                    pre_bid_meeting.clear()
                    pre_bid_meeting.send_keys("No")
                    logging.info("Pre-bid meeting set to 'No'.")
                else:
                    logging.info("Pre-bid meeting field is not editable.")
            except:
                logging.info("Pre-bid meeting field not found.")

            # Click "Done" button
            done_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "button.default-btn[type='submit']")
                )
            )
            done_button.click()
            logging.info("'Done' button clicked.")

        except TimeoutException:
            logging.info("'Become a PB' button not found, continuing with download")

            # Check if we're on the PB page with no attachments
            if "#bidPBDetail" in driver.current_url:
                logging.info(
                    "Redirected to PB page without attachments. No attachments available for this bid."
                )
                return (
                    ""  # Return an empty string instead of "No attachments available"
                )

        # Wait for download to complete
        max_wait_time = 300  # 5 minutes
        check_interval = 60  # 60 seconds

        # Check for downloaded files directly in bid folder
        for _ in range(max_wait_time // check_interval):
            time.sleep(check_interval)
            zip_files = [f for f in os.listdir(bid_folder) if f.endswith(".zip")]

            if zip_files:
                latest_zip = max(
                    [os.path.join(bid_folder, f) for f in zip_files],
                    key=os.path.getctime,
                )

                # Unzip directly to bid folder
                with zipfile.ZipFile(latest_zip, "r") as zip_ref:
                    total_files = len(zip_ref.namelist())
                    for index, file in enumerate(zip_ref.namelist(), start=1):
                        logging.info(
                            f"Extracting file {index} of {total_files}: {file}"
                        )
                        zip_ref.extract(file, bid_folder)

                # Remove the zip file
                os.remove(latest_zip)

                # Check if any files were downloaded and remove empty folder if none were
                if not os.path.exists(bid_folder) or not os.listdir(bid_folder):
                    if os.path.exists(bid_folder):
                        os.rmdir(bid_folder)
                    return ""

                # Get all extracted filenames
                attachment_files = [
                    f
                    for f in os.listdir(bid_folder)
                    if os.path.isfile(os.path.join(bid_folder, f))
                ]
                return ", ".join(attachment_files)

        logging.warning(
            f"No zip file found for bid {bid_number} after {max_wait_time} seconds"
        )
        return ""

    except Exception as e:
        logging.error(f"Error downloading attachments for bid {bid_number}: {str(e)}")
        logging.error(f"Traceback: {traceback.format_exc()}")
        # Clean up empty folder if it exists
        if os.path.exists(bid_folder) and not os.listdir(bid_folder):
            os.rmdir(bid_folder)
        return ""


def process_bids(driver, options, days_to_scrape):
    bids_data = []
    current_date = datetime.now(pytz.timezone("US/Eastern")).date()
    cutoff_date = current_date - timedelta(days=days_to_scrape)
    bid_links = []

    # Create folder structure
    yesterday = datetime.now() - timedelta(days=1)
    yesterday_folder = yesterday.strftime("%Y-%m-%d")
    script_name = "11_PlanetBids_Hartford"

    # Create the main script folder with IN_PROGRESS suffix
    script_folder = f"{script_name}_IN_PROGRESS"
    script_folder_path = os.path.join(yesterday_folder, script_folder)
    os.makedirs(script_folder_path, exist_ok=True)

    # Create temporary download folder inside script folder
    temp_download_folder = os.path.join(script_folder_path, script_name)
    os.makedirs(temp_download_folder, exist_ok=True)

    try:
        while True:
            logging.info("Waiting for bid rows to be present...")
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//table[contains(@class, 'pb-datatable data')]//tbody/tr",
                    )
                )
            )

            bid_rows = driver.find_elements(
                By.XPATH, "//table[contains(@class, 'pb-datatable data')]//tbody/tr"
            )

            logging.info(f"Found {len(bid_rows)} bid rows on the current page.")

            for row in bid_rows:
                try:
                    date_element = row.find_element(By.XPATH, "./td[1]")
                    displayed_date_str = date_element.text.strip()

                    solicitation_number_element = row.find_element(By.XPATH, "./td[3]")
                    solicitation_number = solicitation_number_element.text.strip()

                    row_attribute = row.get_attribute("rowattribute")

                    if row_attribute:
                        bid_detail_url = f"https://vendors.planetbids.com/portal/22554/bo/bo-detail/{row_attribute}"

                        try:
                            displayed_date = datetime.strptime(
                                displayed_date_str, "%m/%d/%Y"
                            ).date()
                            ymd_date_str = displayed_date.strftime("%Y-%m-%d")
                        except ValueError:
                            logging.warning(
                                f"Could not parse date: {displayed_date_str}"
                            )
                            continue

                        if cutoff_date <= displayed_date <= current_date:
                            bid_links.append(
                                {
                                    "url": bid_detail_url,
                                    "posted_date": ymd_date_str,
                                    "solicitation_number": solicitation_number,
                                }
                            )
                            logging.info(
                                f"Added bid link: {bid_detail_url}, Posted Date: {ymd_date_str}, Solicitation Number: {solicitation_number}"
                            )
                    else:
                        logging.warning(
                            f"Could not find rowattribute for bid on {displayed_date_str}"
                        )

                except StaleElementReferenceException:
                    logging.warning(
                        "Stale element encountered. Refreshing page and retrying."
                    )
                    driver.refresh()
                    time.sleep(5)
                    break
                except Exception as e:
                    logging.error(f"Error processing bid row: {str(e)}")
                    logging.error(f"Traceback: {traceback.format_exc()}")

            # Check if there's a next page
            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//a[contains(@class, 'page-link') and contains(text(), 'Next')]",
                        )
                    )
                )
                next_button.click()
                time.sleep(5)  # Wait for next page to load
            except TimeoutException:
                logging.info("No more pages to process.")
                break
            except Exception as e:
                logging.error(f"Error navigating to next page: {str(e)}")
                break

        # Process collected bid links
        total_bids = len(bid_links)
        for index, bid_link in enumerate(bid_links, start=1):
            try:
                logging.info(f"Processing bid {index} out of {total_bids}")

                driver.get(bid_link["url"])
                time.sleep(5)

                bid_details = extract_bid_details(driver, bid_link["posted_date"])
                if bid_details:
                    bid_details["Posted Date"] = bid_link["posted_date"]
                    bid_details["Solicitation Number"] = bid_link["solicitation_number"]

                    # Download attachments
                    attachments = download_attachments(
                        driver, options, bid_link["solicitation_number"]
                    )
                    bid_details["Attachments"] = attachments
                    bids_data.append(bid_details)

                    # Save Excel file inside script folder
                    save_to_excel(bids_data, days_to_scrape, script_folder_path)

                    # Clean up empty bid folder if it exists
                    bid_folder = os.path.join(
                        script_folder_path, bid_link["solicitation_number"]
                    )
                    if os.path.exists(bid_folder) and not os.listdir(bid_folder):
                        os.rmdir(bid_folder)

                else:
                    logging.error("Failed to extract bid details")
                    play_notification_sound()
                    user_input = input(
                        "An error occurred. Press Enter to continue or type 'q' to quit: "
                    )
                    if user_input.lower() == "q":
                        break

            except Exception as e:
                logging.error(f"Error processing bid {index}: {str(e)}")
                play_notification_sound()
                user_input = input(
                    "An error occurred. Press Enter to continue or type 'q' to quit: "
                )
                if user_input.lower() == "q":
                    break

                # Clean up empty bid folder if it exists after an error
                bid_folder = os.path.join(
                    script_folder_path, bid_link["solicitation_number"]
                )
                if os.path.exists(bid_folder) and not os.listdir(bid_folder):
                    os.rmdir(bid_folder)

    except Exception as e:
        logging.error(f"Error processing bids: {str(e)}")
        logging.error(f"Traceback: {traceback.format_exc()}")
        play_notification_sound()
        input("Press Enter to continue...")

    # Clean up temporary download folder if empty
    if os.path.exists(temp_download_folder) and not os.listdir(temp_download_folder):
        os.rmdir(temp_download_folder)

    return bids_data


def save_to_excel(bids_data, days_to_scrape, script_folder_path):
    """Save extracted bid data to an Excel file."""
    df = pd.DataFrame(bids_data)

    # Reorder columns as per the required order
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

    # Add SL No column
    df.insert(0, "SL No", range(1, len(df) + 1))

    # Ensure all required columns are present, add empty ones if missing
    for col in column_order:
        if col not in df.columns:
            df[col] = ""

    # Reorder the DataFrame
    df = df[column_order]

    # Update Attachments column with paths from bid-specific folders
    for index, row in df.iterrows():
        solicitation_number = row["Solicitation Number"]
        folder_path = os.path.join(script_folder_path, solicitation_number)
        if os.path.exists(folder_path):
            attachment_files = [
                f
                for f in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, f))
            ]
            df.at[index, "Attachments"] = ", ".join(attachment_files).rstrip(",")

    # Save Excel file in script folder
    filename = os.path.join(
        script_folder_path,
        f"{os.path.basename(script_folder_path).replace('_IN_PROGRESS', '')}.xlsx",
    )

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = max_length + 2
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

        # Set all cells to text format
        for row in worksheet.iter_rows():
            for cell in row:
                cell.number_format = "@"

    logging.info(
        f"✅ Bids with Posting Range within {days_to_scrape} Days Successfully Extracted and saved to {filename}"
    )


def main(days_to_scrape):
    driver = None
    options = None
    max_retries = 3
    retry_count = 0
    bids_data = None

    # Set up temporary folder path
    yesterday = datetime.now() - timedelta(days=1)
    yesterday_folder = yesterday.strftime("%Y-%m-%d")
    script_name = "11_PlanetBids_Hartford"
    temp_folder = f"{script_name}_IN_PROGRESS"
    temp_folder_path = os.path.join(yesterday_folder, temp_folder)

    # Create IN_PROGRESS folder
    os.makedirs(temp_folder_path, exist_ok=True)

    while retry_count < max_retries:
        try:
            driver, options = setup_driver()

            # Perform login
            login(driver)

            # Add a delay after login
            time.sleep(10)

            set_filters(driver)
            bids_data = process_bids(driver, options, days_to_scrape)
            if bids_data:
                save_to_excel(bids_data, days_to_scrape, temp_folder_path)

                # Clean up temporary download folder first
                temp_download_folder = os.path.join(temp_folder_path, script_name)
                if os.path.exists(temp_download_folder):
                    try:
                        shutil.rmtree(temp_download_folder, ignore_errors=True)
                        print(f"✅ Removed temporary download folder: {temp_download_folder}")
                    except Exception as e:
                        print(f"Error removing temporary download folder: {e}")

                # Rename folder to mark completion
                completed_folder = f"{script_name}_COMPLETED"
                completed_folder_path = os.path.join(yesterday_folder, completed_folder)
                if os.path.exists(completed_folder_path):
                    shutil.rmtree(completed_folder_path)
                os.rename(temp_folder_path, completed_folder_path)

                logging.info(
                    f"🎉 All Bids and Attachments Extraction for the last {days_to_scrape} days Successfully Completed"
                )
                play_notification_sound()

            else:
                logging.info(
                    f"No bids were found in the last {days_to_scrape} days to save to Excel."
                )
            break

        except Exception as e:
            logging.error(
                f"An error occurred (attempt {retry_count + 1}/{max_retries}): {str(e)}"
            )
            logging.error(f"Traceback: {traceback.format_exc()}")
            play_notification_sound()
            retry_count += 1
            if retry_count < max_retries:
                logging.info(f"Retrying in 30 seconds...")
                time.sleep(30)
            else:
                logging.error("Max retries reached. Exiting.")
        finally:
            if driver:
                driver.quit()
            
            # Remove cookies file
            cookie_file = "cookies.pkl"
            if os.path.exists(cookie_file):
                try:
                    os.remove(cookie_file)
                    print("🗑️ Removed cookies file")
                except Exception as e:
                    print(f"Error removing cookies file: {e}")

    if bids_data is None:
        logging.error("Failed to retrieve any bid data after all attempts.")
        play_notification_sound()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Scrape PlanetBids for City of Hartford"
    )
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Number of days to look back for bids (default: 2)",
    )
    args = parser.parse_args()

    main(args.days)
