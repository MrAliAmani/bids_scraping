import os
import sys
import time
import json
import random
import argparse
import pandas as pd
import logging
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
	TimeoutException,
	NoSuchElementException,
	WebDriverException,
	StaleElementReferenceException
)
import zipfile
import requests
from urllib.parse import urljoin
import shutil
from selenium.webdriver.common.action_chains import ActionChains
import glob

# Add utils path to system path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.utils import play_notification_sound, safe_move

def get_cache_file():
	"""Get the path to the cache file"""
	script_name = os.path.splitext(os.path.basename(__file__))[0]
	cache_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'cache')
	os.makedirs(cache_dir, exist_ok=True)
	return os.path.join(cache_dir, f'{script_name}_cache.json')

def load_cache():
	"""Load the cache from file"""
	cache_file = get_cache_file()
	try:
		if os.path.exists(cache_file):
			with open(cache_file, 'r') as f:
				return json.load(f)
	except Exception as e:
		print(f"Error loading cache: {str(e)}")
	return {}

def save_cache(cache_data):
	"""Save the cache to file"""
	cache_file = get_cache_file()
	try:
		with open(cache_file, 'w') as f:
			json.dump(cache_data, f, indent=2)
	except Exception as e:
		print(f"Error saving cache: {str(e)}")

def clean_old_cache_entries(cache_data):
	"""Remove cache entries older than 3 months"""
	three_months_ago = (datetime.now() - timedelta(days=90)).strftime('%Y-%m-%d')
	return {
		url: data for url, data in cache_data.items()
		if data.get('posted_date', '') >= three_months_ago
	}

def should_process_bid(url, bid_number, posted_date, cache_data, days_back=1):
	"""Check if a bid should be processed based on cache data and date range"""
	# First check if the bid is within the date range
	try:
		bid_date = datetime.strptime(posted_date, '%Y-%m-%d')
		cutoff_date = datetime.now() - timedelta(days=days_back)
		if bid_date.date() < cutoff_date.date():
			print(f"Bid {bid_number} posted on {posted_date} is outside the {days_back}-day range, skipping...")
			return False
	except Exception as e:
		print(f"Error checking date range for bid {bid_number}: {str(e)}")
		return False

	# Load existing cache
	existing_cache = load_cache()
	
	# Check if bid number exists in any cache entry
	for _, data in existing_cache.items():
		if data.get('bid_number') == bid_number:
			print(f"Bid {bid_number} found in cache, skipping...")
			return False
	
	# Only remove entries older than 3 months
	three_months_ago = (datetime.now() - timedelta(days=90)).strftime('%Y-%m-%d')
	
	# Keep all entries except those older than 3 months
	for key, data in list(existing_cache.items()):
		if data.get('posted_date', '') < three_months_ago:
			print(f"Removing cache entry older than 3 months: {data.get('bid_number')} posted on {data.get('posted_date')}")
			del existing_cache[key]
	
	# Create unique key with timestamp to prevent overwriting
	timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
	unique_key = f"{bid_number}_{timestamp}"
	
	# Add new bid to cache (append to existing cache)
	existing_cache[unique_key] = {
		'bid_number': bid_number,
		'url': url,
		'posted_date': posted_date,
		'last_checked': datetime.now().strftime('%Y-%m-%d')
	}
	
	# Save updated cache
	save_cache(existing_cache)
	return True

def get_base_folder():
	"""Get the base folder path for yesterday's date"""
	yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
	return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
					   yesterday)

def setup_logger():
	"""Set up logging configuration"""
	try:
		base_folder = get_base_folder()
		if not os.path.exists(base_folder):
			os.makedirs(base_folder)
			
		log_file = os.path.join(base_folder, '19_Pennsylvania_eMarketplace.log')
		
		logging.basicConfig(
			level=logging.INFO,
			format='%(asctime)s - %(levelname)s - %(message)s',
			handlers=[
				logging.FileHandler(log_file, encoding='utf-8'),
				logging.StreamHandler()
			]
		)
		return logging.getLogger(__name__)
	except Exception as e:
		print(f"Error setting up logger: {str(e)}")
		sys.exit(1)

# Initialize logger
logger = setup_logger()

def print_start_message():
	"""Display start message"""
	print("🟢 Bids Extraction Started")
	logger.info("[START] 🟢 Bids Extraction Started")

def notify_completion():
	"""Notify successful completion"""
	print("\n✅ All Bids and Attachments Extraction Successfully Completed")
	logger.info("[COMPLETE] ✅ All Bids and Attachments Extraction Successfully Completed")
	play_notification_sound()

def handle_error(error_msg, play_sound=True, pause=True):
	"""Handle errors with notifications and optional pause"""
	logger.error(error_msg)
	print(f"\nError: {error_msg}")
	
	if play_sound:
		play_notification_sound()
		
	if pause:
		user_input = input("Press Enter to continue, or 'q' to quit: ")
		if user_input.lower() == 'q':
			sys.exit(1)

def setup_driver():
	"""Setup WebDriver with anti-bot measures and download settings"""
	try:
		options = webdriver.ChromeOptions()
		
		# Anti-bot measures
		options.add_argument('--disable-blink-features=AutomationControlled')
		options.add_experimental_option("excludeSwitches", ["enable-automation"])
		options.add_experimental_option('useAutomationExtension', False)
		options.add_argument('--disable-extensions')
		
		# Create and set download folder
		yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
		script_name = os.path.splitext(os.path.basename(__file__))[0]
		base_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), yesterday)
		in_progress_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
		download_folder = os.path.join(in_progress_folder, script_name)
		
		# Create all necessary folders
		os.makedirs(download_folder, exist_ok=True)
		print(f"\nConfigured download folder: {download_folder}")
		
		# Download settings
		prefs = {
			'download.default_directory': download_folder,
			'download.prompt_for_download': False,
			'download.directory_upgrade': True,
			'safebrowsing.enabled': True,
			'profile.default_content_settings.popups': 0,
			'profile.default_content_setting_values.automatic_downloads': 1
		}
		options.add_experimental_option('prefs', prefs)
		
		# Performance settings
		options.add_argument('--no-sandbox')
		options.add_argument('--disable-dev-shm-usage')
		options.add_argument('--disable-gpu')
		options.add_argument('--disable-infobars')
		options.add_argument('--disable-notifications')
		
		# Create driver with configured options
		driver = webdriver.Chrome(options=options)
		driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
			'source': '''
				Object.defineProperty(navigator, 'webdriver', {
					get: () => undefined
				})
			'''
		})
		
		return driver
		
	except Exception as e:
		logger.error(f"[ERROR] Failed to setup WebDriver: {str(e)}")
		return None

def wait_for_element(driver, by, value, timeout=10, retries=3):
	"""Wait for element to be present and visible with retries"""
	for attempt in range(retries):
		try:
			element = WebDriverWait(driver, timeout).until(
				EC.presence_of_element_located((by, value))
			)
			# Verify element is not stale
			element.is_enabled()
			return element
		except (TimeoutException, StaleElementReferenceException):
			if attempt < retries - 1:
				time.sleep(2)
				continue
			return None
	return None

def wait_for_page_load(driver, timeout=30):
	"""Wait for page to fully load"""
	try:
		# Wait for document ready state
		WebDriverWait(driver, timeout).until(
			lambda d: d.execute_script('return document.readyState') == 'complete'
		)
		
		# Wait for any AJAX calls to complete
		time.sleep(5)
		
		return True
	except Exception as e:
		logger.error(f"[ERROR] Page load issue: {str(e)}")
		return False

def format_date(date_str):
	"""Convert date string to YYYY-MM-DD format"""
	try:
		# Clean the input string
		date_str = date_str.strip()
		if ' ' in date_str:  # Remove time component if present
			date_str = date_str.split(' ')[0]
			
		# Try different date formats
		formats = ['%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%m-%d-%Y']
		for fmt in formats:
			try:
				return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
			except ValueError:
				continue
				
		logger.warning(f"[WARNING] Could not parse date format: {date_str}")
		return date_str
		
	except Exception as e:
		logger.error(f"[ERROR] Error formatting date {date_str}: {str(e)}")
		return date_str

def is_within_date_range(posted_date_str, days_back):
	"""Check if the posted date is within the specified range"""
	try:
		posted_date = datetime.strptime(posted_date_str.strip(), '%Y-%m-%d')
		cutoff_date = datetime.now() - timedelta(days=days_back)
		return posted_date >= cutoff_date
	except Exception as e:
		logger.error(f"[ERROR] Error checking date range: {str(e)}")
		return False

def scrape_pennsylvania_emarketplace(driver, days_back=1):
	"""Scrape bids from Pennsylvania eMarketplace"""
	base_folder = get_base_folder()
	script_name = os.path.splitext(os.path.basename(__file__))[0]
	script_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
	os.makedirs(script_folder, exist_ok=True)
	
	# Load cache
	cache_data = load_cache()
	
	try:
		# Access URL
		url = "https://www.emarketplace.state.pa.us/Search.aspx?Cpg=3"
		logger.info(f"[INFO] Accessing URL: {url}")
		print(f"\nAccessing URL: {url}")
		driver.get(url)
		
		if not wait_for_page_load(driver):
			handle_error("Page failed to load")
			return False
		
		# Set date filter for yesterday
		yesterday = (datetime.now() - timedelta(days=days_back)).strftime('%m/%d/%Y')
		
		date_input = wait_for_element(driver, By.ID, "ctl00_MainBody_txtDatePre")
		if date_input:
			date_input.clear()
			date_input.send_keys(yesterday)
			print(f"Set date filter to: {yesterday}")
		else:
			handle_error("Could not find date input field")
			return False
			
		# Click search button
		search_button = wait_for_element(driver, By.ID, "ctl00_MainBody_btnSearch")
		if search_button:
			search_button.click()
			if not wait_for_page_load(driver):
				handle_error("Page failed to load after search")
				return False
			print("Clicked search button")
		else:
			handle_error("Could not find search button")
			return False
			
		# Click Solicitation Start Date to sort
		sort_link = wait_for_element(driver, By.XPATH, "//a[contains(@href, 'Sort$BidStartDate')]")
		if sort_link:
			sort_link.click()
			if not wait_for_page_load(driver):
				handle_error("Page failed to load after sorting")
				return False
			print("Sorted by Solicitation Start Date")
		else:
			handle_error("Could not find sort link")
			return False
			
		# Set results to show ALL
		select_element = wait_for_element(driver, By.ID, "ctl00_MainBody_ddlRows")
		if select_element:
			from selenium.webdriver.support.ui import Select
			Select(select_element).select_by_value("32767")  # ALL option
			if not wait_for_page_load(driver):
				handle_error("Page failed to load after changing results count")
				return False
			print("Set results to show ALL")
		else:
			handle_error("Could not find results dropdown")
			return False
			
		# Process bids
		bids_data = []
		main_window = driver.current_window_handle
		
		# Get all bid rows
		bid_rows = driver.find_elements(By.XPATH, "//tr[td/a[contains(@id, 'HyperLink1')]]")
		if not bid_rows:
			print("No bids found")
			return True
			
		print(f"\nFound {len(bid_rows)} bids")
		
		for row in bid_rows:
			try:
				# Get bid link and info before clicking
				link_element = row.find_element(By.XPATH, ".//a[contains(@id, 'HyperLink1')]")
				solicitation_number = link_element.text.strip()
				bid_url = link_element.get_attribute('href')
				
				# Get other basic info
				title = row.find_element(By.XPATH, ".//span[contains(@id, 'lblTitle')]").text.strip()
				agency = row.find_element(By.XPATH, ".//td[5]").text.strip()
				posted_date = format_date(row.find_element(By.XPATH, ".//td[8]").text.strip())
				response_date = format_date(row.find_element(By.XPATH, ".//td[9]").text.strip().split(' ')[0])
				
				print(f"\nProcessing bid: {solicitation_number} (Posted: {posted_date})")
				
				# Check if bid should be processed (passing days_back)
				if not should_process_bid(bid_url, solicitation_number, posted_date, cache_data, days_back):
					continue
				
				# Create initial bid info
				bid_info = {
					'Posted Date': posted_date,
					'Response Date': response_date,
					'Solicitation Number': solicitation_number,
					'Solicitation Title': title,
					'Agency': agency,
					'Bid Detail Page URL': bid_url,
					'Category': '',
					'Notice Type': '',
					'Description': '',
					'Additional Summary': '',
					'Contracting Office Address': '',
					'Contact Information': '',
					'Attachments': ''
				}
				
				# Open bid in new tab
				driver.execute_script("window.open(arguments[0]);", bid_url)
				
				# Switch to new tab
				new_window = [window for window in driver.window_handles if window != main_window][0]
				driver.switch_to.window(new_window)
				
				if not wait_for_page_load(driver):
					print("Bid detail page failed to load")
					driver.close()
					driver.switch_to.window(main_window)
					continue
				
				# Extract additional details
				bid_details = extract_bid_details_from_page(driver, bid_info)
				if bid_details:
					# Download attachments
					attachments = download_bid_attachments(driver, script_folder, bid_details['Solicitation Number'])
					bid_details['Attachments'] = ', '.join(attachments) if attachments else ''
					
					bids_data.append(bid_details)
					
					# Update Excel after each bid
					update_excel_file(script_folder, script_name, bids_data)
				
				# Close tab and switch back to main window
				driver.close()
				driver.switch_to.window(main_window)
				
			except Exception as e:
				logger.error(f"[ERROR] Error processing bid: {str(e)}")
				print(f"Error processing bid: {str(e)}")
				# Make sure we're back on the main window
				if driver.current_window_handle != main_window:
					driver.close()
					driver.switch_to.window(main_window)
				continue
		
		return True
		
	except Exception as e:
		logger.error(f"[ERROR] Error in scrape_pennsylvania_emarketplace: {str(e)}")
		return False

def extract_bid_info_from_row(row):
	"""Extract bid information from a table row"""
	try:
		print("\nExtracting bid information from row...")
		
		# Get solicitation number and URL
		try:
			link_element = row.find_element(By.XPATH, ".//a[contains(@id, 'HyperLink1')]")
			solicitation_number = link_element.text.strip()
			relative_url = link_element.get_attribute('href')
			
			full_url = urljoin("https://www.emarketplace.state.pa.us/", relative_url)
			print(f"Found solicitation number: {solicitation_number}")
		
		except Exception as e:
			print(f"Error extracting solicitation info: {str(e)}")
			return None
		
		# Get title
		try:
			title_element = row.find_element(By.XPATH, ".//span[contains(@id, 'lblTitle')]")
			title = title_element.text.strip()
			print(f"Found title: {title}")
		
		except Exception as e:
			print(f"Error extracting title: {str(e)}")
			title = ""
		
		# Get agency
		try:
			agency = row.find_element(By.XPATH, ".//td[5]").text.strip()
			print(f"Found agency: {agency}")
		
		except Exception as e:
			print(f"Error extracting agency: {str(e)}")
			agency = ""
		
		# Get dates
		try:
			
			posted_date = format_date(row.find_element(By.XPATH, ".//td[8]").text.strip())
			print(f"Found posted date: {posted_date}")
			
			response_date = format_date(row.find_element(By.XPATH, ".//td[9]").text.strip().split(' ')[0])
			print(f"Found response date: {response_date}")
		
		except Exception as e:
			print(f"Error extracting dates: {str(e)}")
			return None
		
		bid_info = {
			'Posted Date': posted_date,
			'Response Date': response_date,
			'Solicitation Number': solicitation_number,
			'Solicitation Title': title,
			'Agency': agency,
			'Bid Detail Page URL': full_url,
			'Category': '',
			'Notice Type': '',
			'Description': '',
			'Additional Summary': '',
			'Contracting Office Address': '',
			'Contact Information': '',
			'Attachments': ''
		}
		
		print("Successfully extracted bid information")
		return bid_info
		
	except Exception as e:
		logger.error(f"[ERROR] Error extracting bid info from row: {str(e)}")
		print(f"Error extracting bid info from row: {str(e)}")
		return None

def extract_bid_details_from_page(driver, bid_info):
	"""Extract additional details from bid detail page"""
	try:
		# Initialize with existing info
		details = bid_info.copy()
		
		# Wait for page to load
		time.sleep(3)
		
		# Get notice type
		notice_type = wait_for_element(driver, By.ID, "ctl00_MainBody_lblBidTitle", timeout=5)
		if notice_type:
			details['Notice Type'] = notice_type.text.strip()
			
		# Get description
		description = wait_for_element(driver, By.ID, "ctl00_MainBody_lblDesc", timeout=5)
		if description:
			details['Description'] = description.text.strip()
			
		# Get additional summary
		try:
			summary = driver.find_element(By.XPATH, "//td[contains(text(), 'Bids must be received')]")
			if summary:
				details['Additional Summary'] = summary.text.strip()
		except:
			pass
			
		# Get contracting office address
		address = wait_for_element(driver, By.ID, "ctl00_MainBody_lblLocation", timeout=5)
		if address:
			details['Contracting Office Address'] = address.text.strip()
			
		# Get contact information
		contact_info = []
		
		first_name = wait_for_element(driver, By.ID, "ctl00_MainBody_lblFName", timeout=5)
		last_name = wait_for_element(driver, By.ID, "ctl00_MainBody_lblLName", timeout=5)
		if first_name and last_name:
			contact_info.append(f"Name: {first_name.text.strip()} {last_name.text.strip()}")
			
		phone = wait_for_element(driver, By.ID, "ctl00_MainBody_lblPhone", timeout=5)
		if phone:
			contact_info.append(f"Phone: {phone.text.strip()}")
			
		email = wait_for_element(driver, By.ID, "ctl00_MainBody_lblEmail", timeout=5)
		if email:
			contact_info.append(f"Email: {email.text.strip()}")
			
		details['Contact Information'] = "\n".join(contact_info)
		
		return details
		
	except Exception as e:
		logger.error(f"[ERROR] Error extracting bid details: {str(e)}")
		return None

def download_bid_attachments(driver, script_folder, bid_number):
	"""Download attachments for a bid"""
	try:
		print("\nDownloading attachments...")
		downloaded_files = []
		
		# Create bid folder
		bid_folder = os.path.join(script_folder, bid_number)
		os.makedirs(bid_folder, exist_ok=True)
		
		# Find attachment links
		attachment_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'FileDownload.aspx')]")
		
		if not attachment_links:
			print("No attachments found")
			return []
			
		print(f"Found {len(attachment_links)} attachments")
		
		# Get download folder
		download_folder = os.path.join(script_folder, os.path.splitext(os.path.basename(__file__))[0])
		
		# Clear download folder before starting
		for file in os.listdir(download_folder):
			try:
				os.remove(os.path.join(download_folder, file))
			except:
				pass
		
		# Download each attachment
		for link in attachment_links:
			try:
				original_filename = link.text.strip()
				print(f"\nDownloading: {original_filename}")
				
				# Click to download
				link.click()
				
				# Wait for download to start
				time.sleep(3)
				
				# Wait for download to complete (max 30 seconds per file)
				max_wait = 30
				start_time = time.time()
				file_downloaded = False
				
				while time.time() - start_time < max_wait:
					
					files = [f for f in os.listdir(download_folder) 
							if not f.endswith(('.crdownload', '.tmp'))]
					
					if files:
						downloaded_file = files[0]  # Get the first completed download
						source_path = os.path.join(download_folder, downloaded_file)
						
						# Use original filename with extension
						_, ext = os.path.splitext(original_filename)
						if not ext:
							ext = os.path.splitext(downloaded_file)[1]
						
						target_filename = original_filename
						if not target_filename.endswith(ext):
							target_filename += ext
							
						target_path = os.path.join(bid_folder, target_filename)
						
						# Handle duplicate filenames
						counter = 1
						while os.path.exists(target_path):
							base, ext = os.path.splitext(target_filename)
							target_path = os.path.join(bid_folder, f"{base}_{counter}{ext}")
							counter += 1
						
						# Move file to final location
						shutil.move(source_path, target_path)
						downloaded_files.append(os.path.basename(target_path))
						print(f"Successfully downloaded: {os.path.basename(target_path)}")
						file_downloaded = True
						break
					
					time.sleep(1)
				
				if not file_downloaded:
					print(f"Failed to download: {original_filename}")
				
				# Clear any remaining partial downloads
				for file in os.listdir(download_folder):
					try:
						os.remove(os.path.join(download_folder, file))
					except:
						pass
				
			except Exception as e:
				print(f"Error downloading attachment: {str(e)}")
				continue
				
			# Wait between downloads
			time.sleep(2)
		
		print(f"\nDownload summary:")
		print(f"Total attachments found: {len(attachment_links)}")
		print(f"Successfully downloaded: {len(downloaded_files)}")
		if downloaded_files:
			print("Downloaded files:")
			for idx, file in enumerate(downloaded_files, 1):
				print(f"{idx}. {file}")
		
		return downloaded_files
		
	except Exception as e:
		logger.error(f"[ERROR] Error downloading attachments: {str(e)}")
		print(f"Error downloading attachments: {str(e)}")
		return []

def update_excel_file(script_folder, script_name, bids_data):
	"""Update the Excel file with current bid data"""
	try:
		df = pd.DataFrame(bids_data)
		df['SL No'] = range(1, len(df) + 1)
		
		columns = [
			'SL No',
			'Posted Date',
			'Response Date',
			'Notice Type',
			'Solicitation Number',
			'Solicitation Title',
			'Agency',
			'Category',
			'Description',
			'Additional Summary',
			'Contracting Office Address',
			'Contact Information',
			'Bid Detail Page URL',
			'Attachments'
		]
		df = df.reindex(columns=columns)
		
		excel_path = os.path.join(script_folder, f'{script_name}.xlsx')
		with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
			df.to_excel(writer, index=False)
			worksheet = writer.sheets['Sheet1']
			for column in worksheet.columns:
				for cell in column:
					cell.number_format = '@'
		
		print(f"\nUpdated Excel file with {len(bids_data)} bids: {excel_path}")
		return True
	except Exception as e:
		logger.error(f"[ERROR] Error updating Excel: {str(e)}")
		return False

def complete_scraping():
	"""Mark scraping as complete by renaming the folder"""
	try:
		base_folder = get_base_folder()
		script_name = os.path.splitext(os.path.basename(__file__))[0]
		in_progress_path = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
		completed_path = os.path.join(base_folder, f"{script_name}_COMPLETED")
		
		if os.path.exists(in_progress_path):
			try:
				os.rename(in_progress_path, completed_path)
				logger.info(f"[SUCCESS] Renamed folder to {completed_path}")
			except Exception as e:
				logger.error(f"[ERROR] Failed to rename folder: {str(e)}")
	except Exception as e:
		logger.error(f"[ERROR] Error in complete_scraping: {str(e)}")

def main():
	"""Main function to run the scraper"""
	try:
		# Parse command line arguments
		parser = argparse.ArgumentParser(description='Scrape Pennsylvania eMarketplace bids')
		parser.add_argument('--days', type=int, default=1,
						  help='Number of days back to scrape (default: 1)')
		args = parser.parse_args()
		
		# Step 1: Start the Process
		print_start_message()
		
		# Setup logging and get working directory
		base_folder = get_base_folder()
		script_name = os.path.splitext(os.path.basename(__file__))[0]
		
		print(f"\nConfiguration:")
		print(f"- Days to look back: {args.days}")
		print(f"- Working folder: {base_folder}")
		logger.info(f"[CONFIG] Days to look back: {args.days}")
		
		# Initialize WebDriver
		driver = setup_driver()
		if not driver:
			handle_error("[ERROR] Failed to initialize WebDriver")
			return False
			
		try:
			# Run the scraper
			success = scrape_pennsylvania_emarketplace(driver, args.days)
			
			if success:
				# Mark scraping as complete
				complete_scraping()
				notify_completion()
				return True
			else:
				handle_error("[ERROR] Scraping failed")
				return False
				
		finally:
			if driver:
				driver.quit()
		
	except Exception as e:
		error_msg = f"[FATAL] Fatal error in main execution: {str(e)}"
		logger.error(error_msg)
		print(f"\nFatal Error: {str(e)}")
		play_notification_sound()
		return False

if __name__ == "__main__":
	main()