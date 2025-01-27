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

def should_process_bid(url, bid_number, posted_date, cache_data):
	"""Check if a bid should be processed based on cache data"""
	# First load existing cache
	existing_cache = load_cache()
	
	# Check if bid number exists in any cache entry
	for _, data in existing_cache.items():
		if data.get('bid_number') == bid_number:
			print(f"Bid {bid_number} found in cache, skipping...")
			return False
	
	# Remove entries older than 3 months but keep all other existing entries
	three_months_ago = (datetime.now() - timedelta(days=90)).strftime('%Y-%m-%d')
	cleaned_cache = existing_cache.copy()
	
	# Only remove entries older than 3 months
	for key, data in list(cleaned_cache.items()):
		if data.get('posted_date', '') < three_months_ago:
			del cleaned_cache[key]
	
	# Create unique key with timestamp to prevent overwriting
	timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
	unique_key = f"{bid_number}_{timestamp}"
	
	# Add new bid to cache
	cleaned_cache[unique_key] = {
		'bid_number': bid_number,
		'url': url,
		'posted_date': posted_date,
		'last_checked': datetime.now().strftime('%Y-%m-%d')
	}
	
	save_cache(cleaned_cache)
	return True

def get_base_folder():
	"""Get the base folder path for yesterday's date"""
	yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
	return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
					   yesterday)

def setup_logger():
	"""Set up logging configuration"""
	try:
		logging.basicConfig(
			level=logging.INFO,
			format='%(asctime)s - %(levelname)s - %(message)s',
			handlers=[
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
	print("[START] Bids Extraction Started")
	logger.info("[START] Bids Extraction Started")

def notify_completion():
	"""Notify successful completion"""
	print("\n[COMPLETE] All Bids and Attachments Extraction Successfully Completed")
	logger.info("[COMPLETE] All Bids and Attachments Extraction Successfully Completed")
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
	"""Setup WebDriver with anti-bot measures"""
	try:
		options = webdriver.ChromeOptions()
		
		# Anti-bot measures
		options.add_argument('--disable-blink-features=AutomationControlled')
		options.add_experimental_option("excludeSwitches", ["enable-automation"])
		options.add_experimental_option('useAutomationExtension', False)
		options.add_argument('--disable-extensions')
		
		# Create and set download folder
		script_name = os.path.splitext(os.path.basename(__file__))[0]
		base_folder = get_base_folder()
		script_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
		download_folder = os.path.join(script_folder, script_name)
		
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

def complete_scraping():
	"""Mark scraping as complete by renaming the folder"""
	try:
		base_folder = get_base_folder()
		script_name = os.path.splitext(os.path.basename(__file__))[0]
		in_progress_path = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
		completed_path = os.path.join(base_folder, f"{script_name}_COMPLETED")
		
		# Clean up temporary download folder first
		try:
			download_folder = os.path.join(in_progress_path, script_name)
			if os.path.exists(download_folder):
				shutil.rmtree(download_folder, ignore_errors=True)
				logger.info(f"[SUCCESS] Removed temporary download folder: {download_folder}")
		except Exception as e:
			logger.error(f"[ERROR] Failed to remove download folder: {str(e)}")
		
		if os.path.exists(in_progress_path):
			try:
				os.rename(in_progress_path, completed_path)
				logger.info(f"[SUCCESS] Renamed folder to {completed_path}")
			except Exception as e:
				logger.error(f"[ERROR] Failed to rename folder: {str(e)}")
	except Exception as e:
		logger.error(f"[ERROR] Error in complete_scraping: {str(e)}")

def update_excel_file(script_folder, script_name, bids_data):
	"""Update the Excel file with current bid data"""
	try:
		df = pd.DataFrame(bids_data)
		df['SL No'] = range(1, len(df) + 1)
		
		# For each bid, verify and update Attachments column with actual files in the bid folder
		for index, row in df.iterrows():
			bid_number = row['Solicitation Number']
			bid_folder = os.path.join(script_folder, bid_number)
			
			print(f"\nVerifying attachments for bid {bid_number}:")
			
			# Get actual files in the bid folder
			if os.path.exists(bid_folder):
				files = [f for f in os.listdir(bid_folder) if os.path.isfile(os.path.join(bid_folder, f))]
				if files:
					files.sort()  # Sort for consistent order
					attachments_str = ', '.join(files).rstrip(',')  # Remove any trailing comma
					
					# Only update if different from current value
					current_attachments = str(row.get('Attachments', ''))
					if attachments_str != current_attachments:
						df.at[index, 'Attachments'] = attachments_str
						print(f"Updated attachments list:")
						for file in files:
							print(f"- {file}")
					else:
						print(f"Attachments already up to date: {attachments_str}")
				else:
					df.at[index, 'Attachments'] = ''
					print("No files found in bid folder")
			else:
				df.at[index, 'Attachments'] = ''
				print(f"Bid folder not found: {bid_folder}")
		
		# Ensure correct column order
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
		
		# Create empty columns if they don't exist
		for col in columns:
			if col not in df.columns:
				df[col] = ''
		
		# Reorder columns
		df = df[columns]
		
		# Save Excel in the script folder (not in bid folders)
		excel_path = os.path.join(script_folder, f'{script_name}.xlsx')
		
		with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
			df.to_excel(writer, index=False)
			worksheet = writer.sheets['Sheet1']
			
			# Set all cells to text format
			for column in worksheet.columns:
				for cell in column:
					cell.number_format = '@'
		
		print(f"\nUpdated Excel file with {len(bids_data)} bids: {excel_path}")
		logger.info(f"[SUCCESS] Updated Excel file with {len(bids_data)} bids: {excel_path}")
		
		# Verify the Excel file was saved
		if os.path.exists(excel_path):
			# Read back and verify attachments
			verification_df = pd.read_excel(excel_path)
			for _, row in verification_df.iterrows():
				bid_number = row['Solicitation Number']
				attachments = row['Attachments']
				if attachments:
					print(f"\nVerified attachments for bid {bid_number}:")
					print(f"- {attachments}")
		
		return True
	except Exception as e:
		error_msg = f"[ERROR] Error updating Excel: {str(e)}"
		logger.error(error_msg)
		print(f"\n{error_msg}")
		return False

def wait_for_page_load(driver, timeout=30):
	"""Wait for page to fully load"""
	try:
		WebDriverWait(driver, timeout).until(
			lambda d: d.execute_script('return document.readyState') == 'complete'
		)
		time.sleep(5)
		return True
	except Exception as e:
		logger.error(f"[ERROR] Error waiting for page load: {str(e)}")
		return False

def apply_filters(driver):
	"""Apply filters to show only open bids"""
	try:
		print("\nApplying filters...")
		logger.info("[INFO] Applying filters")
		
		# Wait longer for initial page load
		time.sleep(10)
		
		# Click the status dropdown
		print("Clicking status dropdown...")
		status_dropdown = WebDriverWait(driver, 20).until(
			EC.element_to_be_clickable((By.CSS_SELECTOR, "div[data-selector='body_x_selRfxStatus']"))
		)
		status_dropdown.click()
		time.sleep(3)

		# Select "Open" status
		print("Selecting 'Open' status...")
		open_option = WebDriverWait(driver, 10).until(
			EC.element_to_be_clickable((By.CSS_SELECTOR, "li[data-value='val']"))
		)
		open_option.click()
		time.sleep(3)

		# Click search button
		print("Clicking search button...")
		search_button = WebDriverWait(driver, 10).until(
			EC.element_to_be_clickable((By.CSS_SELECTOR, "button[id*='cmdSearchBtn']"))
		)
		search_button.click()
		time.sleep(10)  # Wait longer for search results

		print("Filters applied successfully")
		logger.info("[SUCCESS] Filters applied successfully")
		return True
	except Exception as e:
		error_msg = f"[ERROR] Error applying filters: {str(e)}"
		logger.error(error_msg)
		print(f"\n{error_msg}")
		play_notification_sound()
		return False

def get_bid_links(driver, base_url="https://sdbuynet.sandiegocounty.gov"):
	"""Get all bid detail page URLs from the current page using JavaScript"""
	try:
		# Use JavaScript to get bid links
		script = """
			const baseUrl = 'https://sdbuynet.sandiegocounty.gov';
			const bidLinks = document.querySelectorAll('a[id^="body_x_grid_grd_tr_"][id$="_img___colManagegrid"]');
			const bidUrls = [];
			bidLinks.forEach(link => {
				const href = link.getAttribute('href');
				if (href) {
					bidUrls.push(baseUrl + href);
				}
			});
			return bidUrls;
		"""
		urls = driver.execute_script(script)
		print(f"\nFound {len(urls)} bid URLs on current page:")
		for url in urls:
			print(f"- {url}")
		return urls
	except Exception as e:
		logger.error(f"[ERROR] Error getting bid links: {str(e)}")
		return []

def extract_bid_info(driver, current_row):
	"""Extract bid information from the current row"""
	try:
		# Get the current row's cells
		cells = current_row.find_elements(By.TAG_NAME, "td")
		
		if len(cells) < 8:
			logger.warning("[WARNING] Row has insufficient cells")
			return None

		# Posted Date (RFx Begin Date - 3rd column)
		posted_date = cells[2].get_attribute('textContent').strip()
		if not posted_date:
			logger.warning("[WARNING] Missing posted date")
			return None
		posted_date = format_date(posted_date)

		# Response Date (RFx End Date - 4th column)
		response_date = cells[3].get_attribute('textContent').strip()
		if not response_date:
			logger.warning("[WARNING] Missing response date")
			return None
		response_date = format_date(response_date)

		# Solicitation Title (RFx Name - 2nd column)
		title = cells[1].get_attribute('textContent').strip()
		if not title:
			logger.warning("[WARNING] Missing solicitation title")
			return None

		# Solicitation Number (Code - 8th column)
		code = cells[7].get_attribute('textContent').strip()
		if not code:
			logger.warning("[WARNING] Missing solicitation number")
			return None

		bid_info = {
			'Posted Date': posted_date,
			'Response Date': response_date,
			'Solicitation Title': title,
			'Solicitation Number': code,
			'Notice Type': '',
			'Agency': 'County of San Diego',
			'Category': '',
			'Description': '',
			'Additional Summary': '',
			'Contracting Office Address': '',
			'Contact Information': '',
			'Bid Detail Page URL': '',
			'Attachments': ''
		}
		
		print(f"\nExtracted bid info:")
		print(f"- Solicitation Number: {code}")
		print(f"- Posted Date: {posted_date}")
		print(f"- Response Date: {response_date}")
		print(f"- Title: {title}")
		
		return bid_info
		
	except Exception as e:
		error_msg = f"[ERROR] Error extracting bid info: {str(e)}"
		logger.error(error_msg)
		print(f"\n{error_msg}")
		return None

def extract_bid_details(driver, bid_info):
	"""Extract additional details from bid detail page"""
	try:
		details = {
			'Posted Date': bid_info['posted_date'],
			'Response Date': bid_info['response_date'],
			'Notice Type': 'N/A',
			'Solicitation Number': bid_info['solicitation_number'],
			'Solicitation Title': bid_info['solicitation_title'],
			'Agency': 'County of San Diego',
			'Category': '',
			'Description': '',
			'Additional Summary': '',
			'Contracting Office Address': '',
			'Contact Information': '',
			'Bid Detail Page URL': driver.current_url,
			'Attachments': ''
		}

		time.sleep(3)

		try:
			summary_label = driver.find_element(By.XPATH, "//span[contains(text(), 'Summary')]")
			if summary_label:
				next_row = summary_label.find_element(By.XPATH, "following::tr[1]")
				if next_row:
					content = next_row.text.strip()
					if "Name:" in content or "Phone:" in content or "Email:" in content:
						details['Contact Information'] = content
					else:
						details['Description'] = content
		except:
			logger.warning("[WARNING] Summary not found")

		return details

	except Exception as e:
		logger.error(f"[ERROR] Error in extract_bid_details: {str(e)}")
		return None

def download_attachments(driver, bid_folder):
	"""Download all available attachments for a bid"""
	try:
		print("\nProcessing attachments...")
		attachment_filenames = []
		
		# Use the selector from the markdown file
		attachment_links = driver.find_elements(By.CSS_SELECTOR, "a.iv-download-file")
		
		if not attachment_links:
			print("No attachments found")
			return []

		print(f"Found {len(attachment_links)} attachments")
		
		# Get the download folder from driver options
		script_name = os.path.splitext(os.path.basename(__file__))[0]
		base_folder = get_base_folder()
		script_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
		temp_download_folder = os.path.join(script_folder, script_name)
		os.makedirs(temp_download_folder, exist_ok=True)
		os.makedirs(bid_folder, exist_ok=True)  # Ensure bid folder exists
		
		successfully_moved_files = []
		
		for idx, link in enumerate(attachment_links, 1):
			try:
				filename = link.find_element(By.CSS_SELECTOR, "span[data-iv-role='label']").text.strip()
				print(f"\nDownloading attachment {idx}/{len(attachment_links)}: {filename}")
				
				# Click to download
				ActionChains(driver).move_to_element(link).click().perform()
				time.sleep(5)  # Wait for download to start
				
				# Wait for download to complete
				max_wait = 60  # Maximum wait time in seconds
				wait_time = 0
				while wait_time < max_wait:
					if any(f.endswith('.crdownload') or f.endswith('.tmp') 
						  for f in os.listdir(temp_download_folder)):
						time.sleep(2)
						wait_time += 2
						continue
					break
				
				# After download completes, find the actual downloaded file
				downloaded_files = os.listdir(temp_download_folder)
				for file in downloaded_files:
					source = os.path.join(temp_download_folder, file)
					if os.path.isfile(source):
						destination = os.path.join(bid_folder, file)
						try:
							safe_move(source, destination)
							print(f"Successfully moved {file} to bid folder")
							successfully_moved_files.append(file)
						except Exception as move_error:
							logger.error(f"[ERROR] Failed to move {file}: {str(move_error)}")
				
				if wait_time >= max_wait:
					print(f"Timeout waiting for {filename} to download")
					
			except Exception as e:
				error_msg = f"Error downloading attachment: {str(e)}"
				logger.error(error_msg)
				print(f"{error_msg}")
				continue

		# Verify files in bid folder
		if os.path.exists(bid_folder):
			actual_files = [f for f in os.listdir(bid_folder) if os.path.isfile(os.path.join(bid_folder, f))]
			if actual_files:
				actual_files.sort()
				print(f"\nVerified files in bid folder:")
				for file in actual_files:
					print(f"- {file}")
				return actual_files
			else:
				print("No files found in bid folder after download")
		else:
			print(f"Bid folder not found: {bid_folder}")
		
		return successfully_moved_files

	except Exception as e:
		error_msg = f"[ERROR] Error downloading attachments: {str(e)}"
		logger.error(error_msg)
		print(f"\n❌ {error_msg}")
		play_notification_sound()
		return []

def scrape_bids(driver, days_back=1):
	"""Main function to scrape bids"""
	try:
		base_url = "https://sdbuynet.sandiegocounty.gov"
		url = f"{base_url}/page.aspx/en/rfp/request_browse_public"
		
		print(f"\nAccessing URL: {url}")
		driver.get(url)
		
		if not wait_for_page_load(driver):
			logger.error("[ERROR] Initial page load failed")
			return False
			
		if not apply_filters(driver):
			logger.error("[ERROR] Failed to apply filters")
			return False
			
		# First collect all bid URLs and info from 2 pages
		bids_to_process = []
		processed_bid_numbers = set()  # Track unique bid numbers
		pages_to_scrape = 2
		urls_per_page = 15
		total_urls_needed = 30
		start_row_index = 7  # Skip first 7 rows
		
		print("\n=== Starting URL Collection Phase ===")
		print(f"Collecting unique bids across {pages_to_scrape} pages...")
		
		for page_num in range(1, pages_to_scrape + 1):
			try:
				print(f"\nCollecting bids from page {page_num}/{pages_to_scrape}...")
				
				# Wait for page to load after filter/navigation
				
				# Get rows for extracting details
				rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
				
				# Skip header rows and only process actual bid rows
				valid_rows = rows[start_row_index:]
				page_bids = []  # Store unique bids for this page
				
				for row_index, row in enumerate(valid_rows):
					try:
						actual_row_num = row_index + start_row_index + 1
						
						# Extract bid info
						bid_info = extract_bid_info(driver, row)
						if not bid_info:
							continue
							
						bid_number = bid_info['Solicitation Number']
						
						# Skip if we've already processed this bid number
						if bid_number in processed_bid_numbers:
							print(f"\nSkipping duplicate bid {bid_number}")
							continue
							
						processed_bid_numbers.add(bid_number)
						
						# Get bid URL
						try:
							link_element = row.find_element(By.CSS_SELECTOR, "a[id$='_img___colManagegrid']")
							href = link_element.get_attribute('href')
							if not href:
								continue
								
							bid_url = urljoin(base_url, href)
							bid_info['Bid Detail Page URL'] = bid_url
							
							print(f"\nExtracted bid info (#{len(page_bids) + 1} on page {page_num}):")
							print(f"- Solicitation Number: {bid_info['Solicitation Number']}")
							print(f"- Posted Date: {bid_info['Posted Date']}")
							print(f"- Response Date: {bid_info['Response Date']}")
							print(f"- Title: {bid_info['Solicitation Title']}")
							print(f"- URL: {bid_url}")
							
							if not is_within_date_range(bid_info['Posted Date'], days_back):
								print(f"Skipping bid {bid_number} - outside date range")
								continue
							
							if should_process_bid(bid_url, bid_number, bid_info['Posted Date'], {}):
								page_bids.append(bid_info)
								print(f"Added bid {bid_number} to processing queue")
							else:
								print(f"Skipping bid {bid_number} - already processed")
								
							if len(page_bids) >= urls_per_page:
								print(f"\nReached {urls_per_page} unique bids on page {page_num}")
								break
								
						except Exception as e:
							logger.error(f"[ERROR] Error getting bid URL: {str(e)}")
							continue
							
					except Exception as e:
						logger.error(f"[ERROR] Error processing row {actual_row_num}: {str(e)}")
						continue
				
				print(f"\nFound {len(page_bids)} unique bids on page {page_num}")
				bids_to_process.extend(page_bids)
				
				if page_num < pages_to_scrape:
					try:
						next_button = WebDriverWait(driver, 5).until(
							EC.presence_of_element_located((By.ID, "body_x_grid_PagerBtnNextPage"))
						)
						if not next_button.is_enabled():
							print("Reached last page")
							break
						print(f"\nNavigating to page {page_num + 1}...")
						next_button.click()
						time.sleep(10)
					except Exception as e:
						logger.error(f"[ERROR] Error navigating to next page: {str(e)}")
						break
				
			except Exception as e:
				error_msg = f"[ERROR] Error processing page {page_num}: {str(e)}"
				logger.error(error_msg)
				print(f"\n{error_msg}")
				continue
				
		print(f"\n=== URL Collection Phase Complete ===")
		print(f"Found {len(bids_to_process)} bids to process")
		
		if not bids_to_process:
			logger.error("[ERROR] No bids were found to process")
			return False
			
		print("\nCollected Bid URLs:")
		for index, bid in enumerate(bids_to_process, 1):
			print(f"\n{index}. {bid['Solicitation Number']}")
			print(f"   URL: {bid['Bid Detail Page URL']}")
			print(f"   Title: {bid['Solicitation Title']}")
			print(f"   Posted: {bid['Posted Date']}")
			print(f"   Response: {bid['Response Date']}")
		
		# Now process each collected bid
		print("\n=== Starting Bid Processing Phase ===")
		print(f"Processing {len(bids_to_process)} bids...")
		
		bids_data = []
		for index, bid_info in enumerate(bids_to_process, 1):
			try:
				print(f"\nProcessing bid {index}/{len(bids_to_process)}: {bid_info['Solicitation Number']}")
				
				# Create folder with bid number
				script_name = os.path.splitext(os.path.basename(__file__))[0]
				base_folder = get_base_folder()
				script_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
				bid_folder = os.path.join(script_folder, bid_info['Solicitation Number'])
				download_folder = os.path.join(script_folder, script_name)
				os.makedirs(bid_folder, exist_ok=True)
				os.makedirs(download_folder, exist_ok=True)
				
				# Navigate to bid detail page
				print(f"Accessing bid detail page: {bid_info['Bid Detail Page URL']}")
				driver.get(bid_info['Bid Detail Page URL'])
				if not wait_for_page_load(driver):
					logger.error("[ERROR] Failed to load bid detail page")
					continue
				
				# Extract additional details
				try:
					summary_label = WebDriverWait(driver, 10).until(
						EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Summary')]"))
					)
					if summary_label:
						next_row = summary_label.find_element(By.XPATH, "following::tr[1]")
						if next_row:
							content = next_row.text.strip()
							if "Name:" in content or "Phone:" in content or "Email:" in content:
								bid_info['Contact Information'] = content
								print(f"Extracted contact info: {content}")
							else:
								bid_info['Description'] = content
								print(f"Extracted description: {content}")
				except Exception as e:
					logger.warning(f"[WARNING] Error extracting summary: {str(e)}")
					print("Summary information not found")
				
				# Download attachments
				print("\nDownloading attachments...")
				attachments = download_attachments(driver, bid_folder)
				
				# Wait for files to appear in either folder
				print("\nWaiting for files to appear...")
				files_found = False
				attempt = 1
				
				while not files_found:
					# Check download folder first
					if os.path.exists(download_folder):
						download_files = [f for f in os.listdir(download_folder) if os.path.isfile(os.path.join(download_folder, f))]
						if download_files:
							print(f"\nFound files in download folder, moving to bid folder...")
							for file in download_files:
								source = os.path.join(download_folder, file)
								destination = os.path.join(bid_folder, file)
								try:
									safe_move(source, destination)
									print(f"Moved {file} to bid folder")
								except Exception as e:
									print(f"Failed to move {file}: {str(e)}")
					
					# Check bid folder
					if os.path.exists(bid_folder):
						actual_files = [f for f in os.listdir(bid_folder) if os.path.isfile(os.path.join(bid_folder, f))]
						if actual_files:
							files_found = True
							actual_files.sort()  # Sort for consistent order
							attachments_str = ', '.join(actual_files).rstrip(',')  # Remove any trailing comma
							bid_info['Attachments'] = attachments_str
							print(f"\nFound {len(actual_files)} files in bid folder after attempt {attempt}:")
							for file in actual_files:
								print(f"- {file}")
						else:
							print(f"\nAttempt {attempt}: No files found yet, waiting 5 seconds...")
							time.sleep(5)
							attempt += 1
					else:
						print(f"\nAttempt {attempt}: Bid folder not found, waiting 5 seconds...")
						time.sleep(5)
						attempt += 1
						os.makedirs(bid_folder, exist_ok=True)
				
				bids_data.append(bid_info)
				
				# Update Excel after each bid
				if update_excel_file(script_folder, script_name, bids_data):
					print(f"Bid {bid_info['Solicitation Number']} successfully extracted and saved to Excel.")
					if bid_info['Attachments']:
						print(f"   Attachments saved: {bid_info['Attachments']}")
				else:
					logger.error("[ERROR] Failed to update Excel file")
				
			except Exception as e:
				error_msg = f"[ERROR] Error processing bid {bid_info['Solicitation Number']}: {str(e)}"
				logger.error(error_msg)
				print(f"\n{error_msg}")
				play_notification_sound()
				continue
				
		if bids_data:
			# Final Excel update to ensure all attachments are saved
			if update_excel_file(script_folder, script_name, bids_data):
				print(f"\nFinal Excel update completed with {len(bids_data)} bids.")
			print(f"\n{len(bids_data)} bid links, posted {days_back} day(s) ago, have been successfully extracted and saved.")
			return True
			
		logger.error("[ERROR] No bids were found or processed")
		return False

	except Exception as e:
		error_msg = f"[ERROR] Error in scrape_bids: {str(e)}"
		logger.error(error_msg)
		print(f"\n{error_msg}")
		play_notification_sound()
		return False

def main():
	"""Main function to run the scraper"""
	try:
		print_start_message()
		
		base_folder = get_base_folder()
		script_name = os.path.splitext(os.path.basename(__file__))[0]
		
		# Create main folder structure
		script_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
		download_folder = os.path.join(script_folder, script_name)
		os.makedirs(script_folder, exist_ok=True)
		os.makedirs(download_folder, exist_ok=True)
		
		parser = argparse.ArgumentParser(description='Scrape County of San Diego bids')
		parser.add_argument('--days', type=int, default=1,
						  help='Number of days back to scrape (default: 1)')
		args = parser.parse_args()
		
		print(f"\nConfiguration:")
		print(f"- Days to look back: {args.days}")
		print(f"- Working folder: {base_folder}")
		logger.info(f"[CONFIG] Days to look back: {args.days}")

		driver = setup_driver()
		if not driver:
			logger.error("[ERROR] Failed to initialize WebDriver")
			return False

		try:
			success = scrape_bids(driver, args.days)
			if success:
				complete_scraping()
				notify_completion()
				return True
			else:
				logger.error("[ERROR] No bids were scraped successfully")
				return False

		finally:
			if driver:
				driver.quit()

	except Exception as e:
		error_msg = f"[FATAL] Fatal error in main execution: {str(e)}"
		logger.error(error_msg)
		print(f"\nFatal Error: {str(e)}")
		play_notification_sound()
		input("Press Enter to exit...")
		sys.exit(1)

if __name__ == "__main__":
	main()