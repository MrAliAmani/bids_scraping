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
		base_folder = get_base_folder()
		if not os.path.exists(base_folder):
			os.makedirs(base_folder)
			
		log_file = os.path.join(base_folder, '18_ionwave_scraper.log')
		
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
	print("Bids Extraction Started")
	logger.info("[START] Bids Extraction Started")

def notify_completion():
	"""Notify successful completion"""
	print("\nAll Bids and Attachments Extraction Successfully Completed")
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

def pause_on_error():
	"""Pause script execution on error"""
	play_notification_sound()
	input("Script paused. Press Enter to continue...")

def setup_driver_with_proxy(use_proxy=False):
	"""Setup WebDriver with anti-bot measures and optional proxy support"""
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
		
		if use_proxy:
			proxy_options = get_proxy_options()
			if proxy_options:
				options.add_argument(f'--proxy-server={proxy_options}')
		
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

def download_attachment(driver, url, folder_path, filename):
	"""Download an attachment and save it to the specified folder"""
	logger.info(f"[INFO] Starting download of {filename}")
	
	# Clean filename of invalid characters
	filename = "".join(c for c in filename if c.isalnum() or c in ('-', '_', '.', ' '))
	temp_path = os.path.join(folder_path, f"temp_{filename}")
	file_path = os.path.join(folder_path, filename)

	try:
		# Check if file already exists
		if os.path.exists(file_path):
			logger.warning(f"[WARNING] File already exists: {filename}")
			return True

		# Set up request with timeout and headers
		headers = {
			'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
		}
		logger.info(f"[INFO] Requesting URL: {url}")
		response = requests.get(url, stream=True, timeout=30, headers=headers)
		
		if response.status_code == 200:
			# Get total file size for progress tracking
			total_size = int(response.headers.get('content-length', 0))
			block_size = 8192
			downloaded = 0

			# Download to temporary file with progress tracking
			with open(temp_path, 'wb') as f:
				for chunk in response.iter_content(chunk_size=block_size):
					if chunk:
						f.write(chunk)
						downloaded += len(chunk)
						if total_size > 0:
							progress = (downloaded / total_size) * 100
							logger.info(f"[INFO] Download progress for {filename}: {progress:.1f}%")

			# Verify file was downloaded
			if os.path.getsize(temp_path) == 0:
				raise Exception("Downloaded file is empty")

			# Move file from temp to final location
			shutil.move(temp_path, file_path)
			logger.info(f"[SUCCESS] Successfully downloaded {filename}")

			# Handle zip files
			if filename.lower().endswith('.zip'):
				try:
					with zipfile.ZipFile(file_path, 'r') as zip_ref:
						zip_ref.extractall(folder_path)
						extracted_files = zip_ref.namelist()
						logger.info(f"[SUCCESS] Extracted {len(extracted_files)} files from {filename}")
						return extracted_files
				except Exception as e:
					logger.error(f"[ERROR] Failed to extract zip file {filename}: {str(e)}")
					return [filename]
			
			return [filename]

		else:
			error_msg = f"[ERROR] Failed to download {filename}. Status code: {response.status_code}"
			logger.error(error_msg)
			print(error_msg)
			play_notification_sound()
			input("Press Enter to continue...")
			return []

	except Exception as e:
		error_msg = f"[ERROR] Error downloading {filename}: {str(e)}"
		logger.error(error_msg)
		print(error_msg)
		play_notification_sound()
		input("Press Enter to continue...")
		return []
	finally:
		if os.path.exists(temp_path):
			try:
				os.remove(temp_path)
			except Exception as cleanup_error:
				logger.warning(f"[WARNING] Could not remove temporary file {temp_path}: {str(cleanup_error)}")

def wait_for_page_update(driver, timeout=60):
	"""Wait for page to update after sorting"""
	try:
		# Wait for the loading indicator to disappear (if present)
		try:
			WebDriverWait(driver, timeout).until_not(
				EC.presence_of_element_located((By.CSS_SELECTOR, ".rgLoading"))
			)
		except TimeoutException:
			pass

		# Wait for the bid list to be present and visible
		WebDriverWait(driver, timeout).until(
			EC.presence_of_all_elements_located((By.XPATH, "//tr[contains(@id,'rgBidList_ctl00__')]"))
		)
		
		time.sleep(2)  # Additional wait for stability
		return True
	except Exception as e:
		logger.error(f"[ERROR] Error waiting for page update: {str(e)}")
		return False

def sort_bids_by_date(driver):
	"""Sort bids by clicking Bid Issue Date button"""
	try:
		logger.info("[INFO] Attempting to sort bids by date")
		
		# JavaScript to find and click the sort button
		sort_script = """
			let sortButton = document.querySelector("a[onclick*='OpenDate']") || 
						   document.querySelector("a[title='Click here to sort']") ||
						   document.querySelector("#ctl00_mainContent_rgBidList_ctl00 > thead > tr:nth-child(1) > th:nth-child(6) > a");
			if(sortButton) {
				sortButton.click();
				return true;
			}
			return false;
		"""
		
		# Click sort button twice to sort in descending order
		for i in range(2):
			success = driver.execute_script(sort_script)
			if success:
				logger.info(f"[INFO] Sort button click {i+1} successful")
				if not wait_for_page_update(driver):
					raise Exception("Page did not update after sorting")
				time.sleep(2)  # Wait between clicks
			else:
				raise Exception("Sort button not found")
		
		return True
	except Exception as e:
		logger.error(f"[ERROR] Failed to sort bids: {str(e)}")
		return False

def get_proxy_options():
	"""Get Chrome options configured with proxy settings"""
	logger.info("[INFO] Setting up proxy options")
	
	try:
		proxy_list = [
			"proxy1.example.com:8080",
			"proxy2.example.com:8080",
			# Add more proxies as needed
		]
		
		# Try each proxy
		for proxy in proxy_list:
			try:
				logger.info(f"[INFO] Attempting to use proxy: {proxy}")
				return proxy
			except Exception as e:
				logger.warning(f"[WARNING] Failed to configure proxy {proxy}: {str(e)}")
				continue
		
		logger.warning("[WARNING] No working proxy found, using direct connection")
		return None
		
	except Exception as e:
		logger.error(f"[ERROR] Error configuring proxy options: {str(e)}")
		return None

def wait_for_fresh_element(driver, locator, timeout=10, retries=3):
	"""Wait for element to be fresh (not stale) and return it"""
	for attempt in range(retries):
		try:
			element = WebDriverWait(driver, timeout).until(
				EC.presence_of_element_located(locator)
			)
			# Verify element is not stale
			element.is_enabled()
			return element
		except (StaleElementReferenceException, TimeoutException):
			if attempt < retries - 1:
				time.sleep(2)
				continue
			else:
				return None
	return None

def get_all_dates_from_main_page(driver):
	"""Get all posted and response dates from the main page at once"""
	try:
		# Wait for the table to be present
		WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.XPATH, "//tr[contains(@id,'rgBidList_ctl00__')]"))
		)
		
		# Get all rows at once
		rows = driver.find_elements(By.XPATH, "//tr[contains(@id,'rgBidList_ctl00__')]")
		
		bids_info = []
		for row in rows:
			try:
				# Extract all needed information from the row
				bid_info = {
					'solicitation_number': row.find_element(By.XPATH, ".//td[2]").text.strip(),
					'solicitation_title': row.find_element(By.XPATH, ".//td[3]").text.strip(),
					'agency': row.find_element(By.XPATH, ".//td[5]").text.strip(),
					'posted_date': row.find_element(By.XPATH, ".//td[contains(@class,'rgSorted')]").text.strip(),
					'response_date': row.find_element(By.XPATH, ".//td[7]").text.strip()
				}
				
				# Format dates
				bid_info['formatted_posted_date'] = format_date(bid_info['posted_date'])
				bid_info['formatted_response_date'] = format_date(bid_info['response_date'].split(' ')[0])
				
				bids_info.append(bid_info)
				
				print(f"\nFound bid:")
				print(f"Solicitation Number: {bid_info['solicitation_number']}")
				print(f"Posted Date: {bid_info['posted_date']} -> {bid_info['formatted_posted_date']}")
				print(f"Response Date: {bid_info['response_date']} -> {bid_info['formatted_response_date']}")
				
			except Exception as e:
				logger.error(f"[ERROR] Error extracting data from row: {str(e)}")
				continue
				
		return bids_info
	except Exception as e:
		logger.error(f"[ERROR] Error getting dates from main page: {str(e)}")
		return []

def return_to_listing(driver):
	"""Click the Return to Listing button to go back to main page"""
	try:
		# Find and click the Return to Listing button
		return_script = """
			const returnButton = Array.from(document.querySelectorAll('span.rtbText')).find(
				el => el.textContent.trim() === 'Return to Listing'
			);
			if (returnButton) {
				returnButton.click();
				return true;
			}
			return false;
		"""
		success = driver.execute_script(return_script)
		if success:
			time.sleep(3)  # Wait for page to update
			return wait_for_page_load(driver)
		return False
	except Exception as e:
		logger.error(f"[ERROR] Failed to return to listing: {str(e)}")
		return False

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

def scrape_ionwave_site(url, days_back=1, use_proxy=False):
	"""Scrape bids from an Ionwave site"""
	driver = None
	base_folder = get_base_folder()
	script_name = os.path.splitext(os.path.basename(__file__))[0]
	script_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
	os.makedirs(script_folder, exist_ok=True)
	
	# Load cache
	cache_data = load_cache()
	
	try:
		# Initialize WebDriver
		print("\nInitializing WebDriver...")
		driver = setup_driver_with_proxy(use_proxy)
		if not driver:
			logger.error("[ERROR] Failed to initialize WebDriver")
			return False

		# Access URL
		logger.info(f"[INFO] Accessing URL: {url}")
		print(f"\nAccessing URL: {url}")
		driver.get(url)
		time.sleep(5)

		if not wait_for_page_load(driver):
			return True

		# Sort bids by date
		if not sort_bids_by_date(driver):
			return True

		# Get all dates from main page first
		bids_info = get_all_dates_from_main_page(driver)
		if not bids_info:
			logger.info("[INFO] No bids found on page")
			return True

		bids_data = []
		found_recent_bids = False

		# Process each bid that's within date range
		for bid_info in bids_info:
			try:
				if not is_within_date_range(bid_info['formatted_posted_date'], days_back):
					print(f"Bid {bid_info['solicitation_number']} is outside date range")
					continue

				found_recent_bids = True
				
				# Check cache before processing
				bid_url = driver.current_url
				if not should_process_bid(bid_url, bid_info['solicitation_number'], 
									   bid_info['formatted_posted_date'], cache_data):
					continue
				
				# Click view button for this bid
				try:
					view_button = WebDriverWait(driver, 10).until(
						EC.element_to_be_clickable((By.XPATH, f"//tr[contains(@id,'rgBidList_ctl00__')]//td[2][contains(text(),'{bid_info['solicitation_number']}')]/..//td[1]/span[@class='flaticon-grid_View']"))
					)
					view_button.click()
					time.sleep(3)
				except Exception as e:
					logger.error(f"[ERROR] Could not click view button: {str(e)}")
					continue

				# Extract additional bid details
				bid_details = extract_bid_details(driver, bid_info)
				if not bid_details:
					continue

				# Download attachments
				print("\nProcessing attachments...")
				attachments = download_attachments(driver, None)  # Pass None initially
				
				if attachments:
					# Only create bid folder if we have attachments
					bid_folder = os.path.join(script_folder, bid_details['Solicitation Number'])
					os.makedirs(bid_folder, exist_ok=True)
					
					# Move files from download folder to bid folder
					downloads_folder = get_download_folder()
					processed_files = []
					
					for file in os.listdir(downloads_folder):
						if file.endswith('.crdownload') or file.endswith('.tmp'):
							continue
							
						file_path = os.path.join(downloads_folder, file)
						
						try:
							if file.lower().endswith('.zip'):
								print(f"Processing zip file: {file}")
								with zipfile.ZipFile(file_path, 'r') as zip_ref:
									for zip_info in zip_ref.infolist():
										if not zip_info.filename.endswith('/'):  # Skip directories
											filename = os.path.basename(zip_info.filename)  # Get just filename
											source = zip_ref.read(zip_info.filename)
											target_path = os.path.join(bid_folder, filename)
											with open(target_path, 'wb') as f:
												f.write(source)
											processed_files.append(filename)
								os.remove(file_path)  # Remove zip after extraction
							else:
								destination = os.path.join(bid_folder, file)
								if safe_move(file_path, destination):
									processed_files.append(file)
						except Exception as e:
							print(f"Error processing file {file}: {str(e)}")
							play_notification_sound()
					
					print(f"Downloaded attachments: {', '.join(processed_files)}")
					bid_details['Attachments'] = ', '.join(processed_files)
				else:
					print("No attachments downloaded")
					bid_details['Attachments'] = ''

				bids_data.append(bid_details)
				print(f"Successfully processed bid: {bid_details['Solicitation Number']}")

				# Update Excel file after each bid
				update_excel_file(script_folder, script_name, bids_data)

				# Return to listing page before processing next bid
				if not return_to_listing(driver):
					logger.error("[ERROR] Failed to return to listing page")
					driver.get(url)  # Fallback to loading URL if button click fails
					time.sleep(5)
					if not wait_for_page_load(driver):
						continue
					if not sort_bids_by_date(driver):
						continue

			except Exception as e:
				logger.error(f"[ERROR] Error processing bid: {str(e)}")
				print(f"Error processing bid: {str(e)}")
				continue

		# Final status
		if bids_data:
			print(f"\nProcessed {len(bids_data)} bids successfully")
			return True
		else:
			if found_recent_bids:
				logger.warning("[WARNING] No bid data collected despite finding recent bids")
				return False
			else:
				logger.info("[INFO] No bids found within the specified date range")
				return True

	except Exception as e:
		logger.error(f"[ERROR] Error in scrape_ionwave_site: {str(e)}")
		return False
	finally:
		if driver:
			try:
				driver.quit()
			except:
				pass

def extract_bid_details(driver, bid_info):
	"""Extract additional details from bid detail page"""
	try:
		# Initialize details with data we already have
		details = {
			'Posted Date': bid_info['formatted_posted_date'],
			'Response Date': bid_info['formatted_response_date'],
			'Notice Type': '',
			'Solicitation Number': bid_info['solicitation_number'],
			'Solicitation Title': bid_info['solicitation_title'],
			'Agency': bid_info['agency'],
			'Category': '',
			'Description': '',
			'Additional Summary': '',
			'Contracting Office Address': '',
			'Contact Information': '',
			'Bid Detail Page URL': driver.current_url,
			'Attachments': []
		}

		# Wait for details page to load
		time.sleep(3)

		# Notice Type
		try:
			notice_type = wait_for_fresh_element(
				driver,
				(By.ID, "ctl00_mainContent_lblType"),
				timeout=5
			)
			if notice_type:
				details['Notice Type'] = notice_type.text
		except:
			logger.warning("[WARNING] Notice Type not found")

		# Description
		try:
			description = wait_for_fresh_element(
				driver,
				(By.ID, "ctl00_mainContent_lblNotes"),
				timeout=5
			)
			if description:
				details['Description'] = description.text
		except:
			logger.warning("[WARNING] Description not found")

		# Contact Information
		try:
			name = wait_for_fresh_element(driver, (By.ID, "ctl00_mainContent_lblName"), timeout=5)
			phone = wait_for_fresh_element(driver, (By.ID, "ctl00_mainContent_lblPhone"), timeout=5)
			email = wait_for_fresh_element(driver, (By.ID, "ctl00_mainContent_lblEmail"), timeout=5)
			
			contact_info = []
			if name and name.text.strip():
				contact_info.append(f"Name: {name.text}")
			if phone and phone.text.strip():
				contact_info.append(f"Phone: {phone.text}")
			if email and email.text.strip():
				contact_info.append(f"Email: {email.text}")
			
			details['Contact Information'] = "\n".join(contact_info) if contact_info else ''
		except:
			logger.warning("[WARNING] Contact information not found")

		# Contracting Office Address
		try:
			address = wait_for_fresh_element(
				driver,
				(By.ID, "ctl00_mainContent_lblAddress"),
				timeout=5
			)
			if address:
				details['Contracting Office Address'] = address.text
		except:
			logger.warning("[WARNING] Address not found")

		return details

	except Exception as e:
		logger.error(f"[ERROR] Error in extract_bid_details: {str(e)}")
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
		
		if os.path.exists(in_progress_path):
			try:
				os.rename(in_progress_path, completed_path)
				logger.info(f"[SUCCESS] Renamed folder to {completed_path}")
			except Exception as e:
				logger.error(f"[ERROR] Failed to rename folder: {str(e)}")
	except Exception as e:
		logger.error(f"[ERROR] Error in complete_scraping: {str(e)}")

def scrape_bid_details(driver, bid_info, bid_folder):
	"""Scrape details for a single bid including attachments"""
	try:
		print(f"\nProcessing bid details for: {bid_info['solicitation_number']}")
		
		# Click view button for this bid
		view_button = WebDriverWait(driver, 10).until(
			EC.element_to_be_clickable((By.XPATH, 
				f"//tr[contains(@id,'rgBidList_ctl00__')]//td[2][contains(text(),'{bid_info['solicitation_number']}')]/..//td[1]/span[@class='flaticon-grid_View']"))
		)
		view_button.click()
		time.sleep(10)

		# Wait for bid details page to load
		WebDriverWait(driver, 10).until(
			EC.presence_of_element_located((By.ID, "ctl00_mainContent_lblType"))
		)
		time.sleep(5)

		# Extract bid details
		bid_details = extract_bid_details(driver, bid_info)
		if not bid_details:
			print("Failed to extract bid details")
			return None

		# Return to listing page
		print("\nReturning to main page...")
		return_button = WebDriverWait(driver, 10).until(
			EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'rtbText') and text()='Return to Listing']"))
		)
		return_button.click()
		time.sleep(5)

		return bid_details

	except Exception as e:
		print(f"Error in scrape_bid_details: {str(e)}")
		logger.error(f"[ERROR] Error scraping bid details: {str(e)}")
		play_notification_sound()
		return None

def get_download_folder():
	"""Get the download folder path for the current script"""
	yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
	script_name = os.path.splitext(os.path.basename(__file__))[0]
	base_folder = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), yesterday)
	in_progress_folder = os.path.join(base_folder, f"{script_name}_IN_PROGRESS")
	download_folder = os.path.join(in_progress_folder, script_name)
	os.makedirs(download_folder, exist_ok=True)
	return download_folder

def get_download_script():
	"""Return the JavaScript code for downloading attachments"""
	return """
	async function downloadAllAttachments() {
		const downloadedFiles = new Set(); // Track downloaded files to prevent duplicates
		const downloadedLinks = new Set(); // Track clicked links to prevent duplicates

		async function clickElements(xpathPattern) {
			let rowCount = 1;
			
			while (true) {
				const attachmentXPath = xpathPattern.replace('{row}', rowCount);
				const attachmentElement = document.evaluate(
					attachmentXPath, 
					document, 
					null, 
					XPathResult.FIRST_ORDERED_NODE_TYPE, 
					null
				).singleNodeValue;

				if (!attachmentElement) {
					break; // No more elements found with this pattern
				}

				const fileName = attachmentElement.textContent.trim();
				const linkHref = attachmentElement.getAttribute('href');
				const linkId = `${fileName}_${linkHref}`; // Create unique identifier for link

				// Only click if we haven't processed this link before
				if (!downloadedLinks.has(linkId)) {
					console.log('Found attachment:', fileName);
					try {
						attachmentElement.click();
						downloadedFiles.add(fileName);
						downloadedLinks.add(linkId);
						console.log(`Started download for: ${fileName}`);
						await new Promise(resolve => setTimeout(resolve, 5000));
					} catch (error) {
						console.error('Error clicking attachment:', error);
					}
				}
				rowCount++;
			}
		}

		const xpathPatterns = [
			'/html/body/form/div[3]/div[3]/div/table/tbody/tr[{row}]/td[1]/a',
			'/html/body/form/div[3]/div[4]/div/table/tbody/tr[{row}]/td[1]/a'
		];

		for (const pattern of xpathPatterns) {
			await clickElements(pattern);
		}

		return Array.from(downloadedFiles);
	}
	return await downloadAllAttachments();
	"""

def download_attachments(driver, bid_folder):
	"""Download all available attachments for a bid"""
	try:
		print("\n=== Starting Attachment Download Process ===")
		
		# Verify we're on the bid detail page
		try:
			WebDriverWait(driver, 10).until(
				EC.presence_of_element_located((By.ID, "ctl00_mainContent_lblType"))
			)
			print("✓ Confirmed on bid detail page")
		except:
			print("Not on bid detail page - attachment download may fail")
			play_notification_sound()
			return []

		# Look for attachment links
		attachment_links = driver.find_elements(By.XPATH, "//div[contains(@id, 'rgBidDocuments') or contains(@id, 'rgBidAttachments')]//a[contains(@class, 'procLink')]")
		if not attachment_links:
			print("No attachment links found")
			return []
			
		print(f"✓ Found {len(attachment_links)} attachment links:")
		downloaded_files = set()  # Track which files we've processed
		
		# Process each attachment link
		for idx, link in enumerate(attachment_links, 1):
			filename = link.text.strip()
			print(f"  - Attachment {idx}: {filename}")
			
			if filename not in downloaded_files:
				try:
					link.click()
					downloaded_files.add(filename)
					print(f"Started download for: {filename}")
					time.sleep(5)  # Wait between downloads
				except Exception as e:
					print(f"Error clicking link for {filename}: {str(e)}")
					continue
		
		if downloaded_files:
			print("\nProcessing downloads...")
			time.sleep(30)  # Initial wait for downloads to start
			
			attachments = []
			downloads_folder = get_download_folder()
			
			# Wait for and process each downloaded file
			for filename in downloaded_files:
				print(f"\nProcessing: {filename}")
				download_complete = False
				start_time = time.time()
				max_wait_time = 300  # 5 minutes per file
				
				while not download_complete and (time.time() - start_time) < max_wait_time:
					# Check for active downloads
					downloading_files = [f for f in os.listdir(downloads_folder) 
									  if f.endswith('.crdownload') or f.endswith('.tmp')]
					if downloading_files:
						print(f"⏳ Still downloading: {downloading_files}")
						time.sleep(10)
						continue
					
					# Look for completed downloads
					for file in os.listdir(downloads_folder):
						if file.endswith('.crdownload') or file.endswith('.tmp'):
							continue
							
						file_path = os.path.join(downloads_folder, file)
						
						if time.time() - os.path.getctime(file_path) < max_wait_time:
							print(f"✓ Found downloaded file: {file}")
							attachments.append(file)
							download_complete = True
							break
					
					if not download_complete:
						print(f"⏳ Waiting for {filename}...")
						time.sleep(10)
				
				if not download_complete:
					print(f"Download failed for {filename}")
					play_notification_sound()
			
			print(f"\n=== Download Process Complete ===")
			print(f"Total attachments: {len(attachments)}")
			if attachments:
				print("Files:")
				for idx, att in enumerate(attachments, 1):
					print(f"  {idx}. {att}")
			return list(set(attachments))  # Remove any duplicates
		
		print("\nNo files queued for download")
		return []
		
	except Exception as e:
		print(f"\nError in download_attachments: {str(e)}")
		logger.error(f"[ERROR] Error in download_attachments: {str(e)}")
		play_notification_sound()
		return []

def wait_for_page_load(driver, timeout=30):
	"""Wait for page to fully load"""
	try:
		# Wait for document ready state
		WebDriverWait(driver, timeout).until(
			lambda d: d.execute_script('return document.readyState') == 'complete'
		)
		
		# Wait for any AJAX calls to complete
		time.sleep(5)
		
		# Check if page has error message or is empty
		try:
			error_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Error') or contains(text(), 'error')]")
			if error_elements:
				logger.info("[INFO] Page contains error message - may be temporarily unavailable")
				return False
		except:
			pass

		# Try to find bid list
		try:
			bid_elements = driver.find_elements(By.XPATH, "//tr[contains(@id,'rgBidList_ctl00__')]")
			if not bid_elements:
				logger.info("[INFO] No bid listings found on page")
				return False
		except:
			logger.info("[INFO] Could not locate bid list - page may be empty")
			return False
		
		return True
	except Exception as e:
		logger.info(f"[INFO] Page load issue: {str(e)}")
		return False

def main():
	"""Main function to run the scraper"""
	try:
		# Step 1: Start the Process
		print_start_message()
		
		# Setup logging and get working directory
		base_folder = get_base_folder()
		
		# Parse command line arguments
		parser = argparse.ArgumentParser(description='Scrape Ionwave bids')
		parser.add_argument('--days', type=int, default=1,
						  help='Number of days back to scrape (default: 1)')
		args = parser.parse_args()
		
		print(f"\nConfiguration:")
		print(f"- Days to look back: {args.days}")
		print(f"- Working folder: {base_folder}")
		logger.info(f"[CONFIG] Days to look back: {args.days}")

		# Ensure the working folder exists
		if not os.path.exists(base_folder):
			os.makedirs(base_folder)
			logger.info(f"[INFO] Created working folder: {base_folder}")

		# URLs in specified order
		urls = [
			"https://miamiu-ohiousourcing.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://lexingtoncounty.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://washco-md.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://lauderhill.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://mcallenisd.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://humbleisd.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://iowadotebid.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://geus.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://cfisdbid.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://cbiabids.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://moline.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://planotx.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://sanjac.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://brazosbid.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://newcaneyisd.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://ura.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://ccisd.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://davenport.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://arlington-tx.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://carrolltonbids.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://jocogov.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://dentontx.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://unlebid.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://blounttn.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://galenaparkisd.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://southlakecarroll.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://pisd.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://stillwater.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://flower-mound.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://garlandtx.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://citycrbids.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://lbwl.ionwave.net/SourcingEvents.aspx?SourceType=1",
			"https://wylieisd.ionwave.net/CurrentSourcingEvents.aspx",
			"https://lcpscm.ionwave.net/CurrentSourcingEvents.aspx"
		]

		# Remove duplicates while preserving order
		seen = set()
		urls = [x for x in urls if not (x in seen or seen.add(x))]
		
		total_urls = len(urls)
		processed_urls = 0
		successful_urls = 0
		failed_urls = []

		print(f"\nTotal URLs to process: {total_urls}")
		logger.info(f"[INFO] Total URLs to process: {total_urls}")

		excel_path = os.path.join(base_folder, '18_Ionwave.xlsx')
		logger.info(f"[INFO] Excel file will be saved as: {excel_path}")

		for url in urls:
			try:
				processed_urls += 1
				print(f"\nProcessing URL {processed_urls}/{total_urls}")
				print(f"URL: {url}")
				logger.info(f"[PROGRESS] Processing URL {processed_urls}/{total_urls}: {url}")
				
				success = scrape_ionwave_site(url, args.days)
				
				if success:
					successful_urls += 1
					print(f"Status: Success")
					logger.info(f"[SUCCESS] Successfully processed URL: {url}")
				else:
					failed_urls.append(url)
					print(f"Status: Failed")
					logger.error(f"[ERROR] Failed to process URL: {url}")
					play_notification_sound()
					user_input = input("Press Enter to continue to next URL, or 'q' to quit: ")
					if user_input.lower() == 'q':
						break
				
				print(f"Progress: {processed_urls}/{total_urls} URLs processed")
				time.sleep(2)

			except Exception as e:
				logger.info(f"[INFO] Skipping URL {url}: {str(e)}")
				print(f"Status: Skipped")
				time.sleep(2)
				continue

		# Print summary
		print("\n" + "="*50)
		print("Execution Summary:")
		print(f"Total URLs processed: {processed_urls}/{total_urls}")
		print(f"Successful: {successful_urls}")
		print(f"Failed: {len(failed_urls)}")
		
		if failed_urls:
			print("\nFailed URLs:")
			for url in failed_urls:
				print(f"- {url}")

		logger.info("\n" + "="*50)
		logger.info("[SUMMARY] Execution Summary")
		logger.info(f"[SUMMARY] Working folder: {base_folder}")
		logger.info(f"[SUMMARY] Excel file: {excel_path}")
		logger.info(f"[SUMMARY] Total URLs processed: {processed_urls}/{total_urls}")
		logger.info(f"[SUMMARY] Successful URLs: {successful_urls}")
		logger.info(f"[SUMMARY] Failed URLs: {len(failed_urls)}")
		
		if failed_urls:
			logger.info("[SUMMARY] Failed URLs:")
			for url in failed_urls:
				logger.info(f"[SUMMARY] - {url}")

		notify_completion()

	except Exception as e:
		error_msg = f"[FATAL] Fatal error in main execution: {str(e)}"
		logger.error(error_msg)
		print(f"\nFatal Error: {str(e)}")
		play_notification_sound()
		input("Press Enter to exit...")
		sys.exit(1)

if __name__ == "__main__":
	main()