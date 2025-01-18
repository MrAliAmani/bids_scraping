import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import functools

# Override the print function to always flush
print = functools.partial(print, flush=True)

import time
import argparse
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver import ActionChains
from fake_useragent import UserAgent
from selenium_stealth import stealth
import pandas as pd
from utils.utils import safe_move, play_notification_sound
import signal
from pathlib import Path
import re


class CalProcureScraper:
    def __init__(self, days_back=2):
        # First define all instance variables
        self.base_url = "https://caleprocure.ca.gov/pages/Events-BS3/event-search.aspx"
        self.days_back = days_back
        # Use yesterday's date for folder name
        self.today = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        self.bids_data = []
        self.current_bid_number = None
        self.max_retries = 3
        self.page_load_timeout = 30

        # Updated business units mapping based on the provided codes
        self.business_units = {
            "Statewide Business Unit": "0000",
            "State of California Emergency": "00001",
            "Legislative/Judicial/Executive": "0010",
            "Legislature": "0100",
            "Senate": "0110",
            "Assembly": "0120",
            "Legislative Joint Expenses": "0130",
            "Legislative Counsel Bureau": "0160",
            "Judicial Branch": "0250",
            "SUPREME COURT": "0260",
            "JUDICIAL COUNCIL": "0270",
            "Comm. on Judicial Performance": "0280",
            "Habeas Resource Center": "0290",
            "DISTRICT COURTS OF APPEAL": "0300",
            "1st DISTRICT COURT OF APPEAL": "0310",
            "2nd DISTRICT COURT OF APPEAL": "0320",
            "3rd DISTRICT COURT OF APPEAL": "0330",
            "Governor's Office": "0500",
            "California Technology Agency": "0502",
            "Business & Economic Developmnt": "0509",
            "Sec. for State & Consumer Svcs": "0510",
            "Sec., Gov't Operations Agency": "0511",
            "Bus, Consmer Svcs & Hsng Secty": "0515",
            "Sec Business Trans & Housing": "0520",
            "Sec., Transportation Agency": "0521",
            "Sec., Health & Human Services": "0530",
            "Ofc Technology and Solutions I": "0531",
            "Sec., Natural Resources": "0540",
            "Inspector General Office": "0552",
            "Sec., Environment Protection": "0555",
            "Secretary for Education": "0558",
            "Sec., Labor/Workforce Develop": "0559",
            "Office of Emergency Services": "0690",
            "Department of Justice": "0820",
            "State Controller": "0840",
            "Department of Insurance": "0845",
            "CA State Lottery Commission": "0850",
            "CA Gambling Control Commission": "0855",
            "State Board of Equalization": "0860",
            "Office of Tax Appeals": "0870",
            "Secretary of State": "0890",
            "Citizens Redistricting Comm": "0911",
            "State Treasurer": "0950",
            "Scholarshare Investment Board": "0954",
            "Debt & Investment Advisory Com": "0956",
            "Debt Limit Allocation Commitee": "0959",
            "Industrial Develop Fin Comm": "0965",
            "Tax Credit Allocation Commitee": "0968",
            "Alternative Energy & Adv Trans": "0971",
            "Pollution Control Fin Auth": "0974",
            "Health Facilities Fin Auth": "0977",
            "California ABLE Act Board": "0981",
            "CalSavers Retirement Savings B": "0984",
            "CA School Finance Authority": "0985",
            "Educational Facilities Auth": "0989",
            "Business, Consumer Srvs & Hous": "1000",
            "Sec., Biz, Con Srvs, & Housing": "1015",
            "No Subagency - DO NOT USE": "1020",
            "Sec., State & Consumer Service": "1030",
            "Cannabis Control Appeals Panel": "1045",
            "CA Science Center": "1100",
            "CA African American Museum": "1105",
            "Consumer Affairs-Reg Boards": "1110",
            "Department of Consumer Affairs": "1111",
            "Department of Cannabis Control": "1115",
            "A. E. Alquist Seismic Safety": "1690",
            "Civil Rights Department": "1700",
            "Dept of Finan Protec and Innov": "1701",
            "California Privacy Protection": "1703",
            "Fair Employment & Housing Comm": "1705",
            "Franchise Tax Board": "1730",
            "Horse Racing Board": "1750",
            "Department of General Services": "7760",  # Updated from 1760 to 7760
            "Victim Comp & Govt Claims Bd": "1870",
            "State Personnel Board": "1880",
            "Public Employees' Retirement": "1900",
            "State Teachers' Retirement": "1920",
            "GO Bonds - Debt Service  - BCH": "1996",
            "Business, Transport & Housing": "2000",
            "Business & Housing": "2010",
            "Transportation": "2020",
            "Sec., Business, Trans & Housng": "2030",
            "Dept. Alcoholic Beverage Cntrl": "2100",
            "Alcoholic Beverage Cntl Appeal": "2120",
            "Dept. of Financial Institution": "2150",
            "Department of Corporations": "2180",
            "St Asst Fd Enterprise,Bus & In": "2222",
            "Housing & Community Developmnt": "2240",
            "CA Housing Finance Agency": "2245",
            "Office Real Estate Appraisers": "2310",
            "Department of Real Estate": "2320",
            "Dept of Managed Health Care": "2400",
            "Transportation": "2500",
            "Sec. for Transportation Agency": "2521",
            "CA Transportation Commission": "2600",
            "State Transit Assistance": "2640",
            "Department of Transportation": "2660",
            "High Speed Rail Authority": "2665",
            "High-Speed Rail Auth Ofc Inspe": "2667",
            "Board of Pilot Commissioners": "2670",
            "Office of Traffic Safety": "2700",
            "Dept of the CA Highway Patrol": "2720",
            "Department of Motor Vehicles": "2740",
            "GO Bonds-Transportation": "2830",
            "Natural Resources": "3000",
            "Sec., Natural Resources Agency": "3030",
            "Exposition Park": "3100",
            "CA African American Museum": "3105",
            "Special Resources Programs": "3110",
            "CA Tahoe Conservancy": "3125",
            "Geothermal ResourcesDevProgram": "3180",
            "Environmental Protection Pgm": "3210",
            "CA Conservation Corps": "3340",
            "Office of Energy Infrastructur": "3355",
            "Energy Resources Conservation": "3360",
            "Renew Res Protect Pgm": "3370",
            "Colorado River Board of CA": "3460",
            "Department of Conservation": "3480",
            "Resources Recycling & Recovery": "3500",
            "CAL FIRE": "3540",
            "State Lands Commission": "3560",
            "Department of Fish & Wildlife": "3600",
            "Wildlife Conservation Board": "3640",
            "Dept of Boating & Waterways": "3680",
            "CA Coastal Commission": "3720",
            "State Coastal Conservancy": "3760",
            "Native American Heritage Comm": "3780",
            "Dept of Parks & Recreation": "3790",
            "Santa Monica Mtns Conservancy": "3810",
            "Salton Sea Conservancy": "3815",
            "SF Bay Conservation Commission": "3820",
            "San Gabriel & Lower LA Rivers": "3825",
            "San Joaquin River Conservancy": "3830",
            "Baldwin Hills and Urban Waters": "3835",
            "Delta Protection Commission": "3840",
            "San Diego River Conservancy": "3845",
            "Coachella Valley Mtns Conser": "3850",
            "Sierra Nevada Conservancy": "3855",
            "Department of Water Resources": "3860",
            "Sacramento-San Joaquin Delta": "3875",
            "GO Bonds Resources": "3882",
            "Delta Stewardship Council": "3885",
            "Environmental Protection": "3890",
            "Sec., Environmental Protectio": "3895",
            "State Air Resources Board": "3900",
            "Dept of Pesticide Regulation": "3930",
            "State Water Resources Control": "3940",
            "Dept. Toxic Substances Control": "3960",
            "Resources Recycling & Recovery": "3970",
            "Env'l Health Hazard Assessment": "3980",
            "GO Bonds Env Protect": "3996",
            "Health & Human Services": "4000",
            "Sec., Health & Human Srvs Agy": "4020",
            "Developmental Disabilities": "4100",
            "Emergency Medical Service Auth": "4120",
            "Health Care Access and Informa": "4140",
            "Dept of Managed Health Care": "4150",
            "California Department of Aging": "4170",
            "Commission on Aging": "4180",
            "CA Senior Legislature": "4185",
            "Dept. Alcohol & Drug Programs": "4200",
            "First 5 California": "4250",
            "State Dept Hlth Care Services": "4260",
            "Department of Public Health": "4265",
            "Medical Assistance Commission": "4270",
            "Managed Risk Medical Insurance": "4280",
            "Dept of Developmental Services": "4300",
            "Developmental Services - HQ": "4310",
            "State Hospitals": "4320",
            "Department of State Hospitals": "4440",
            "State Hospitals Sacramento": "4450",
            "State Hospital - Atascadero": "4470",
            "State Hospital Metropolitan": "4490",
            "State Hospital Napa": "4500",
            "State Hospital Patton": "4510",
            "State Hospital Stockton": "4520",
            "State Hospital Vacaville": "4530",
            "State Hospital Coalinga": "4540",
            "State Hospital Salinas": "4550",
            "Mental Health Srvcs Oversight": "4560",
            "Community Srvcs & Development": "4700",
            "CA Health Benefit Exchange": "4800",
            "Department of Rehabilitation": "5160",
            "Dept of Youth and Community Re": "5165",
            "Independent Living Council": "5170",
            "Dept of Child Support Services": "5175",
            "Department of Social Services": "5180",
            "State - Local Realignment 1991": "5195",
            "State - Local Realignment 2011": "5196",
            "GO Bonds -HHS": "5206",
            "Corrections & Rehabilitation": "5210",
            "Dept of Corrections & Rehab": "5225",
            "State & Community Corrections": "5227",
            "Prison Industry Authority": "5420",
            "Department of Education": "6100",
            "CA State Library": "6120",
            "Education Audit Appeals Panel": "6125",
            "CA School for the Blind": "6200",
            "School for the Deaf-Fremont": "6240",
            "School for the Deaf-Riverside": "6250",
            "State Summer School for Arts": "6255",
            "Teacher Credentialing Comm": "6360",
            "University of California": "6440",
            "Institute for Regenerative Med": "6445",
            "UC Office of the President": "6491",
            "UC Berkeley": "6500",
            "UC Davis": "6510",
            "UC Davis Medical Center": "6511",
            "UC Irvine": "6520",
            "UC Irvine Med Center": "6521",
            "UCLA": "6530",
            "UC San Francisco": "6560",
            "UCSF Medical Center": "6561",
            "UC Santa Cruz": "6580",
            "UC Merced": "6590",
            "College of the Law, San Fran": "6600",
            "Cal State University": "6610",
            "CSU Statewide Programs": "6620",
            "CSU Systemwide Offices": "6630",
            "CSU Campuses": "6640",
            "CSU Health Ben for Ret Annuit": "6645",
            "CSU, Bakersfield": "6650",
            "CSU, San Bernardino": "6660",
            "CSU, Stanislaus": "6670",
            "CSU, Dominguez Hills": "6690",
        }

        # Then call setup methods
        self.setup_folders()
        self.setup_driver()
        self.setup_signal_handler()

    def setup_folders(self):
        """Setup necessary folders for downloads and bid data"""
        try:
            # Create main folder with yesterday's date
            self.main_folder = os.path.join(os.getcwd(), self.today)
            os.makedirs(self.main_folder, exist_ok=True)
            print(f"✅ Created main folder: {self.main_folder}")

            # Create script-specific folder with IN_PROGRESS suffix
            self.script_folder_name = "16_CalProcure_IN_PROGRESS"
            self.script_folder = os.path.join(self.main_folder, self.script_folder_name)
            os.makedirs(self.script_folder, exist_ok=True)
            print(f"✅ Created script folder: {self.script_folder}")

            # Create temp download folder inside script folder (not main folder)
            self.temp_download_folder = os.path.join(
                self.script_folder, "16_CalProcure"
            )
            os.makedirs(self.temp_download_folder, exist_ok=True)
            print(f"✅ Created temp download folder: {self.temp_download_folder}")

        except Exception as e:
            print(f"❌ Error creating folders: {str(e)}")
            raise

    def setup_driver(self):
        """Configure and setup Chrome WebDriver with anti-bot measures"""
        options = webdriver.ChromeOptions()
        ua = UserAgent()
        options.add_argument(f"user-agent={ua.random}")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--start-maximized")  # Start maximized
        options.add_argument("--ignore-certificate-errors")  # Ignore SSL errors
        options.add_argument("--disable-extensions")  # Disable extensions
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": self.temp_download_folder,
                "download.prompt_for_download": False,
                "plugins.always_open_pdf_externally": True,
                "profile.default_content_setting_values.notifications": 2,
                "profile.default_content_settings.popups": 0,
            },
        )

        self.driver = webdriver.Chrome(options=options)
        # Increase timeouts
        self.driver.set_page_load_timeout(90)  # Increased from 60 to 90 seconds
        self.driver.implicitly_wait(30)  # Increased from 20 to 30 seconds

        stealth(
            self.driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
        )

    def wait_for_element(self, by, selector, timeout=30):
        """Wait for element to be present and visible"""
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
            # Additional wait for element to be visible and interactable
            WebDriverWait(self.driver, timeout).until(
                EC.visibility_of_element_located((by, selector))
            )
            return element
        except TimeoutException:
            print(f"❌ Timeout waiting for element: {selector}")
            return None

    def apply_filters(self):
        """Apply search filters"""
        print("🔍 Applying search filters...")

        try:
            # Wait for and click Advanced Search button
            advanced_search = WebDriverWait(self.driver, 25).until(
                EC.element_to_be_clickable((By.ID, "RESP_INQA_WK_AUC_SRCH_ADVBSC_PB"))
            )
            self.driver.execute_script("arguments[0].click();", advanced_search)
            print("✅ Clicked Advanced Search button")

            # Wait for date input to be present and visible
            date_input = WebDriverWait(self.driver, 25).until(
                EC.presence_of_element_located(
                    (By.ID, "RESP_INQA_WK_AUC_FROM_START_DT")
                )
            )
            WebDriverWait(self.driver, 20).until(EC.visibility_of(date_input))

            # Set start date using JavaScript with retry mechanism
            start_date = (datetime.now() - timedelta(days=self.days_back)).strftime(
                "%m/%d/%Y"
            )
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    # Clear existing value first
                    self.driver.execute_script("arguments[0].value = '';", date_input)
                    time.sleep(5)  # Increased wait after clearing

                    # Set new value
                    self.driver.execute_script(
                        f"arguments[0].value = '{start_date}';", date_input
                    )
                    time.sleep(7)  # Wait after setting value

                    # Verify the value was set
                    actual_value = self.driver.execute_script(
                        "return arguments[0].value;", date_input
                    )

                    if actual_value == start_date:
                        print(f"✅ Set start date to {start_date}")
                        break
                    else:
                        raise Exception("Date value not set correctly")

                except Exception as e:
                    if attempt == max_retries - 1:
                        raise
                    print(f"⚠️ Attempt {attempt + 1} failed to set date: {str(e)}")
                    time.sleep(3)  # Increased wait between retries

            # Click Search button and wait for results
            for attempt in range(max_retries):
                try:
                    search_button = WebDriverWait(self.driver, 20).until(
                        EC.element_to_be_clickable(
                            (By.ID, "RESP_INQA_WK_INQ_AUC_GO_PB")
                        )
                    )
                    self.driver.execute_script("arguments[0].click();", search_button)
                    print("✅ Clicked Search button")

                    # Wait for results to load
                    print("⏳ Waiting for search results...")
                    time.sleep(20)  # Initial wait

                    # Wait for table to be present
                    table = WebDriverWait(self.driver, 30).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "table tbody tr")
                        )
                    )

                    # Wait for specific bid data cells to be present
                    WebDriverWait(self.driver, 30).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "td[data-if-label='tdEventId']")
                        )
                    )
                    WebDriverWait(self.driver, 30).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "td[data-if-label='tdDepName']")
                        )
                    )

                    # Additional wait for data to fully load
                    time.sleep(5)

                    # Get bid rows
                    bid_rows = self.driver.find_elements(
                        By.CSS_SELECTOR, "tr[id^='trRESP_INQA_HD_VW_GR$0_row']"
                    )
                    if not bid_rows:
                        raise Exception("No bid rows found after search")

                    print(f"✅ Found {len(bid_rows)} bid results")
                    return True

                except Exception as e:
                    if attempt == max_retries - 1:
                        print(
                            f"❌ Failed to get search results after {max_retries} attempts"
                        )
                        raise
                    print(
                        f"⚠️ Attempt {attempt + 1} failed to get search results: {str(e)}"
                    )
                    time.sleep(20)  # Increased wait between retries

        except Exception as e:
            print(f"❌ Error applying filters: {str(e)}")
            raise

    def parse_date(self, date_str):
        """Convert date string to YYYY-MM-DD format"""
        try:
            date_str = date_str.split()[0]  # Remove time part
            date_obj = datetime.strptime(date_str, "%m/%d/%Y")
            return date_obj.strftime("%Y-%m-%d")
        except:
            return ""

    def extract_bid_details(self, row_num, bid_info):
        """Extract details from bid page"""
        try:
            print("\n📄 Extracting bid details...")
            bid_data = {
                "SL No": row_num,
                "Posted Date": "",
                "Response Date": bid_info["response_date"],  # From main page
                "Notice Type": "",
                "Solicitation Number": bid_info["bid_number"],  # From main page
                "Solicitation Title": bid_info["title"],  # From main page
                "Agency": "",
                "Category": "",
                "Description": "",
                "Additional Summary": "",
                "Contracting Office Address": "",
                "Contact Information": "",
                "Bid Detail Page URL": self.driver.current_url,
                "Attachments": "",
            }

            # Posted Date - Format: 11/04/2024 4:04PM PST -> YYYY-MM-DD
            try:
                posted_date = self.driver.find_element(
                    By.CSS_SELECTOR, "span[data-if-label='eventStartDate']"
                ).text.strip()
                bid_data["Posted Date"] = self.parse_date(posted_date)
                print(f"📅 Posted Date: {bid_data['Posted Date']}")
            except:
                print("⚠️ Could not extract Posted Date")

            # Notice Type - Concatenate Format/Type with "/"
            try:
                format1 = self.driver.find_element(
                    By.CSS_SELECTOR, "span[data-if-label='format1']"
                ).text.strip()
                format2 = self.driver.find_element(
                    By.CSS_SELECTOR, "span[data-if-label='format2']"
                ).text.strip()
                bid_data["Notice Type"] = f"{format1}/{format2}"
                print(f"📋 Notice Type: {bid_data['Notice Type']}")
            except:
                print("⚠️ Could not extract Notice Type")

            # Agency (Department)
            try:
                agency = self.driver.find_element(
                    By.CSS_SELECTOR, "span[data-if-label='dept']"
                ).text.strip()
                bid_data["Agency"] = agency
                print(f"🏢 Agency: {bid_data['Agency']}")
            except:
                print("⚠️ Could not extract Agency")

            # Category (UNSPSC Codes) - Only first two categories
            try:
                categories = []
                category_rows = self.driver.find_elements(
                    By.CSS_SELECTOR, "#unspscTable tbody tr"
                )[
                    :2
                ]  # Only first two

                for row in category_rows:
                    code = row.find_element(By.CSS_SELECTOR, "td.bold").text.strip()
                    desc = row.find_element(
                        By.CSS_SELECTOR, "td:nth-child(2)"
                    ).text.strip()
                    categories.append(f"{code} - {desc}")

                bid_data["Category"] = "; ".join(categories)
                print(f"🏷️ Categories: {bid_data['Category']}")
            except:
                print("⚠️ Could not extract Categories")

            # Description
            try:
                description = self.driver.find_element(
                    By.CSS_SELECTOR, "div[data-if-label='descriptiondetails']"
                ).text.strip()
                bid_data["Description"] = description
                print(f"📝 Description found: {len(description)} characters")
            except:
                print("⚠️ Could not extract Description")

            # Contact Information
            try:
                contact_info = []

                # Name
                name = self.driver.find_element(
                    By.CSS_SELECTOR, "strong[data-if-label='contactName']"
                ).text.strip()
                contact_info.append(f"Name: {name}")

                # Title (always empty with label)
                contact_info.append("Title: ")

                # Phone
                phone = self.driver.find_element(
                    By.CSS_SELECTOR, "span[data-if-label='phoneText']"
                ).text.strip()
                if phone:
                    contact_info.append(f"Phone: {phone}")

                # Email
                email = self.driver.find_element(
                    By.CSS_SELECTOR, "a[data-if-label='emailAnchor']"
                ).text.strip()
                if email:
                    contact_info.append(f"Email: {email}")

                bid_data["Contact Information"] = "\n".join(contact_info)
                print(f"👤 Contact information extracted")
            except:
                print("⚠️ Could not extract Contact Information")

            # Additional Summary (Pre Bid Conference)
            try:
                conference_info = []
                conference_spans = self.driver.find_elements(
                    By.CSS_SELECTOR, "span[data-if-container-label='conferenceRow']"
                )

                for i in range(0, len(conference_spans), 2):
                    label = conference_spans[i].text.strip().rstrip(":")
                    value = (
                        conference_spans[i + 1].text.strip()
                        if i + 1 < len(conference_spans)
                        else ""
                    )
                    if label and not value.isspace():
                        conference_info.append(f"{label}: {value}")

                bid_data["Additional Summary"] = "\n".join(conference_info)
                print(f"📑 Pre-bid conference information extracted")
            except:
                print("⚠️ Could not extract Pre-bid Conference information")

            print("✅ Bid details extraction completed")
            return bid_data

        except Exception as e:
            print(f"❌ Error extracting bid details: {str(e)}")
            play_notification_sound()
            return None

    def get_business_unit_code(self, department_name):
        """Get business unit code from department name using mapping"""
        try:
            # Clean up department name
            dept_name = department_name.strip()
            print(f"🔍 Looking up business unit for: {dept_name}")

            # Try exact match first
            if dept_name in self.business_units:
                code = self.business_units[dept_name]
                print(f"✅ Found exact match: {code}")
                return code

            # Try case-insensitive match
            dept_lower = dept_name.lower()
            for name, code in self.business_units.items():
                if name.lower() == dept_lower:
                    print(f"✅ Found case-insensitive match: {code}")
                    return code

            # Try partial match
            for name, code in self.business_units.items():
                name_lower = name.lower()
                if name_lower in dept_lower or dept_lower in name_lower:
                    print(f"✅ Found partial match with {name}: {code}")
                    return code

            # If no match found, log warning and return default
            print(f"⚠️ No business unit match found for: {dept_name}")
            print("⚠️ Using default business unit code: 2660")
            return "2660"

        except Exception as e:
            print(f"❌ Error getting business unit code: {str(e)}")
            return "2660"

    def download_attachments(self, solicitation_number):
        """Download all bid attachments"""
        self.current_bid_number = solicitation_number
        try:
            # Ensure download folder exists
            if not os.path.exists(self.temp_download_folder):
                print(f"⚠️ Recreating download folder: {self.temp_download_folder}")
                os.makedirs(self.temp_download_folder, exist_ok=True)

            # Create bid-specific folder
            bid_folder = os.path.join(self.script_folder, solicitation_number)
            os.makedirs(bid_folder, exist_ok=True)
            print(f"✅ Created bid folder: {bid_folder}")

            # Click View Event Package
            print("🔍 Looking for View Event Package button...")
            try:
                package_btn = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located(
                        (By.ID, "RESP_INQ_DL0_WK_AUC_DOWNLOAD_PB")
                    )
                )

                self.driver.execute_script(
                    "arguments[0].scrollIntoView(true);", package_btn
                )
                time.sleep(5)

                self.driver.execute_script("arguments[0].click();", package_btn)
                print("✅ Clicked View Event Package button")
                time.sleep(10)

            except Exception as e:
                print(f"❌ Error clicking View Event Package button: {str(e)}")
                return ""

            # Check for "No attachments found"
            try:
                no_attachments = self.driver.find_element(
                    By.CSS_SELECTOR, "#noAttachmentsText > span"
                )
                if no_attachments.is_displayed():
                    print("ℹ️ No attachments found for this bid")
                    return ""
            except NoSuchElementException:
                pass

            # Wait for attachments table
            print("⏳ Waiting for attachments table...")
            try:
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
                )
                time.sleep(8)
            except TimeoutException:
                print("⚠️ No attachments table found")
                return ""

            moved_files = []
            row_num = 2  # Start with row 2
            max_attempts = 10  # Maximum number of rows to try

            # Process attachments using XPath
            while row_num <= max_attempts:
                try:
                    # XPath for the download button in current row
                    button_xpath = f"/html/body/div[8]/div/div/div[2]/div[2]/form/div[5]/div/div/div[3]/table/tbody/tr[{row_num}]/td[3]/button[2]"

                    # Try to find the button
                    try:
                        download_button = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, button_xpath))
                        )
                    except TimeoutException:
                        print(f"No more download buttons found after row {row_num-1}")
                        break

                    print(f"\n📥 Processing attachment at row {row_num}")

                    # Scroll button into view and click
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView(true);", download_button
                    )
                    time.sleep(5)

                    print("🖱️ Clicking download button...")
                    self.driver.execute_script("arguments[0].click();", download_button)
                    time.sleep(8)

                    try:
                        print("⏳ Waiting for final download button...")
                        final_download = WebDriverWait(self.driver, 30).until(
                            EC.element_to_be_clickable((By.ID, "downloadButton"))
                        )

                        # Additional wait for popup to fully load
                        time.sleep(10)

                        print("🖱️ Clicking final download button...")
                        self.driver.execute_script(
                            "arguments[0].click();", final_download
                        )
                        time.sleep(8)

                        # Wait for download
                        print("⏳ Waiting for download to complete...")
                        start_time = time.time()
                        download_complete = False

                        while not download_complete and (time.time() - start_time) < 60:
                            # Check for completed downloads
                            for filename in os.listdir(self.temp_download_folder):
                                if not filename.endswith(
                                    ".crdownload"
                                ) and not filename.endswith(".tmp"):
                                    source = os.path.join(
                                        self.temp_download_folder, filename
                                    )
                                    destination = os.path.join(
                                        bid_folder, filename
                                    )  # Move to bid-specific folder

                                    try:
                                        if safe_move(source, destination):
                                            moved_files.append(filename)
                                            print(
                                                f"✅ Moved: {filename} to bid folder {solicitation_number}"
                                            )
                                            download_complete = True
                                            break
                                    except Exception as e:
                                        print(
                                            f"⚠️ Error moving file {filename}: {str(e)}"
                                        )

                        if not download_complete:
                            time.sleep(2)

                    except Exception as e:
                        print(f"⚠️ Error with final download button: {str(e)}")

                    row_num += 1  # Move to next row

                except Exception as e:
                    print(f"⚠️ Error processing row {row_num}: {str(e)}")
                    row_num += 1  # Try next row even if current fails
                    continue

            if moved_files:
                print(
                    f"\n✅ Successfully downloaded and moved {len(moved_files)} attachments"
                )
                return ", ".join(moved_files)
            else:
                print("\nℹ️ No attachments were downloaded")
                return ""

        except Exception as e:
            print(f"❌ Error downloading attachments: {str(e)}")
            play_notification_sound()
            return ""

    def load_page_with_retry(self, url, max_retries=None):
        """Load page with retry mechanism"""
        if max_retries is None:
            max_retries = self.max_retries

        for attempt in range(max_retries):
            try:
                print(f"📄 Loading page attempt {attempt + 1}/{max_retries}...")

                # Clear cookies and cache before each attempt
                self.driver.delete_all_cookies()

                # Try to load the page in steps
                self.driver.get("about:blank")  # Clear current page
                time.sleep(2)  # Short pause

                # Load the actual URL
                self.driver.get(url)

                # Final check for error page after retries
                if "error.aspx" in self.driver.current_url:
                    print("❌ Site is unavailable, skipping this bid.")
                    return None

                # Wait for initial page load
                WebDriverWait(self.driver, 30).until(
                    lambda driver: driver.execute_script("return document.readyState")
                    == "complete"
                )

                # Additional checks for page load
                try:
                    # Wait for body to be present
                    WebDriverWait(self.driver, 20).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )

                    # Wait for any loading spinners to disappear (adjust selector as needed)
                    spinners = self.driver.find_elements(
                        By.CSS_SELECTOR, ".loading-spinner, .loader"
                    )
                    if spinners:
                        WebDriverWait(self.driver, 20).until(
                            EC.invisibility_of_elements_located(
                                (By.CSS_SELECTOR, ".loading-spinner, .loader")
                            )
                        )
                except TimeoutException:
                    print("⚠️ Some elements didn't load, but continuing...")

                # Verify we're on the correct page
                if "caleprocure.ca.gov" in self.driver.current_url:
                    # Additional wait for dynamic content
                    time.sleep(5)
                    print("✅ Page loaded successfully")
                    return True
                else:
                    raise Exception("Landed on incorrect page")

            except Exception as e:
                print(f"⚠️ Attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_retries - 1:
                    print("🔄 Retrying after longer delay...")
                    time.sleep(10)  # Increased delay between retries
                    continue
                else:
                    print("❌ All attempts failed")
                    return False
        return False

    def scrape(self):
        """Main scraping function"""
        print("🟢 Bids Extraction Started")

        try:
            # Load the main page with retry
            if not self.load_page_with_retry(self.base_url):
                print("❌ Failed to load main page")
                return

            # Apply filters
            self.apply_filters()

            # Get list of bids with Response Date and Title
            bids = self.get_bids_list()
            if not bids:
                print("❌ No bids found")
                return

            print(f"\n🎯 Found {len(bids)} bids to process")

            # Process each bid
            for index, bid in enumerate(bids, 1):
                try:
                    print(f"\n📄 Processing bid {index}/{len(bids)}")
                    print(f"🔢 Bid Number: {bid['bid_number']}")
                    print(f"🏢 Department: {bid['department']}")
                    print(f"📅 Response Date: {bid['response_date']}")
                    print(f"📋 Title: {bid['title']}")
                    print(f"🔗 URL: {bid['url']}")

                    # Load bid details page
                    if not self.load_page_with_retry(bid["url"]):
                        print("⚠️ Failed to load bid page, skipping...")
                        continue

                    # Extract bid details (passing bid info from main page)
                    bid_details = self.extract_bid_details(index, bid)
                    if bid_details:
                        # Download attachments
                        attachments = self.download_attachments(bid["bid_number"])
                        if attachments:
                            bid_details["Attachments"] = attachments

                        self.bids_data.append(bid_details)
                        self.save_to_excel()

                except Exception as e:
                    print(f"⚠️ Error processing bid {index}: {str(e)}")
                    continue

        except Exception as e:
            print(f"❌ Error during scraping: {str(e)}")
            play_notification_sound()
        finally:
            self.cleanup()

    def save_to_excel(self):
        """Save extracted data to Excel"""
        try:
            if self.bids_data:
                # Save Excel file in script folder
                excel_path = os.path.join(self.script_folder, "16_CalProcure.xlsx")
                df = pd.DataFrame(self.bids_data)

                # Ensure all required columns exist
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

                # Reorder columns to match required order
                df = df[required_columns]

                df.to_excel(excel_path, index=False)
                print(f"📊 Data saved to {excel_path}")
        except Exception as e:
            print(f"❌ Error saving to Excel: {str(e)}")
            play_notification_sound()

    def setup_signal_handler(self):
        """Setup handler for graceful shutdown on CTRL+C"""

        def signal_handler(signum, frame):
            print("\n⚠️ Graceful shutdown initiated...")
            self.cleanup()
            sys.exit(0)

        signal.signal(signal.SIGINT, signal_handler)

    def cleanup(self):
        """Cleanup function to handle incomplete downloads and browser closure"""
        print("🧹 Cleaning up...")
        try:
            # Move remaining downloads
            self.move_remaining_downloads()

            # Close browser
            if hasattr(self, "driver"):
                self.driver.quit()
                print("🌐 Browser closed successfully")

            # Rename script folder to COMPLETED if everything was successful
            if hasattr(self, "script_folder") and os.path.exists(self.script_folder):
                completed_folder = self.script_folder.replace(
                    "_IN_PROGRESS", "_COMPLETED"
                )
                try:
                    os.rename(self.script_folder, completed_folder)
                    print(f"✅ Renamed folder to: {completed_folder}")
                except Exception as e:
                    print(f"⚠️ Error renaming folder to COMPLETED: {str(e)}")

        except Exception as e:
            print(f"❌ Error during cleanup: {str(e)}")
            play_notification_sound()

    def extract_categories(self):
        """Extract UNSPSC categories"""
        categories = []
        try:
            rows = self.driver.find_elements(By.CSS_SELECTOR, "#unspscTable tbody tr")
            for i, row in enumerate(rows[:2]):  # Only first two categories
                code = row.find_element(By.CSS_SELECTOR, "td.bold").text
                desc = row.find_element(By.CSS_SELECTOR, "td:nth-child(2)").text
                categories.append(f"{code} - {desc}")
        except Exception as e:
            print(f"⚠️ Error extracting categories: {str(e)}")
        return "; ".join(categories)

    def extract_additional_summary(self):
        """Extract Pre Bid Conference information"""
        try:
            # Execute JavaScript to get conference information
            js_code = """
            function getPreBidConference() {
                const spans = document.querySelectorAll("span[data-if-container-label='conferenceRow']");
                if (!spans.length) return null;
                
                const conferenceInfo = [];
                for (let i = 0; i < spans.length; i += 2) {
                    const label = spans[i]?.textContent.trim().replace(':', '');
                    let value = spans[i + 1]?.textContent.trim();
                    
                    if (!value || value === '&nbsp;' || value === '\xa0') {
                        value = '';
                    }
                    
                    if (label) {
                        conferenceInfo.push(`${label}: ${value}`);
                    }
                }
                return conferenceInfo.length > 0 ? conferenceInfo.join('\\n') : null;
            }
            return getPreBidConference();
            """

            conference_info = self.driver.execute_script(js_code)
            if conference_info:
                print(f"✅ Extracted pre-bid conference info: {conference_info}")
                return conference_info
            else:
                print("ℹ️ No pre-bid conference information found")
                return ""

        except Exception as e:
            print(f"⚠️ Error extracting additional summary: {str(e)}")
            return ""

    def move_remaining_downloads(self):
        """Move any remaining downloads to bid-specific folder"""
        print("📦 Moving remaining downloads...")
        try:
            if not hasattr(self, "current_bid_number") or not self.current_bid_number:
                print("⚠️ No current bid number available for file movement")
                return

            if not os.path.exists(self.temp_download_folder):
                print("⚠️ Temporary download folder not found")
                return

            # Get bid-specific folder
            bid_folder = os.path.join(self.script_folder, self.current_bid_number)
            os.makedirs(bid_folder, exist_ok=True)

            # Wait for any ongoing downloads to complete
            self.wait_for_downloads()

            # Move any remaining files to bid-specific folder
            files_moved = []
            for filename in os.listdir(self.temp_download_folder):
                if not filename.endswith(".crdownload") and not filename.endswith(
                    ".tmp"
                ):
                    source = os.path.join(self.temp_download_folder, filename)
                    destination = os.path.join(bid_folder, filename)

                    try:
                        if safe_move(source, destination):
                            files_moved.append(filename)
                    except Exception as e:
                        print(f"⚠️ Error moving file {filename}: {str(e)}")

            if files_moved:
                print(
                    f"✅ Moved {len(files_moved)} files to bid folder {self.current_bid_number}"
                )
                print(f"📄 Files moved: {', '.join(files_moved)}")
            else:
                print("ℹ️ No files to move")

        except Exception as e:
            print(f"❌ Error during file movement: {str(e)}")
            play_notification_sound()

    def wait_for_downloads(self):
        """Wait for all downloads to complete"""
        max_wait_time = 300  # 5 minutes maximum wait time
        start_time = time.time()

        while time.time() - start_time < max_wait_time:
            if not any(
                filename.endswith(".crdownload") or filename.endswith(".tmp")
                for filename in os.listdir(self.temp_download_folder)
            ):
                return True
            time.sleep(5)

        print("⚠️ Download wait time exceeded maximum limit")
        return False

    def get_pre_bid_conference(self, driver):
        """Get pre-bid conference information"""
        try:
            js_code = """
                // script that does the concatenation:
                function getPreBidConference() {
                    // Find all spans with data-if-container-label="conferenceRow"
                    const spans = document.querySelectorAll("span[data-if-container-label='conferenceRow']");
                    
                    if (!spans.length) {
                        console.error("No conference spans found");
                        return null;
                    }

                    // Array to store the formatted conference information
                    const conferenceInfo = [];
                    
                    // Process spans in pairs (label + value)
                    for (let i = 0; i < spans.length; i += 2) {
                        const label = spans[i]?.textContent.trim();
                        let value = spans[i + 1]?.textContent.trim();
                        
                        // Replace &nbsp; with empty string
                        if (!value || value === '&nbsp;') {
                            value = '';
                        }
                        
                        // Add label-value pair if label exists
                        if (label) {
                            conferenceInfo.push(`${label}: ${value}`);
                        }
                    }

                    // Join all values with newlines
                    return conferenceInfo.join('\n');
                }

                // Execute and log the result
                return console.log(getPreBidConference());
            """

            conference_info = driver.execute_script(js_code)
            return conference_info if conference_info else ""

        except Exception as e:
            self.logger.error(f"Error getting pre-bid conference info: {str(e)}")
            return ""

    def get_bid_details(self, driver):
        """Get all bid details from the bid page"""
        try:
            # ... existing code ...

            bid_details = {
                # ... existing fields ...
                "Additional Summary": self.get_pre_bid_conference(driver),
                # ... other fields ...
            }

            return bid_details

        except Exception as e:
            self.logger.error(f"Error getting bid details: {str(e)}")
            return None

    def get_bid_url(self, department_name, bid_number):
        """Construct bid URL using department name and bid number"""
        try:
            # Clean up department name
            dept_name = department_name.strip()

            # Get business unit code from mapping
            business_unit = self.business_units.get(dept_name)

            if not business_unit:
                print(f"⚠️ No business unit found for department: {dept_name}")
                # Use default business unit code if no match found
                business_unit = "2660"

            # Construct and return URL
            bid_url = f"https://caleprocure.ca.gov/event/{business_unit}/{bid_number}"
            print(f"🔗 Generated URL: {bid_url}")
            return bid_url

        except Exception as e:
            print(f"❌ Error constructing bid URL: {str(e)}")
            return None

    def get_bids_list(self):
        """Get list of all bids from main page"""
        bids = []
        try:
            # Wait for table to be present
            print("⏳ Waiting for bids table...")
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody"))
            )

            # Get all bid rows
            rows = self.driver.find_elements(
                By.CSS_SELECTOR, "tr[id^='trRESP_INQA_HD_VW_GR$0_row']"
            )
            if not rows:
                print("❌ No bid rows found")
                return []

            print(f"📋 Found {len(rows)} bids")

            # Process each row
            for row in rows:
                try:
                    # Get bid number (Event ID)
                    bid_number = row.find_element(
                        By.CSS_SELECTOR, "td[data-if-label='tdEventId']"
                    ).text.strip()

                    # Get department name
                    department = row.find_element(
                        By.CSS_SELECTOR, "td[data-if-label='tdDepName']"
                    ).text.strip()

                    # Get Response Date
                    response_date = row.find_element(
                        By.CSS_SELECTOR, "td[data-if-label='tdEndDate']"
                    ).text.strip()
                    response_date = self.parse_date(
                        response_date
                    )  # Convert to YYYY-MM-DD

                    # Get Solicitation Title
                    title = row.find_element(
                        By.CSS_SELECTOR, "td[data-if-label='tdEventName']"
                    ).text.strip()

                    if not bid_number or not department:
                        continue

                    print(f"\n🔍 Processing bid:")
                    print(f"Bid Number: {bid_number}")
                    print(f"Department: {department}")
                    print(f"Response Date: {response_date}")
                    print(f"Title: {title}")

                    # Get business unit code from hardcoded mapping
                    business_unit = self.get_business_unit_code(department)
                    print(f"Business Unit Code: {business_unit}")

                    # Construct bid URL
                    bid_url = (
                        f"https://caleprocure.ca.gov/event/{business_unit}/{bid_number}"
                    )
                    print(f"URL: {bid_url}")

                    bids.append(
                        {
                            "bid_number": bid_number,
                            "department": department,
                            "business_unit": business_unit,
                            "url": bid_url,
                            "response_date": response_date,
                            "title": title,
                        }
                    )

                except Exception as e:
                    print(f"⚠️ Error processing bid row: {str(e)}")
                    continue

            print(f"\n✅ Successfully processed {len(bids)} bids")
            return bids

        except Exception as e:
            print(f"❌ Error getting bids list: {str(e)}")
            return []


def main():
    parser = argparse.ArgumentParser(description="Scrape CalProcure bids")
    parser.add_argument(
        "--days", type=int, default=2, help="Number of days back to scrape"
    )
    args = parser.parse_args()

    scraper = CalProcureScraper(days_back=args.days)
    scraper.scrape()


if __name__ == "__main__":
    main()
