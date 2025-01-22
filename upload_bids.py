from datetime import datetime, timedelta
import os
import boto3
import sys
from botocore.exceptions import NoCredentialsError, ClientError
import shutil
import argparse
import glob
import pandas as pd
import requests
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from rich.console import Console

import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

console = Console()

# MinIO/S3 Configuration
AWS_ACCESS_KEY_ID = "minioadmin"  # Your MinIO access key
AWS_SECRET_ACCESS_KEY = "minioadmin"  # Your MinIO secret key
AWS_DEFAULT_REGION = "us-east-2"  # Your MinIO region
AWS_BUCKET_NAME = "bidsdocuments-bucket"  # Your MinIO bucket name
AWS_ENDPOINT_URL = "http://localhost:9000"  # Your MinIO endpoint URL
AWS_USE_PATH_STYLE_ENDPOINT = "false"  # Use path-style endpoint
AWS_PATH = "https://bidsdocuments.s3.us-east-2.amazonaws.com/"  # S3 path

# Base Path Configuration
BASE_PATH = r"D:\Projects\web-scraping"  # Base path for the project


def enable_versioning(s3_client, bucket_name):
    """Enable versioning on the S3 bucket"""
    try:
        s3_client.put_bucket_versioning(
            Bucket=bucket_name, VersioningConfiguration={"Status": "Enabled"}
        )
        print("[OK] Versioning enabled for bucket: {bucket_name}")
    except Exception as e:
        print(f"[ERROR] Failed to enable versioning: {e}")


def find_scraper_folders(base_path):
    """Find all scraper folders (ending with _IN_PROGRESS or _COMPLETED)"""
    scraper_folders = []
    for root, dirs, _ in os.walk(base_path):
        scraper_folders.extend(
            os.path.join(root, dir_name)
            for dir_name in dirs
            if dir_name.endswith("_IN_PROGRESS") or dir_name.endswith("_COMPLETED")
        )
    return scraper_folders


def is_bid_folder(folder_name):
    """Check if the folder name looks like a solicitation number"""
    # Most solicitation numbers contain both numbers and hyphens
    folder_name = os.path.basename(folder_name)
    return any(char.isdigit() for char in folder_name)


def upload_to_s3(s3_client, local_path, bucket_name, s3_folder):
    """Upload bid folders and Excel files to S3"""
    success = True

    try:
        if not os.path.exists(local_path):
            print(f"[ERROR] Path not found: {local_path}")
            return False

        # First, check if the path itself is a completed folder
        if os.path.basename(local_path).endswith("_COMPLETED"):
            scraper_folders = [local_path]
        else:
            # Find all scraper folders
            scraper_folders = find_scraper_folders(local_path)

        if not scraper_folders:
            print(f"[WARNING] No scraper folders found in: {local_path}")
            return False

        print(f"[INFO] Found {len(scraper_folders)} scraper folders")

        for scraper_folder in scraper_folders:
            scraper_name = (
                os.path.basename(scraper_folder)
                .split("_IN_PROGRESS")[0]
                .split("_COMPLETED")[0]
            )
            print(f"[INFO] Processing scraper: {scraper_name}")

            # Look for Excel file in both the scraper folder and its parent directory
            excel_file = f"{scraper_name}.xlsx"
            possible_excel_paths = [
                os.path.join(scraper_folder, excel_file),  # In the completed folder
                os.path.join(
                    os.path.dirname(scraper_folder), excel_file
                ),  # In the parent directory
                os.path.join(
                    os.path.dirname(os.path.dirname(scraper_folder)), excel_file
                ),  # In the root
            ]

            excel_uploaded = False
            for excel_path in possible_excel_paths:
                if os.path.exists(excel_path):
                    s3_file_path = f"{s3_folder}/{excel_file}".replace("\\", "/")
                    try:
                        s3_client.upload_file(excel_path, bucket_name, s3_file_path)
                        print(f"[OK] Uploaded Excel file: {s3_file_path}")
                        excel_uploaded = True
                        break
                    except Exception as e:
                        print(f"[ERROR] Failed to upload Excel file {excel_path}: {e}")
                        success = False

            if not excel_uploaded:
                print(f"[WARNING] No Excel file found for {scraper_name}")

            # Upload bid folders
            bid_folders_found = False
            for item in os.listdir(scraper_folder):
                item_path = os.path.join(scraper_folder, item)
                if os.path.isdir(item_path) and is_bid_folder(item):
                    bid_folders_found = True
                    for root, _, files in os.walk(item_path):
                        for file in files:
                            local_file_path = os.path.join(root, file)
                            relative_path = os.path.relpath(
                                local_file_path, scraper_folder
                            )
                            s3_file_path = f"{s3_folder}/{relative_path}".replace(
                                "\\", "/"
                            )

                            try:
                                s3_client.upload_file(
                                    local_file_path, bucket_name, s3_file_path
                                )
                                print(f"[OK] Uploaded: {s3_file_path}")
                            except Exception as e:
                                print(
                                    f"[ERROR] Failed to upload {local_file_path}: {e}"
                                )
                                success = False

            if not bid_folders_found:
                print(f"[INFO] No bid folders found in {scraper_folder}")

    except Exception as e:
        print(f"[ERROR] Error during upload: {e}")
        success = False

    return success


def remove_empty_folders(s3_client, bucket_name, prefix):
    """Remove empty folders from S3 bucket"""
    try:
        # List all objects with the given prefix
        paginator = s3_client.get_paginator("list_objects_v2")
        pages = paginator.paginate(Bucket=bucket_name, Prefix=prefix)

        # Collect all folder paths
        folders = set()
        files = set()

        for page in pages:
            if "Contents" in page:
                for obj in page["Contents"]:
                    key = obj["Key"]
                    # Skip if the object is zero bytes (likely a folder marker)
                    if obj["Size"] == 0:
                        folders.add(key)
                    else:
                        files.add(key)
                        # Add all parent folders
                        parts = key.split("/")
                        for i in range(len(parts) - 1):
                            folders.add("/".join(parts[: i + 1]) + "/")

        # Find empty folders (those not containing any files)
        empty_folders = []
        for folder in folders:
            has_files = False
            for file in files:
                if file.startswith(folder):
                    has_files = True
                    break
            if not has_files:
                empty_folders.append(folder)

        # Delete empty folders
        for folder in empty_folders:
            try:
                s3_client.delete_object(Bucket=bucket_name, Key=folder)
                print(f"🗑️ Removed empty folder: {folder}")
            except Exception as e:
                print(f"⚠️ Failed to remove folder {folder}: {e}")

    except Exception as e:
        print(f"❌ Error while removing empty folders: {e}")


def get_default_path():
    """Get the default path using yesterday's date"""
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    # Assuming the script is in the project root directory
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, yesterday)


def cleanup_resources(folder_path: str) -> bool:
    """Clean up resources after successful upload"""
    try:
        if not os.path.exists(folder_path):
            print(f"[WARNING] Path not found for cleanup: {folder_path}")
            return False

        print(f"\n🧹 Cleaning up resources in: {folder_path}")
        
        # Remove the folder and all its contents
        try:
            shutil.rmtree(folder_path)
            print(f"✅ Removed folder and contents: {folder_path}")
            return True
        except Exception as e:
            print(f"❌ Error removing folder {folder_path}: {str(e)}")
            return False

    except Exception as e:
        print(f"❌ Error during cleanup: {str(e)}")
        return False


class BidsUploader:
    def __init__(self):
        self.api_endpoint = "https://bidsportal.com/api/uploadBids"
        self.headers = {
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        
    def clean_data(self, value) -> str:
        """Clean data by handling NaN values and converting to string"""
        if pd.isna(value) or value is None:
            return ""
        return str(value).strip()
        
    def upload_bid(self, bid_data: Dict) -> bool:
        """Upload a single bid to the API"""
        try:
            # Clean the data before sending
            cleaned_data = {
                key: self.clean_data(value)
                for key, value in bid_data.items()
            }
            
            response = requests.post(
                self.api_endpoint,
                json=cleaned_data,
                headers=self.headers,
                verify=False
            )
            response.raise_for_status()
            return True
        except Exception as e:
            logger.error(f"Error uploading bid: {str(e)}")
            return False

    def process_excel_file(self, excel_path: str) -> bool:
        """Process and upload bids from a single Excel file"""
        try:
            # Read Excel file
            df = pd.read_excel(excel_path)
            
            if df.empty:
                logger.warning(f"Empty Excel file: {excel_path}")
                return False
                
            # Check for required API columns
            required_columns = ['API_Category', 'API_Category_ID', 'API_Notice_Type', 'API_Agency', 'API_State']
            if not all(col in df.columns for col in required_columns):
                logger.error(f"Missing required API columns in {excel_path}")
                return False
                
            total_rows = len(df)
            successful_uploads = 0
            
            # Process each row
            for index, row in df.iterrows():
                try:
                    # Calculate progress
                    progress = int(((index + 1) / total_rows) * 100)
                    print(f"\rUploading bid {index + 1}/{total_rows} ({progress}%)", end='')
                    
                    # Prepare bid data
                    bid_data = {
                        "title": row.get('Title') or row.get('Solicitation Title', ''),
                        "description": row.get('Description', ''),
                        "category": row.get('API_Category'),
                        "category_id": row.get('API_Category_ID'),
                        "notice_type": row.get('API_Notice_Type'),
                        "agency": row.get('API_Agency'),
                        "state": row.get('API_State'),
                        "bid_url": row.get('Bid Detail Page URL', ''),
                        "bid_number": row.get('Bid Number', ''),
                        "posted_date": row.get('Posted Date', ''),
                        "submission_date": row.get('Submission Date', '')
                    }
                    
                    # Upload bid
                    if self.upload_bid(bid_data):
                        successful_uploads += 1
                    
                except Exception as e:
                    logger.error(f"Error processing row {index + 1}: {str(e)}")
                    continue
                    
            print(f"\n✅ Uploaded {successful_uploads}/{total_rows} bids from {Path(excel_path).name}")
            return successful_uploads > 0
            
        except Exception as e:
            logger.error(f"Error processing file {excel_path}: {str(e)}")
            return False

def upload_data(folder_path: str) -> Tuple[bool, str]:
    """Upload data to S3 using MinIO"""
    try:
        # Initialize S3 client
        s3_client = boto3.client(
            "s3",
            aws_access_key_id=AWS_ACCESS_KEY_ID,
            aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
            region_name=AWS_DEFAULT_REGION,
            endpoint_url=AWS_ENDPOINT_URL,
        )

        # Calculate the target folder name (yesterday's date)
        current_folder = str(datetime.now().date() - timedelta(days=1))
        s3_folder = f"State/attachments/{current_folder}"

        # Enable versioning on the bucket
        enable_versioning(s3_client, AWS_BUCKET_NAME)

        # Upload files
        success = upload_to_s3(s3_client, folder_path, AWS_BUCKET_NAME, s3_folder)

        # Remove empty folders after upload
        print("Cleaning up empty folders...")
        remove_empty_folders(s3_client, AWS_BUCKET_NAME, s3_folder)

        if success:
            # Clean up resources after successful upload
            if cleanup_resources(folder_path):
                return True, "Upload and cleanup successful"
            else:
                return False, "Upload successful but cleanup failed"
        else:
            return False, "Upload failed"

    except Exception as e:
        error_msg = f"Error during upload: {str(e)}"
        print(f"❌ {error_msg}")
        return False, error_msg

def upload_bids_from_cli(base_path: str = None) -> bool:
    """Upload bids from Excel files in yesterday's COMPLETED folders"""
    try:
        # Get yesterday's date folder
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        
        if not base_path:
            base_path = os.path.join(os.getcwd(), yesterday)
        
        if not os.path.exists(base_path):
            print(f"❌ Yesterday's folder not found: {base_path}")
            return False
            
        print(f"\nUploading bids from: {base_path}")

        # Find all COMPLETED folders
        success = True
        folders_processed = 0
        
        if base_path.endswith('COMPLETED'):
            # If specific COMPLETED folder provided
            completed_folders = [base_path]
        else:
            # Find all COMPLETED folders in base path
            completed_folders = []
            for root, dirs, files in os.walk(base_path):
                if root.endswith('COMPLETED'):
                    completed_folders.append(root)

        if not completed_folders:
            print("No COMPLETED folders found")
            return False

        # Process each completed folder
        for folder in completed_folders:
            print(f"\n📁 Processing folder: {folder}")
            upload_success, upload_message = upload_data(folder)
            
            if upload_success:
                print(f"✅ Successfully processed: {folder}")
                folders_processed += 1
            else:
                print(f"❌ Failed to process {folder}: {upload_message}")
                success = False

        print(f"\n🎉 Upload complete! Processed {folders_processed} folders")
        return success

    except Exception as e:
        print(f"\n❌ Error during upload: {str(e)}")
        return False

if __name__ == "__main__":
    # Allow optional base path argument
    base_path = sys.argv[1] if len(sys.argv) > 1 else None
    success = upload_bids_from_cli(base_path)
    sys.exit(0 if success else 1)
