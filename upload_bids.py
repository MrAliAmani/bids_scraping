from datetime import datetime, timedelta
import os
import boto3
import sys
from botocore.exceptions import NoCredentialsError, ClientError
import shutil
import argparse
import glob

import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

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


def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Upload files or folders to MinIO/S3")
    parser.add_argument(
        "path",
        nargs="?",
        default=None,
        help="Path to date folder (defaults to yesterday's date folder)",
    )
    parser.add_argument(
        "--backup",
        action="store_true",
        help="Create a backup copy instead of removing source files",
    )
    args = parser.parse_args()

    # Use default path if none provided
    if args.path is None:
        args.path = get_default_path()

    # Normalize the path to handle any path separators
    args.path = os.path.normpath(args.path)

    # Verify the path exists before proceeding
    if not os.path.exists(args.path):
        print(f"❌ Error: Path does not exist: {args.path}")
        sys.exit(1)

    # Calculate the target folder name (yesterday's date)
    current_folder = str(datetime.now().date() - timedelta(days=1))
    s3_folder = f"State/attachments/{current_folder}"

    # Initialize S3 client
    s3_client = boto3.client(
        "s3",
        aws_access_key_id=AWS_ACCESS_KEY_ID,
        aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
        region_name=AWS_DEFAULT_REGION,
        endpoint_url="http://localhost:9000",
    )

    # Enable versioning on the bucket
    enable_versioning(s3_client, AWS_BUCKET_NAME)

    print(f"Processing: {args.path}")
    success = upload_to_s3(s3_client, args.path, AWS_BUCKET_NAME, s3_folder)

    # Remove empty folders after upload
    print("Cleaning up empty folders...")
    remove_empty_folders(s3_client, AWS_BUCKET_NAME, s3_folder)

    # Handle backup if requested and upload was successful
    if success and args.backup and os.path.isdir(args.path):
        try:
            backup_path = (
                f"{args.path}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
            shutil.copytree(args.path, backup_path)
            print(f"📁 Created backup at: {backup_path}")
        except Exception as e:
            print(f"⚠️ Warning: Could not create backup of {args.path}: {e}")

    if success:
        if args.backup:
            try:
                # Create backup before cleanup
                backup_path = f"{args.path}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                shutil.copytree(args.path, backup_path)
                print(f"📁 Created backup at: {backup_path}")
            except Exception as e:
                print(f"⚠️ Warning: Could not create backup of {args.path}: {e}")
        
        # Clean up resources after successful upload
        if cleanup_resources(args.path):
            print("✅ Resources cleaned up successfully")
        else:
            print("⚠️ Warning: Some resources could not be cleaned up")
            
        print("✅ Upload process completed successfully")
    else:
        print("❌ Upload process completed with errors")
        sys.exit(1)


if __name__ == "__main__":
    main()
