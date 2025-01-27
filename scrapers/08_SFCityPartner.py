import subprocess
import time
import json
import os
from datetime import datetime, timedelta
import shutil
import sys


def get_progress_file_path():
    """Get the path to the progress file."""
    yesterday = datetime.now() - timedelta(days=1)
    base_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    date_folder = os.path.join(base_folder, yesterday.strftime("%Y-%m-%d"))
    main_folder = os.path.join(date_folder, "08_SFCityPartner_IN_PROGRESS")
    return os.path.join(main_folder, "scraping_progress.json")


def get_total_bids():
    """Get total number of bids from progress file or return default of 3."""
    progress_file = get_progress_file_path()
    if os.path.exists(progress_file):
        with open(progress_file, "r") as f:
            data = json.load(f)
            return data.get("total_bids", 3)  # Default to 3 if not found
    return 3  # Default number of bids


def get_scraped_bids_count():
    """Get number of already processed bids (both scraped and skipped)."""
    progress_file = get_progress_file_path()
    if os.path.exists(progress_file):
        with open(progress_file, "r") as f:
            data = json.load(f)
            scraped = len(data.get("scraped_bids", []))
            skipped = len(data.get("skipped_bids", []))
            return scraped + skipped
    return 0


def check_all_bids_processed():
    """Check if all bids have been processed (either scraped or skipped)."""
    progress_file = get_progress_file_path()
    if os.path.exists(progress_file):
        try:
            with open(progress_file, "r") as f:
                data = json.load(f)
                total_bids = data.get("total_bids", 0)
                scraped_bids = set(data.get("scraped_bids", []))
                skipped_bids = set(data.get("skipped_bids", []))

                processed_count = len(scraped_bids) + len(skipped_bids)
                if total_bids > 0 and processed_count >= total_bids:
                    print("\n✅ All bids have already been processed!")
                    print(f"Total bids: {total_bids}")
                    print(
                        f"Scraped bids ({len(scraped_bids)}): {sorted(list(scraped_bids))}"
                    )
                    print(
                        f"Skipped bids ({len(skipped_bids)}): {sorted(list(skipped_bids))}"
                    )
                    print("\nNo more bids to process. Exiting...")
                    return True

                print(f"\nProgress: {processed_count}/{total_bids} bids processed")
                print(f"- Scraped: {len(scraped_bids)} bids")
                print(f"- Skipped: {len(skipped_bids)} bids")
                print(f"- Remaining: {total_bids - processed_count} bids")

        except Exception as e:
            print(f"Error reading progress file: {e}")
            return False
    return False


def main():
    print("🚀 Starting SF City Partner Multi-Run Scraper")

    # Check if all bids are already processed
    if check_all_bids_processed():
        return

    # Path to the scraper utility script (in the same directory)
    scraper_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "08_SFCityPartner_util.py"
    )

    total_bids = get_total_bids()
    current_bid = get_scraped_bids_count()

    print(f"Total bids to scrape: {total_bids}")
    print(f"Already scraped: {current_bid}")
    print(f"Remaining bids: {total_bids - current_bid}")

    try:
        while current_bid < total_bids:
            try:
                print(f"\n📑 Starting scrape for bid {current_bid + 1} of {total_bids}")
                print("=" * 50)

                # Get days argument from command line
                days_arg = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] == '--days' else None
                days_value = sys.argv[2] if days_arg and len(sys.argv) > 2 else None

                # Run the scraper script with days argument if provided
                cmd = ["python", scraper_path]
                if days_value:
                    cmd.extend(["--days", days_value])
                
                result = subprocess.run(cmd, check=True, text=True)

                # Check if the folder was marked as completed (indicating no bids found)
                yesterday = datetime.now() - timedelta(days=1)
                base_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                date_folder = os.path.join(base_folder, yesterday.strftime("%Y-%m-%d"))
                completed_folder = os.path.join(date_folder, "08_SFCityPartner_COMPLETED")
                
                if os.path.exists(completed_folder):
                    print("\n✅ No bids found - Scraping completed")
                    return

                # Check if all bids are processed after each run
                if check_all_bids_processed():
                    break

                # Update progress
                current_bid = get_scraped_bids_count()

                if current_bid < total_bids:
                    print(f"\n⏳ Waiting 30 seconds before next bid...")
                    time.sleep(30)  # Wait between runs

            except subprocess.CalledProcessError as e:
                print(f"❌ Error running scraper: {e}")
                retry = input("\nRetry this bid? (y/n): ").lower().strip()
                if retry != "y":
                    break
            except KeyboardInterrupt:
                print("\n⚠️ Script interrupted by user")
                break
            except Exception as e:
                print(f"❌ Unexpected error: {e}")
                break

    finally:
        # Ensure cleanup of temporary download folder
        yesterday = datetime.now() - timedelta(days=1)
        base_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        date_folder = os.path.join(base_folder, yesterday.strftime("%Y-%m-%d"))
        main_folder = os.path.join(date_folder, "08_SFCityPartner_IN_PROGRESS")
        temp_download_folder = os.path.join(main_folder, "08_SFCityPartner")
        
        try:
            if os.path.exists(temp_download_folder):
                shutil.rmtree(temp_download_folder)
                print(f"✅ Removed temporary download folder: {temp_download_folder}")
        except Exception as e:
            print(f"⚠️ Error removing temporary folder: {str(e)}")


if __name__ == "__main__":
    main()
