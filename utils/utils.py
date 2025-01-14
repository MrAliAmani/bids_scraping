import os
import shutil
import winsound


def list_excel_files():
    """
    Create a list of Excel file names based on the script names in the scrapers folder.

    Returns:
        list: A list of Excel file names.
    """
    excel_files = ["12_BonfireSites.xlsx"]

    # Assuming the scrapers directory is a sibling of the utils directory
    utils_dir = os.path.dirname(__file__)
    scrapers_dir = os.path.join(utils_dir, "..", "scrapers")
    scrapers_dir = os.path.abspath(scrapers_dir)

    if not os.path.isdir(scrapers_dir):
        print(f"Scrapers directory not found: {scrapers_dir}")
        return excel_files

    for file in os.listdir(scrapers_dir):
        if file.endswith(".py"):
            excel_name = os.path.splitext(file)[0] + ".xlsx"
            excel_files.append(excel_name)

    return excel_files


def safe_move(source_path, destination_path):
    """
    Safely moves the file from source_path to destination_path, but skips the move
    if the file is in the list of Excel files generated from scrapers.

    Args:
        source_path (str): The full path to the source file.
        destination_path (str): The full path to the destination.

    Returns:
        bool: True if the file was moved, False if it was skipped.
    """
    # Get the list of Excel files
    excel_files = list_excel_files()

    # Get the file name (without the directory) from source_path
    file_name = os.path.basename(source_path)

    # Check if the file is in the list of Excel files
    if file_name in excel_files:
        print(f"Skipping move for {file_name} as it is an Excel file.")
        return False

    # Move the file using shutil.move
    shutil.move(source_path, destination_path)
    print(f"File {file_name} moved to {destination_path}.")
    return True


def play_notification_sound():
    """Play a notification sound."""
    winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
