import os
import sys
from pathlib import Path

# Add the project root directory to Python path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root)

from utils.excel_processor import ExcelProcessor
import pandas as pd


def test_florida_marketplace_excel():
    """Test the Excel processor with the Florida Marketplace Excel file"""
    print("\n🚀 Starting Excel Processor Test\n")

    # Specific Excel file path - using Path for better path handling
    excel_path = (
        Path(os.getcwd())
        / "2024-11-29"
        / "12_BonfireSites_IN_PROGRESS"
        / "12_BonfireSites.xlsx"
    )

    try:
        # Verify file exists
        if not excel_path.exists():
            print(f"❌ Excel file not found at: {excel_path}")
            return

        print(f"Found Excel file: {excel_path}")

        # Read original Excel to show structure
        original_df = pd.read_excel(excel_path)
        print("\nOriginal Excel structure:")
        print(f"Columns: {original_df.columns.tolist()}")
        print(f"Number of rows: {len(original_df)}")

        # Initialize processor
        print("\nInitializing Excel Processor...")
        processor = ExcelProcessor()
        print("✅ Excel Processor initialized")

        # Process the Excel file
        print("\nProcessing Excel file...")
        result = processor.process_excel_file(str(excel_path))

        if result:
            print("✅ Excel processing completed successfully")

            # Check the output file
            output_file = excel_path.with_name(f"{excel_path.name}")
            if output_file.exists():
                print(f"\nReading processed file: {output_file}")
                processed_df = pd.read_excel(output_file)

                # Verify all required columns were added
                required_columns = {
                    "API Category": "Category",
                    "API Category ID": "Category",
                    "API Notice Type": "Notice Type",
                    "API Agency": "Agency",
                    "API State": None,  # No corresponding original column
                }

                print("\nVerifying new columns:")
                for new_col, original_col in required_columns.items():
                    if new_col in processed_df.columns:
                        if original_col and original_col in processed_df.columns:
                            # Get index positions to verify correct ordering
                            new_idx = processed_df.columns.get_loc(new_col)
                            orig_idx = processed_df.columns.get_loc(original_col)

                            # Check if new column is before original column
                            if new_idx < orig_idx:
                                print(
                                    f"✅ {new_col:<15} | Correctly placed before {original_col}"
                                )
                            else:
                                print(
                                    f"❌ {new_col:<15} | Incorrectly placed after {original_col}"
                                )
                        else:
                            print(f"✅ {new_col:<15} | Added to columns")

                        # Show sample value
                        value = processed_df[new_col].iloc[0]
                        value_str = str(value) if pd.notna(value) else "None"
                        print(f"   Sample value: {value_str}")
                    else:
                        print(f"❌ {new_col:<15} | Column missing")

                # Compare data between original and processed files
                print("\nData comparison (first row):")
                print("\nOriginal data:")
                for col in ["Notice Type", "Agency", "Category"]:
                    if col in original_df.columns:
                        print(f"{col}: {original_df[col].iloc[0]}")

                print("\nProcessed data:")
                for col in required_columns:
                    if col in processed_df.columns:
                        value = processed_df[col].iloc[0]
                        value_str = str(value) if pd.notna(value) else "None"
                        print(f"{col}: {value_str}")
            else:
                print(f"❌ Output file not found: {output_file}")
        else:
            print("❌ Excel processing failed")

    except Exception as e:
        print(f"❌ Error during testing: {str(e)}")
        import traceback

        print(f"Stack trace:\n{traceback.format_exc()}")

    print("\n=== Test Complete ===")


def test_notice_type_matching():
    """Test notice type matching according to excel_extra_cols.md"""
    processor = ExcelProcessor()

    # Get notice types from API
    notice_types = processor.fetch_api_data("notice")

    test_cases = [
        ("RFP for Construction", "Request For Proposal"),
        ("Request for Quote (RFQ)", "Request For Proposal"),
        ("Invitation for Bid", "Request For Proposal"),
        ("Sources Sought Notice", "Sources Sought / RFI"),
        ("RFI Document", "Sources Sought / RFI"),
        ("Award Notice Posted", "Award Notice"),
    ]

    print("\nTesting Notice Type Matching:")
    for text, expected in test_cases:
        notice_type, notice_id = processor.determine_notice_type(text, notice_types)
        print(f"\nInput: {text}")
        print(f"Expected: {expected}")
        print(f"Got: {notice_type}")
        print("✅ Match" if notice_type == expected else "❌ No match")


def test_category_matching():
    """Test category matching according to excel_extra_cols.md"""
    processor = ExcelProcessor()

    # Get categories from API
    categories = processor.fetch_api_data("category")

    test_cases = [
        {
            "title": "Parking Lot Maintenance",
            "description": "Repair and maintain parking facilities",
            "category": "72103301 - Parking lot maintenance",
            "expected": "060 - Building Repairs and Maintenance",
        },
        {
            "title": "Tower Light System",
            "description": "Replace telecommunication tower lighting",
            "category": "72141118 - Telecommunication transmission tower",
            "expected": "013 - Communications – Broadcasting and Telecommunication",
        },
    ]

    print("\nTesting Category Matching:")
    for case in test_cases:
        category, category_id = processor.find_best_category_match(
            case["title"], case["description"], case["category"], categories
        )
        print(f"\nInput Category: {case['category']}")
        print(f"Expected: {case['expected']}")
        print(f"Got: {category}")
        print("✅ Match" if category == case["expected"] else "❌ No match")


if __name__ == "__main__":
    test_florida_marketplace_excel()
