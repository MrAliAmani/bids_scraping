import json
import os
import sys
from pathlib import Path

# Add project root to Python path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.append(str(PROJECT_ROOT))

import pandas as pd
from utils.category_matcher import CategoryMatcher
from utils.excel_processor import ExcelProcessor  # For getting API categories
from rich.console import Console
from rich.table import Table
from datetime import datetime
from sklearn.metrics import precision_score, recall_score, f1_score
import numpy as np
from typing import Dict, List, Tuple, Optional

console = Console()


def calculate_accuracy_metrics(predictions_df: pd.DataFrame) -> Dict:
    """Calculate accuracy metrics comparing against QC Category if present"""
    metrics = {
        "total": len(predictions_df),
        "correct": 0,
        "incorrect": 0,
        "unpredicted": 0,
        "accuracy": 0.0,
    }

    # Only calculate accuracy if QC Category exists
    if "QC Category" not in predictions_df.columns:
        console.print(
            "[yellow]No QC Category column found - skipping accuracy calculation[/yellow]"
        )
        return metrics

    for idx, row in predictions_df.iterrows():
        predicted = row.get("API_Category", "").strip()
        actual = row.get("QC Category", "").strip()

        if not predicted:
            metrics["unpredicted"] += 1
        elif predicted.lower() == actual.lower():  # Case-insensitive comparison
            metrics["correct"] += 1
        else:
            metrics["incorrect"] += 1

    # Calculate accuracy based only on predicted items
    predicted_total = metrics["correct"] + metrics["incorrect"]
    if predicted_total > 0:
        metrics["accuracy"] = metrics["correct"] / predicted_total

    return metrics


def process_excel_with_methods(input_excel: str, method: str = "all"):
    """
    Process Excel file with specified matching method

    Args:
        input_excel (str): Path to input Excel file
        method (str): Matching method to use ('similarity', 'weighted_fuzzy', 'hybrid', 'llm', 'ai_enhanced')
    """
    # Initialize ExcelProcessor and get categories
    excel_processor = ExcelProcessor()
    api_categories = excel_processor.fetch_api_data("category")

    if not api_categories:
        console.print("[red]Failed to fetch categories from API[/red]")
        return None

    # Initialize matcher
    matcher = CategoryMatcher(api_categories)

    # Define available methods
    reliable_methods = {
        "similarity": matcher.match_by_similarity,
        "weighted_fuzzy": matcher.match_by_weighted_fuzzy,
        "hybrid": matcher.match_by_hybrid,
        "fuzzy_ollama": matcher.match_by_fuzzy_ollama,
        "confident_hybrid": matcher.match_by_confident_hybrid,
    }

    llm_methods = {
        "llm": matcher.match_by_llm,
        "ai_enhanced": matcher.match_by_ai_enhanced,
    }

    # Select methods to run
    if method == "all":
        methods = reliable_methods  # Only run reliable methods by default
    elif method in reliable_methods:
        methods = {method: reliable_methods[method]}
    elif method in llm_methods:
        methods = {method: llm_methods[method]}
    else:
        console.print(
            f"[red]Invalid method: {method}. Available methods: {', '.join(reliable_methods.keys() | llm_methods.keys())}[/red]"
        )
        return None

    # Setup output directory
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = PROJECT_ROOT / f"results_{timestamp}"
    output_dir.mkdir(exist_ok=True)
    base_name = Path(input_excel).stem

    # Read and clean data
    df = pd.read_excel(input_excel)
    df = df.fillna("")

    # Track metrics
    metrics = {
        "total_records": len(df),
        "duplicate_count": 0,
        "processed_count": 0,
        "methods": {},
    }

    # Process with each method separately
    for method_name, method_func in methods.items():
        console.print(f"\n[bold]Processing with {method_name} method[/bold]")

        sl_no = 1
        matcher.processed_bids.clear()  # Reset duplicate tracking for each method

        # Create DataFrame to store all results, including non-matches
        all_results = []

        for idx, row in df.iterrows():
            # Get fields
            possible_title_columns = ["Title", "Solicitation Title", "title", "TITLE"]
            title = next(
                (
                    str(row.get(col, "")).strip()
                    for col in possible_title_columns
                    if col in row
                ),
                "",
            )
            description = str(row.get("Description", "")).strip()
            category = str(row.get("Category", "")).strip()

            # Skip empty records
            if not any([title, description, category]):
                continue

            # Check duplicates
            if matcher.is_duplicate_bid(title, description or category):
                metrics["duplicate_count"] += 1
                continue

            metrics["processed_count"] += 1

            try:
                # Process bid
                if not description and category:
                    match, score = method_func(title, category, category)
                else:
                    match, score = method_func(title, description, category)

                row_data = row.copy()
                if match and score > 0:
                    row_data["API_Category"] = match["category_name"]
                    row_data["API_Category_ID"] = match["category_id"]
                    row_data["Match_Score"] = score
                else:
                    # Leave blank for low confidence matches
                    row_data["API_Category"] = ""
                    row_data["API_Category_ID"] = ""
                    row_data["Match_Score"] = 0.0

                row_data["SL No"] = sl_no
                all_results.append(row_data)
                sl_no += 1

            except Exception as e:
                console.print(f"[red]Error processing row {idx + 1}: {str(e)}[/red]")

        # Save all results, including blank predictions
        if all_results:
            result_df = pd.DataFrame(all_results)
            output_file = output_dir / f"{base_name}_{method_name}.xlsx"
            result_df.to_excel(output_file, index=False)
            console.print(
                f"[green]Saved {method_name} results to {output_file}[/green]"
            )

            # Calculate accuracy metrics only for non-blank predictions
            predicted_df = result_df[
                result_df["API_Category"].notna() & (result_df["API_Category"] != "")
            ]
            method_metrics = calculate_accuracy_metrics(predicted_df)
            metrics["methods"][method_name] = method_metrics

            # Add prediction coverage metrics
            total_rows = len(result_df)
            predicted_rows = len(predicted_df)
            coverage = predicted_rows / total_rows if total_rows > 0 else 0
            metrics["methods"][method_name]["coverage"] = coverage
            console.print(f"[blue]Prediction coverage: {coverage:.2%}[/blue]")

    # Save combined metrics to single JSON file
    metrics_file = output_dir / f"{base_name}_metrics.json"
    with open(metrics_file, "w") as f:
        json.dump(metrics, f, indent=2)

    # Print summary
    console.print("\n[bold]Processing Summary:[/bold]")
    console.print(f"Total records: {metrics['total_records']}")
    console.print(f"Duplicate records: {metrics['duplicate_count']}")
    console.print(f"Processed records: {metrics['processed_count']}")

    console.print("\n[bold]Method Results:[/bold]")
    for method_name, method_metrics in metrics["methods"].items():
        console.print(f"\n[cyan]{method_name.title()}:[/cyan]")
        console.print(f"Total Matches: {method_metrics['total']}")
        console.print(f"Correct: {method_metrics['correct']}")
        console.print(f"Incorrect: {method_metrics['incorrect']}")
        console.print(f"Accuracy: {method_metrics['accuracy']:.2%}")

    return output_dir


def process_multiple_excel_files(method: str = "all"):
    """Process multiple Excel files with specified matching method"""
    # Define the paths to both Excel files
    excel_files = [
        PROJECT_ROOT / "2024-12-03/03_TXSMartBuy_COMPLETED/03_TXSMartBuy.xlsx",
        PROJECT_ROOT
        / "2024-12-07/06_MyFloridaMarketPlace_COMPLETED/06_MyFloridaMarketPlace.xlsx",
    ]

    # Initialize ExcelProcessor and get categories once
    excel_processor = ExcelProcessor()
    api_categories = excel_processor.fetch_api_data("category")

    if not api_categories:
        console.print("[red]Failed to fetch categories from API[/red]")
        return None

    # Initialize matcher
    matcher = CategoryMatcher(api_categories)

    # Define available methods
    reliable_methods = {
        "similarity": matcher.match_by_similarity,
        "weighted_fuzzy": matcher.match_by_weighted_fuzzy,
        "hybrid": matcher.match_by_hybrid,
        "confident_hybrid": matcher.match_by_confident_hybrid,
    }

    llm_methods = {
        "llm": matcher.match_by_llm,
        "ai_enhanced": matcher.match_by_ai_enhanced,
        "fuzzy_ollama": matcher.match_by_fuzzy_ollama,
    }

    # Select methods to run
    if method == "all":
        methods = reliable_methods  # Only run reliable methods by default
    elif method in reliable_methods:
        methods = {method: reliable_methods[method]}
    elif method in llm_methods:
        methods = {method: llm_methods[method]}
    else:
        console.print(
            f"[red]Invalid method: {method}. Available methods: {', '.join(reliable_methods.keys() | llm_methods.keys())}[/red]"
        )
        return None

    # Setup output directory
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = PROJECT_ROOT / f"results_{timestamp}"
    output_dir.mkdir(exist_ok=True)

    # Process each Excel file
    for excel_path in excel_files:
        if excel_path.exists():
            console.print(f"\n[bold]Processing {excel_path.name}[/bold]")
            try:
                process_excel_with_methods(str(excel_path), method)
            except Exception as e:
                console.print(
                    f"[red]Error processing {excel_path.name}: {str(e)}[/red]"
                )
        else:
            console.print(f"[yellow]File not found: {excel_path}[/yellow]")

    return output_dir


if __name__ == "__main__":
    # Add argument parsing
    import argparse

    parser = argparse.ArgumentParser(
        description="Process Excel files with category matching"
    )
    parser.add_argument(
        "--method",
        type=str,
        default="all",
        choices=[
            "all",
            "similarity",
            "weighted_fuzzy",
            "hybrid",
            "fuzzy_ollama",
            "llm",
            "ai_enhanced",
            "confident_hybrid",
        ],
        help="Matching method to use (llm and ai_enhanced require explicit selection)",
    )
    args = parser.parse_args()

    # Warn if LLM method is selected
    if args.method in ["llm", "ai_enhanced"]:
        console.print(
            "[yellow]Warning: LLM-based methods may encounter rate limits or API issues[/yellow]"
        )
        proceed = input("Do you want to continue? (y/n): ")
        if proceed.lower() != "y":
            sys.exit(0)

    try:
        output_dir = process_multiple_excel_files(args.method)
        if output_dir:
            console.print(f"\n[green]All results saved in: {output_dir}[/green]")
    except Exception as e:
        console.print(f"[red]Error during processing: {str(e)}[/red]")
