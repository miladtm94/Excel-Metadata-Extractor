#!/usr/bin/env python3
"""
Excel Metadata Extractor

Reads an Excel file (.xlsx) and extracts sheet names along with their column headers.
Produces a new Excel file where each column represents a sheet and rows contain the
headers of that sheet. Empty sheets are represented by blank cells.

Default input file: Data.xlsx
Default output file: Result.xlsx
"""

import argparse
import logging
import sys
from pathlib import Path

import pandas as pd


def setup_logging() -> None:
    """Configure basic logging to console."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def extract_headers(file_path: Path) -> dict[str, pd.Index]:
    """
    Extract column headers from each sheet of an Excel file.

    Args:
        file_path: Path to the Excel file.

    Returns:
        A dictionary mapping sheet names to their column headers (as pandas Index).
        Sheets that are completely empty will have an empty Index.

    Raises:
        FileNotFoundError: If the Excel file does not exist.
        ValueError: If the file cannot be read as an Excel file.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        # Read only the header row (nrows=0) from every sheet
        xl = pd.ExcelFile(file_path)
        headers = {}
        for sheet_name in xl.sheet_names:
            # Use nrows=0 to fetch only the column headers, no data rows
            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=0)
            headers[sheet_name] = df.columns
        return headers
    except Exception as e:
        raise ValueError(f"Failed to read Excel file: {e}") from e


def create_metadata_dataframe(headers: dict[str, pd.Index]) -> pd.DataFrame:
    """
    Convert the headers dictionary into a DataFrame suitable for export.

    The resulting DataFrame has sheet names as column headers. Each column contains
    the list of headers for that sheet, padded with NaN to make all columns the same
    length.

    Args:
        headers: Dictionary mapping sheet names to column headers.

    Returns:
        A pandas DataFrame where columns are sheet names and rows are header values.
    """
    # Determine the maximum number of headers among all sheets
    max_len = max((len(h) for h in headers.values()), default=0)

    # Build a dictionary of lists, each of length max_len, padded with None
    data = {}
    for sheet, cols in headers.items():
        # Convert Index to list, then pad with None
        col_list = list(cols)
        padded = col_list + [None] * (max_len - len(col_list))
        data[sheet] = padded

    # Create DataFrame; orientation='index' would put sheets as rows, so we transpose
    df = pd.DataFrame(data)
    return df


def save_metadata(df: pd.DataFrame, output_path: Path) -> None:
    """
    Save the metadata DataFrame to an Excel file.

    Args:
        df: DataFrame containing the metadata.
        output_path: Path where the Excel file will be written.
    """
    try:
        df.to_excel(output_path, index=False)
        logging.info(f"Metadata successfully written to {output_path}")
    except Exception as e:
        raise IOError(f"Failed to write output file: {e}") from e


def parse_arguments() -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Extract sheet names and column headers from an Excel file."
    )
    parser.add_argument(
        "input_file",
        nargs="?",  # Makes it optional
        default="Data.xlsx",
        help="Path to the input Excel file (default: Data.xlsx).",
    )
    parser.add_argument(
        "-o", "--output",
        default="Result.xlsx",
        help="Path for the output Excel file (default: Result.xlsx).",
    )
    return parser.parse_args()


def main() -> None:
    """Main execution flow."""
    setup_logging()
    args = parse_arguments()

    input_path = Path(args.input_file)
    output_path = Path(args.output)

    try:
        logging.info(f"Reading headers from {input_path}")
        headers = extract_headers(input_path)

        if not headers:
            logging.warning("No sheets found in the Excel file.")
            return

        logging.info(f"Found {len(headers)} sheet(s). Creating metadata table.")
        df = create_metadata_dataframe(headers)

        logging.info(f"Saving metadata to {output_path}")
        save_metadata(df, output_path)

        # Print a summary to console
        print("\nExtracted sheet names:")
        for i, sheet in enumerate(headers.keys(), start=1):
            print(f"  {i}. {sheet} ({len(headers[sheet])} headers)")

    except (FileNotFoundError, ValueError, IOError) as e:
        logging.error(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()