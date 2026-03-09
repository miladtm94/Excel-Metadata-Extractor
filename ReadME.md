# Excel Metadata Extractor

A lightweight Python tool to extract sheet names and column headers from an Excel file (.xlsx) and export them into a structured table. Useful for documenting the structure of Excel workbooks, especially when dealing with many sheets or complex headers.

## Features

- Reads all sheets from an Excel file without loading the actual data (only headers are fetched).
- Handles empty sheets gracefully (represented by blank cells in the output).
- Produces a clean Excel output where each column corresponds to a sheet, and rows list the column headers of that sheet.
- Command-line interface with input and output file arguments.
- Informative logging and error handling.

## Requirements

- Python 3.8 or higher
- pandas
- openpyxl (for reading/writing .xlsx files)

Install dependencies with:

```bash
pip install pandas openpyxl
```

## Installation

Clone the repository or download the script `excel_metadata_extractor.py`. No additional installation is required.

## Usage

Run the script from the command line, providing the path to your input Excel file:

```bash
python excel_metadata_extractor.py path/to/your_file.xlsx
```

By default, the output is saved as `Result.xlsx` in the current directory. You can specify a different output file using the `-o` or `--output` option:

```bash
python excel_metadata_extractor.py data.xlsx -o metadata_output.xlsx
```

### Example

Given an Excel file `Data.xlsx` with three sheets:

- **Sheet1** headers: `['Header A', 'Header B', 'Header C', 'Header D']`
- **Sheet2** headers: `['Header a', 'Header b', 'Header c', 'Header d']`
- **Sheet3** headers: `['Header 1', 'Header 2', 'Header 3', 'Header 4', 'Header 5', 'Header 6', 'Header 7', 'Header 8']`
- **Sheet4** (empty)
- **Sheet5** headers: `['Header 10', 'Header 11', 'Header 12', 'Header 13', 'Header 14', 'Header 15', 'Header 16', 'Header 17']`

Running the command:

```bash
python excel_metadata_extractor.py Data.xlsx
```

Produces `Result.xlsx` with the following structure:

| Sheet1     | Sheet2     | Sheet3   | Sheet4 | Sheet5     |
|------------|------------|----------|--------|------------|
| Header A   | Header a   | Header 1 |        | Header 10  |
| Header B   | Header b   | Header 2 |        | Header 11  |
| Header C   | Header c   | Header 3 |        | Header 12  |
| Header D   | Header d   | Header 4 |        | Header 13  |
|            |            | Header 5 |        | Header 14  |
|            |            | Header 6 |        | Header 15  |
|            |            | Header 7 |        | Header 16  |
|            |            | Header 8 |        | Header 17  |

## How It Works

1. The script reads the Excel file and lists all sheet names.
2. For each sheet, it loads **only the header row** (`nrows=0` in pandas) to obtain the column names.
3. It builds a dictionary mapping sheet names to their headers.
4. It creates a new DataFrame where each column is a sheet, and rows contain the headers (padded with `NaN` to align columns of different lengths).
5. The DataFrame is saved to a new Excel file without an index.

## Contributing

Contributions, issues, and feature requests are welcome. Feel free to check the [issues page](https://github.com/yourusername/excel-metadata-extractor/issues) if you have ideas for improvement.

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Author

Milad Tatar Mamaghani – original concept (2022)  
Refactored with best practices for GitHub.

