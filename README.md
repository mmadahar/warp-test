# Warp Test

A Python utility project for generating and managing Excel workbooks with random data, as well as file system operations.

## Project Overview

This project provides tools for:
- Generating large numbers of Excel workbooks with random data
- File system operations (listing files and folders)
- Verification of Excel workbook contents
- CSV data management utilities

## Structure

### Source Files (`src/`)

- `utils.py`: Core utility functions for directory creation and CSV operations
- `get_files.py`: Functions for retrieving and managing file lists
- `get_folders.py`: Functions for retrieving and managing folder lists

### Test Files (`test/`)

- `excel_generator.py`: Main script for generating Excel workbooks with random data
- `verify_workbook.py`: Tool for verifying the structure and content of generated Excel workbooks

## Installation

1. Ensure you have Python 3.11+ installed
2. Clone this repository
3. Install dependencies:

```bash
# Using uv (recommended)
uv add openpyxl pandas faker

# Or using pip
pip install openpyxl pandas faker
```

## Usage

### Generating Excel Workbooks

The `excel_generator.py` script generates Excel workbooks with random data. Each workbook contains 10 worksheets with random content.

```bash
# Basic usage (generates 1000 workbooks in the specified directory)
PYTHONPATH=. uv run test/excel_generator.py --write_folder=./output_dir

# Specify a custom number of workbooks
PYTHONPATH=. uv run test/excel_generator.py --write_folder=./output_dir --num_workbooks=50
```

### Verifying Excel Workbooks

The `verify_workbook.py` script checks the structure and content of generated workbooks.

```bash
# Verify a specific workbook
PYTHONPATH=. uv run test/verify_workbook.py path/to/workbook.xlsx
```

### Utility Functions

The project includes utility functions for:

- Creating directories if they don't exist
- Saving data to CSV files
- Reading data from CSV files

Example usage in Python:

```python
from src.utils import create_directory_if_not_exists, save_to_csv

# Create a directory
output_dir = create_directory_if_not_exists("./data/output")

# Save data to CSV
items = ["file1.txt", "file2.txt", "file3.txt"]
csv_path = save_to_csv(items, "files.csv", "filename")
```

## Excel Workbook Structure

Each generated workbook:
- Contains exactly 10 worksheets
- Has random worksheet names
- Contains random data in various formats:
  - Text values
  - Numeric values
  - Dates
  - Sentences and phrases
- Varies in structure (rows and columns) across worksheets

## Dependencies

- Python 3.11+
- openpyxl: Excel file manipulation
- pandas: Data handling and CSV operations
- Faker: Generation of realistic random data

## License

This project is made available under the terms of the MIT license.

