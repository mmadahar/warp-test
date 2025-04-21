# Excel to Delta Table Converter

A Python utility project for reading Excel files and converting them to Delta tables with a unified schema. The project supports multiple worksheets and maintains data lineage through metadata columns.

## Project Overview

This project provides tools for:
- Reading Excel files and converting them to Delta tables with a unified schema
- Handling multiple worksheets from a single file
- Maintaining data lineage by including metadata (file path, name, extension, worksheet name, row number)
- Supporting various Excel engines (openpyxl, xlrd, odf, pyxlsb, calamine)
- Generating test Excel workbooks with random data for testing purposes

## Structure

### Source Files (`src/`)

- `excel.py`: Core module for reading Excel files and converting them to Delta tables
- `excel_safe.py`: Enhanced version of excel.py that handles unexpected kernel crashes and OOM errors
- `utils.py`: Utility functions for directory creation and CSV operations
- `get_files.py`: Functions for retrieving and managing file lists
- `get_folders.py`: Functions for retrieving and managing folder lists

### Test Files (`test/`)

- `excel_generator.py`: Main script for generating Excel workbooks with random data
- `verify_workbook.py`: Tool for verifying the structure and content of generated Excel workbooks
- `get_schema.py`: Tool for displaying the schema of the generated Delta tables
- `sample_data.py`: Utility for sampling data from the Delta tables

## Installation

1. Ensure you have Python 3.11+ installed
2. Clone this repository
3. Install dependencies:

```bash
# Using uv (recommended)
uv add openpyxl pandas faker deltalake polars

# Or using pip
pip install openpyxl pandas faker deltalake polars
```

## Usage

### Converting Excel Files to Delta Tables

The main functionality allows you to read Excel files and convert their content to Delta tables with a unified schema:

```bash
# Basic usage
PYTHONPATH=. uv run src/excel.py path/to/excel_file.xlsx

# Specify a custom Excel engine
PYTHONPATH=. uv run src/excel.py path/to/excel_file.xlsx openpyxl

# Specify a custom Delta table path
PYTHONPATH=. uv run src/excel.py path/to/excel_file.xlsx openpyxl ./data/custom_delta
```

You can also use the functionality programmatically:

```python
from src.excel import read_and_process_excel

# Read an Excel file and write to Delta table
success, dataframes = read_and_process_excel(
    "path/to/excel_file.xlsx", 
    "./data/excel", 
    engine="openpyxl"
)

# Check if the operation was successful
if success:
    print("Conversion completed successfully!")
```

### Examining Delta Tables

You can examine the schema and sample data from the Delta tables:

```bash
# Display the schema of the Delta table
PYTHONPATH=. uv run test/get_schema.py

# Sample data from the Delta table
PYTHONPATH=. uv run test/sample_data.py
```

### Generating Test Excel Workbooks

The `excel_generator.py` script generates Excel workbooks with random data for testing. Each workbook contains 10 worksheets with random content.

```bash
# Basic usage (generates 1000 workbooks in the specified directory)
PYTHONPATH=. uv run test/excel_generator.py --write_folder=./test/excel

# Specify a custom number of workbooks
PYTHONPATH=. uv run test/excel_generator.py --write_folder=./test/excel --num_workbooks=50
```

### Complete End-to-End Example

Here's a complete example showing how to generate test data and convert it to Delta format:

```bash
# 1. Generate test Excel files
PYTHONPATH=. uv run test/excel_generator.py --write_folder=./test/excel --num_workbooks=10

# 2. Convert one of the Excel files to Delta format
PYTHONPATH=. uv run src/excel.py ./test/excel/workbook_0001.xlsx

# 3. Examine the schema of the resulting Delta table
PYTHONPATH=. uv run test/get_schema.py

# 4. Sample data from the Delta table
PYTHONPATH=. uv run test/sample_data.py
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

## Delta Table Schema

When Excel files are converted to Delta tables, the following schema is applied:

- `path`: Absolute path to the source Excel file
- `name`: Name of the source Excel file (without extension)
- `ext`: Extension of the source Excel file
- `worksheet`: Name of the worksheet in the Excel file
- `row`: Row number in the original Excel file
- Additional columns: All data from the original Excel file with auto-generated column names

The Delta tables are partitioned by `worksheet` to optimize querying by worksheet name.

## Excel Workbook Structure

Each generated test workbook:
- Contains exactly 10 worksheets
- Has random worksheet names
- Contains random data in various formats:
  - Text values
  - Numeric values
  - Dates
  - Sentences and phrases
- Varies in structure (rows and columns) across worksheets

## Technical Details

### Delta Table Configuration

The Delta tables are written with the following configuration:
- Mode: `append` - Data is appended to existing tables
- Partition by: `worksheet` - Data is partitioned by worksheet name
- Schema mode: `merge` - Schema is merged with existing schema if table already exists

### Excel Engines Support

The project supports multiple Excel engines:
- `openpyxl` (default) for .xlsx files
- `xlrd` for .xls files 
- `odf` for .ods files
- `pyxlsb` for .xlsb files
- `calamine` for fast reading of Excel files

## Dependencies

- Python 3.11+
- openpyxl: Excel file manipulation
- pandas: Data handling and CSV operations
- deltalake: Delta table operations
- polars: Fast DataFrame operations and Delta table querying
- Faker: Generation of realistic random data

## Error Handling

### excel_safe.py

The `excel_safe.py` module was specifically created to handle unexpected kernel crashes and Out of Memory (OOM) errors that can occur when processing large Excel files. It includes:

- Robust error handling for kernel crashes during Excel file processing
- Memory management techniques to prevent OOM errors when dealing with large files
- Graceful fallback mechanisms that allow processing to continue even after errors
- Detailed logging to help diagnose issues with problematic files
- Automatic retries with different Excel engines when a particular engine fails

Example usage:

```python
from src.excel_safe import read_and_process_excel_safely

# Process an Excel file with safety mechanisms enabled
success, dataframes = read_and_process_excel_safely(
    "path/to/large_excel_file.xlsx", 
    "./data/excel",
    max_memory_mb=1024,  # Set memory limit
    retry_count=3        # Number of retry attempts
)
```

```bash
# Command line usage with safety mechanisms
PYTHONPATH=. uv run src/excel_safe.py path/to/large_excel_file.xlsx
```

## License

This project is made available under the terms of the MIT license.

