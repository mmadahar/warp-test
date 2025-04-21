#!/usr/bin/env python3
"""
Test script for demonstrating the process_excel function.

This script creates a sample Excel file with multiple sheets,
processes it with the process_excel function, and verifies the results.
"""
import os
import sys
import pandas as pd
import numpy as np
from pathlib import Path

# Add the src directory to the Python path
sys.path.append(str(Path(__file__).parent.parent / "src"))

# Import the process_excel function
from read_excel import process_excel

# Directory setup
TEST_DIR = Path(__file__).parent
DATA_DIR = TEST_DIR / "data"
EXCEL_FILE = DATA_DIR / "sample_test.xlsx"
DELTA_DIR = DATA_DIR / "delta_output"

def create_test_excel():
    """Create a test Excel file with multiple sheets."""
    # Create directories if they don't exist
    DATA_DIR.mkdir(exist_ok=True)
    DELTA_DIR.mkdir(exist_ok=True)
    
    # Create sample DataFrames for different sheets
    # Sheet 1: Employee data
    employees = pd.DataFrame({
        'ID': [101, 102, 103, 104, 105],
        'Name': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Brown', 'Charlie Davis'],
        'Department': ['Engineering', 'Marketing', 'HR', 'Finance', 'IT'],
        'Salary': [85000, 75000, 65000, 78000, 88000],
        'Hire Date': ['2020-05-15', '2019-07-22', '2021-03-10', '2018-11-30', '2022-01-05']
    })
    
    # Sheet 2: Product data
    products = pd.DataFrame({
        'Product ID': ['P001', 'P002', 'P003', 'P004', 'P005'],
        'Name': ['Widget A', 'Widget B', 'Gadget X', 'Gadget Y', 'Tool Z'],
        'Category': ['Widgets', 'Widgets', 'Gadgets', 'Gadgets', 'Tools'],
        'Price': [19.99, 24.99, 49.99, 59.99, 99.99],
        'In Stock': ['Yes', 'No', 'Yes', 'Yes', 'No']
    })
    
    # Sheet 3: Sales data
    sales = pd.DataFrame({
        'Date': pd.date_range(start='2023-01-01', periods=10, freq='D'),
        'Product ID': np.random.choice(['P001', 'P002', 'P003', 'P004', 'P005'], size=10),
        'Quantity': np.random.randint(1, 50, size=10),
        'Total': np.random.uniform(10, 1000, size=10).round(2),
        'Customer ID': [f'C{i:03d}' for i in np.random.randint(1, 100, size=10)]
    })
    
    # Write DataFrames to Excel file with different sheets
    with pd.ExcelWriter(EXCEL_FILE) as writer:
        employees.to_excel(writer, sheet_name='Employees', index=False)
        products.to_excel(writer, sheet_name='Products', index=False)
        sales.to_excel(writer, sheet_name='Sales', index=False)
    
    print(f"Created test Excel file at: {EXCEL_FILE}")
    return EXCEL_FILE

def test_process_excel():
    """Test the process_excel function with the sample Excel file."""
    # Create the test Excel file
    excel_file = create_test_excel()
    
    # Process the Excel file and write to Delta
    print("\nProcessing Excel file and writing to Delta format...")
    success, dfs = process_excel(
        file=str(excel_file),
        delta_path=str(DELTA_DIR),
        partition_col='worksheet',
        engine=None  # Let pandas choose the appropriate engine
    )
    
    # Print results
    print(f"\nProcess completed with success: {success}")
    print(f"Number of sheets processed: {len(dfs)}")
    
    # Print information about each sheet
    for sheet_name, df in dfs.items():
        print(f"\nSheet: {sheet_name}")
        print(f"Shape: {df.shape}")
        print(f"Columns: {df.columns.tolist()}")
        print(f"First 3 rows:")
        print(df.head(3))
        print("-" * 50)
    
    # Verify Delta output directory
    delta_files = list(DELTA_DIR.glob("**/*"))
    print(f"\nDelta output directory: {DELTA_DIR}")
    print(f"Number of files/directories in Delta output: {len(delta_files)}")
    print("Delta partition structure:")
    for f in delta_files[:10]:  # Show first 10 files/directories
        if f.is_dir():
            print(f"  DIR: {f.relative_to(DELTA_DIR)}")
        else:
            print(f"  FILE: {f.relative_to(DELTA_DIR)}")
    
    if len(delta_files) > 10:
        print(f"  ... and {len(delta_files) - 10} more files/directories")
    
    return success

if __name__ == "__main__":
    # Ensure the script can be run from anywhere
    os.chdir(Path(__file__).parent.parent)
    
    # Run the test
    print("Starting process_excel test...")
    result = test_process_excel()
    
    # Exit with appropriate status code
    sys.exit(0 if result else 1)

