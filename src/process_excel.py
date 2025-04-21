#!/usr/bin/env python3
"""
Excel file reader using pandas.

This module provides the core functionality to read Excel files
and convert sheets to pandas DataFrames.
"""
import sys
import subprocess
import os
import pandas as pd
from deltalake import write_deltalake
from read_excel import read_excel

def safe_read_excel(file):
    """
    Safely read Excel file by first trying calamine engine in a subprocess,
    falling back to default engine if it fails.
    
    Args:
        file (str): Path to the Excel file
            
    Returns:
        tuple: (dict, str) where:
            - dict: Dictionary of pandas DataFrames where key = sheet name
            - str: Absolute path to the Excel file
    """
    try:
        # Try reading with calamine in a subprocess first
        cmd = [
            "uv", "run", "-q", "-c",
            f"import pandas as pd; pd.read_excel('{file}', sheet_name=None, header=None, engine='calamine', dtype=str)"
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)  # 5 min timeout
        
        if result.returncode == 0:
            print("Calamine subprocess check successful, reading with calamine...")
            return read_excel(file, engine="calamine")
            
    except subprocess.TimeoutExpired:
        print("Calamine engine timed out, falling back to default engine")
    except subprocess.CalledProcessError as e:
        print(f"Calamine engine failed: {e}, falling back to default engine")
    except Exception as e:
        print(f"Unexpected error with calamine: {e}, falling back to default engine")
    
    print("Using default engine...")
    return read_excel(file)
def write_dfs_to_delta(df_dict, base_path, partition_col, excel_file_path):
    """
    Write each DataFrame to Delta table with partitioning, append mode, and schema merging.
    Adds metadata columns from the Excel file path.
    
    Args:
        df_dict: Dictionary of pandas DataFrames where key = sheet name
        base_path: Base path for the Delta table
        partition_col: Column name to use for partitioning
        excel_file_path: Path to the Excel file to extract metadata from
        
    Returns:
        bool: True if all sheets were written successfully
    """
    # Process each DataFrame, adding sheet_name as a column
    # Extract file metadata
    full_path = os.path.abspath(excel_file_path)
    file_name = os.path.splitext(os.path.basename(full_path))[0]
    file_ext = os.path.splitext(os.path.basename(full_path))[1][1:]  # Remove the leading dot
    
    # Process each DataFrame, adding metadata columns
    for sheet_name, df in df_dict.items():
        # Create a copy of the DataFrame
        df = df.copy()
        
        # Add metadata columns
        df.insert(0, 'path', full_path)
        df.insert(1, 'name', file_name)
        df.insert(2, 'ext', file_ext)
        df.insert(3, 'worksheet', sheet_name)  # Formerly sheet_name
        df.insert(4, 'row', range(1, len(df) + 1))  # Formerly row_number
        
        write_deltalake(
            base_path,
            df,
            mode="append",
            partition_by=['worksheet'],  # Updated from sheet_name to worksheet
            schema_mode="merge"
        )
        
        print(f"Wrote sheet {sheet_name} to Delta table at {base_path}")
    
    return True


def safe_process_workbook(filepath, delta_path, partition_col):
    """
    Safely process an Excel workbook by trying different engines if needed.
    First attempts with calamine engine, then falls back to default if needed.
    
    Args:
        filepath (str): Path to the Excel file
        delta_path (str): Path for Delta table output
        partition_col (str): Column name to use for partitioning
        
    Returns:
        bool: True if processing was successful, False otherwise
    """
    try:
        # Try reading with safe_read_excel
        dfs, file_path = safe_read_excel(filepath)
        
        # Write to Delta table
        success = write_dfs_to_delta(dfs, delta_path, partition_col, file_path)
        return success
    except Exception as e:
        print(f"Failed to process workbook: {e}")
        return False


def read_write_excel(file, delta_path, partition_col, engine=None):
    """
    Read all sheets from an Excel file and write them to a Delta table.
    
    Args:
        file (str): Path to the Excel file
        delta_path (str): Path for the Delta table
        partition_col (str): Column name to use for partitioning
        engine (str, optional): Excel engine to use
            
    Returns:
        tuple: (bool, dict) - Success status and dictionary of pandas DataFrames
    """
    # Read all sheets from Excel
    dfs, file_path = read_excel(file, engine)
    
    # Write to a Delta table
    success = write_dfs_to_delta(dfs, delta_path, partition_col, file_path)
    return success, dfs


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Read Excel files and optionally write to a Delta table')
    parser.add_argument('file', help='Path to Excel file')
    parser.add_argument('--engine', help='Excel engine to use (openpyxl, xlrd, etc.)', default=None)
    parser.add_argument('--write-delta', action='store_true', 
                        help='Write DataFrames to a Delta table in ./data/excel')
    parser.add_argument('--delta-path', default='./data/excel',
                        help='Path where Delta tables will be written (default: ./data/excel)')
    parser.add_argument('--partition-col', default='0', 
                        help='Column name to use for partitioning (default: first column)')
    
    args = parser.parse_args()
    
    if args.write_delta:
        # Ensure directory exists
        os.makedirs(args.delta_path, exist_ok=True)
        
        if args.engine:
            # If engine is specified explicitly, use it directly
            dfs, file_path = read_excel(args.file, args.engine)
            success = write_dfs_to_delta(dfs, args.delta_path, args.partition_col, file_path)
            if success:
                print("Successfully wrote all Excel sheets to Delta tables")
        else:
            # Use the safe process approach with fallback
            success = safe_process_workbook(args.file, args.delta_path, args.partition_col)
            if success:
                print("Successfully processed workbook")
            else:
                print("Failed to process workbook")
    else:
        # Just read the Excel file
        if args.engine:
            dfs, _ = read_excel(args.file, args.engine)
        else:
            dfs, _ = safe_read_excel(args.file)
        
        # Print information about the sheets
        for sheet_name, df in dfs.items():
            print(f"\nSheet: {sheet_name}")
            print(f"Shape: {df.shape}")
            print(f"Columns: {df.columns.tolist()}")
            print("-" * 50)
