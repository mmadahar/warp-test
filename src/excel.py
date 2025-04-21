#!/usr/bin/env python3
"""
Excel file reader using pandas.

This module provides the core functionality to read Excel files
and convert sheets to pandas DataFrames.
"""

import os
import sys
import pandas as pd
from deltalake import write_deltalake
def read_and_process_excel(file, delta_path, engine=None):
    """
    Read all sheets from an Excel file, process them, and write to a Delta table.
    
    This function:
    1. Reads all sheets from an Excel file into DataFrames
    2. Adds metadata columns to each DataFrame (path, name, ext, worksheet, row)
    3. Writes each DataFrame to a Delta table partitioned by worksheet name
    
    All operations are wrapped in a try-except block to handle any errors that
    might occur during file reading, path operations, DataFrame manipulation,
    or Delta table writing.
    
    Args:
        file (str): Path to the Excel file
        delta_path (str): Path for the Delta table
        engine (str, optional): Excel engine to use. Options include:
            'openpyxl' (default) for .xlsx files
            'xlrd' for .xls files
            'odf' for .ods files
            'pyxlsb' for .xlsb files
            'calamine' for fast reading of Excel files
            None to let pandas decide based on file extension
            
    Returns:
        tuple: (bool, dict) - Success status and dictionary of pandas DataFrames
    """
    # Initialize variables
    success = False
    dfs = None
    
    try:
        # Read all sheets with no header and convert all data to strings
        dfs = pd.read_excel(file, sheet_name=None, header=None, engine=engine, dtype=str)
        
        # Get absolute path to the file
        absolute_file_path = os.path.abspath(file)
        
        # Extract file metadata
        file_name = os.path.splitext(os.path.basename(absolute_file_path))[0]
        file_ext = os.path.splitext(os.path.basename(absolute_file_path))[1][1:]  # Remove the leading dot
        
        # Process each DataFrame, adding metadata columns and writing to Delta table
        success = True
        for sheet_name, df in dfs.items():
            # Create a copy of the DataFrame
            df = df.copy()
            
            # Add metadata columns
            df.insert(0, 'path', absolute_file_path)
            df.insert(1, 'name', file_name)
            df.insert(2, 'ext', file_ext)
            df.insert(3, 'worksheet', sheet_name)
            df.insert(4, 'row', range(1, len(df) + 1))
            
            write_deltalake(
                delta_path,
                df,
                mode="append",
                partition_by=["worksheet"],
                schema_mode="merge"
            )
            
            print(f"Wrote sheet {sheet_name} to Delta table at {delta_path}")
    except Exception as e:
        error_source = "reading Excel file"
        if dfs is not None:
            error_source = "processing Excel data or writing to Delta table"
        print(f"Error {error_source}: {e}")
        success = False
    
    return success, dfs


if __name__ == "__main__":
    """
    Command-line interface for the read_excel.py script.
    
    Usage:
        uv run read_excel.py <file_path> [engine] [delta_path]
    
    Arguments:
        file_path (str): Path to the Excel file
        engine (str, optional): Excel engine to use (default: None)
        delta_path (str, optional): Path for the Delta table (default: './data/excel')
    
    Returns:
        Exit code 0 for success, 1 for failure
    """
    # Check for required file argument
    if len(sys.argv) < 2:
        print("Error: Excel file path is required")
        print("Usage: uv run read_excel.py <file_path> [engine] [delta_path]")
        sys.exit(1)
    
    # Parse command-line arguments
    file_path = sys.argv[1]
    
    # Get optional engine argument
    engine = None
    if len(sys.argv) >= 3:
        engine_arg = sys.argv[2]
        # Only set engine if it's not the delta path
        if not engine_arg.startswith('./') and not engine_arg.startswith('/'):
            engine = engine_arg
    
    # Get optional delta path argument
    delta_path = './data/excel'
    if len(sys.argv) >= 4:
        delta_path = sys.argv[3]
    elif len(sys.argv) == 3 and (sys.argv[2].startswith('./') or sys.argv[2].startswith('/')):
        # If the second argument looks like a path, use it as delta_path
        delta_path = sys.argv[2]
        engine = None
    
    print(f"Processing Excel file: {file_path}")
    print(f"Engine: {engine or 'default'}")
    print(f"Delta path: {delta_path}")
    
    # Call the main function
    success, _ = read_and_process_excel(file_path, delta_path, engine)
    
    # Return appropriate exit code
    if success:
        print("Processing completed successfully")
        sys.exit(0)
    else:
        print("Processing failed")
        sys.exit(1)
