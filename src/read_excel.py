#!/usr/bin/env python3
"""
Excel file reader using pandas.

This module provides the core functionality to read Excel files
and convert sheets to pandas DataFrames.
"""

import os
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
