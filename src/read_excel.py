#!/usr/bin/env python3
"""
Excel file reader using pandas.

This module provides the core functionality to read Excel files
and convert sheets to pandas DataFrames.
"""

import os
import pandas as pd
from deltalake import write_deltalake
def read_excel(file, engine=None):
    """
    Read all sheets from an Excel file and return them as a dictionary of DataFrames.
    
    Args:
        file (str): Path to the Excel file
        engine (str, optional): Excel engine to use. Options include:
            'openpyxl' (default) for .xlsx files
            'xlrd' for .xls files
            'odf' for .ods files
            'pyxlsb' for .xlsb files
            'calamine' for fast reading of Excel files
            None to let pandas decide based on file extension
            
    Returns:
        tuple: (dict, str) where:
            - dict: Dictionary of pandas DataFrames where key = sheet name
            - str: Absolute path to the Excel file
    """
    # Read all sheets with no header and convert all data to strings
    dfs = pd.read_excel(file, sheet_name=None, header=None, engine=engine, dtype=str)
    
    # Get absolute path to the file
    absolute_file_path = os.path.abspath(file)
    
    # Return both the DataFrames dictionary and the absolute file path
    return dfs, absolute_file_path


def process_excel(file, delta_path, partition_col, engine=None):
    """
    Read all sheets from an Excel file, convert to DataFrames, and write to a Delta table.
    
    This function combines the functionality of read_excel and write_dfs_to_delta.
    
    Args:
        file (str): Path to the Excel file
        delta_path (str): Path for the Delta table
        partition_col (str): Column name to use for partitioning
        engine (str, optional): Excel engine to use
            
    Returns:
        tuple: (bool, dict) - Success status and dictionary of pandas DataFrames
    """
    # Read all sheets from Excel using the existing read_excel function
    dfs, file_path = read_excel(file, engine)
    
    # Extract file metadata
    full_path = file_path  # Already an absolute path from read_excel
    file_name = os.path.splitext(os.path.basename(full_path))[0]
    file_ext = os.path.splitext(os.path.basename(full_path))[1][1:]  # Remove the leading dot
    
    # Process each DataFrame, adding metadata columns
    success = True
    try:
        for sheet_name, df in dfs.items():
            # Create a copy of the DataFrame
            df = df.copy()
            
            # Add metadata columns
            df.insert(0, 'path', full_path)
            df.insert(1, 'name', file_name)
            df.insert(2, 'ext', file_ext)
            df.insert(3, 'worksheet', sheet_name)
            df.insert(4, 'row', range(1, len(df) + 1))
            
            write_deltalake(
                delta_path,
                df,
                mode="append",
                partition_by=[partition_col],
                schema_mode="merge"
            )
            
            print(f"Wrote sheet {sheet_name} to Delta table at {delta_path}")
    except Exception as e:
        print(f"Error writing to Delta table: {e}")
        success = False
    
    return success, dfs
