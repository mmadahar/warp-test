#!/usr/bin/env python3
"""
Simple script to display the schema of a Delta table using polars.
"""

import polars as pl

def get_delta_schema(delta_path: str) -> None:
    """
    Retrieve and print the schema of a Delta table.
    
    Args:
        delta_path: Path to the Delta table
    """
    try:
        # Read just the schema (no data) from the Delta table
        df = pl.scan_delta(delta_path)
        
        # Collect schema using the recommended method to avoid performance warning
        schema = df.collect_schema()
        
        # Print schema in a readable format
        print(f"Schema for Delta table at {delta_path}:")
        for field in schema.items():
            print(f"  {field[0]}: {field[1]}")
            
    except Exception as e:
        print(f"Error reading Delta table: {e}")

if __name__ == "__main__":
    # Path to the Delta table
    delta_path = "./data/excel"
    
    # Get and display the schema
    get_delta_schema(delta_path)

