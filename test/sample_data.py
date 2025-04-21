#!/usr/bin/env python3
"""
Simple script to sample 10 rows from a Delta table using polars.
"""

import polars as pl
import random

def sample_delta_data(delta_path: str, num_rows: int = 10) -> None:
    """
    Sample and print rows from a Delta table.
    
    Args:
        delta_path: Path to the Delta table
        num_rows: Number of rows to sample (default: 10)
    """
    try:
        # Read the Delta table
        df = pl.scan_delta(delta_path).collect()
        
        # Get total number of rows
        total_rows = df.height
        print(f"Total rows in Delta table: {total_rows}")
        
        # Sample rows (use min in case there are fewer than num_rows)
        sample_size = min(num_rows, total_rows)
        sampled_df = df.sample(sample_size, seed=42)
        
        # Print the sampled rows
        print(f"\n{sample_size} sample rows from Delta table at {delta_path}:")
        print(sampled_df)
            
    except Exception as e:
        print(f"Error reading Delta table: {e}")

if __name__ == "__main__":
    # Path to the Delta table
    delta_path = "./data/excel"
    
    # Sample and display data
    sample_delta_data(delta_path)

