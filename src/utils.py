"""
Utility functions for the file system operations.
"""
from pathlib import Path
from typing import Union, List
import pandas as pd

def create_directory_if_not_exists(directory_path: Union[str, Path]) -> Path:
    """
    Create a directory if it doesn't exist.
    
    Args:
        directory_path: Path to the directory to create (string or Path object)
        
    Returns:
        Path object of the created or existing directory
    """
    # Convert to Path object if it's a string
    path = Path(directory_path)
    
    # Create directory if it doesn't exist
    if not path.exists():
        path.mkdir(parents=True)
        print(f"Created directory: {path}")
    
    return path


def save_to_csv(items: List[str], output_file: str = "items.csv", column_name: str = "path"):
    """
    Save a list of paths (files or folders) to a CSV file in the data directory.
    
    Args:
        items: List of paths to save
        output_file: Name of the output file (default: 'items.csv')
        column_name: Name of the column in CSV (default: 'path')
    
    Returns:
        Path to the saved CSV file
    """
    # Create data directory if it doesn't exist
    data_dir = create_directory_if_not_exists("./data")
    
    # Create DataFrame and save to CSV
    df = pd.DataFrame(items, columns=[column_name])
    csv_path = data_dir / output_file
    df.to_csv(csv_path, index=False)
    
    # Determine item type for message (files or folders)
    item_type = "folders" if "folder" in column_name.lower() else "files"
    print(f"Saved {len(items)} {item_type} to {csv_path}")
    
    return csv_path


def read_from_csv(csv_path: str, column_name: str = None) -> List[str]:
    """
    Read paths from a CSV file.
    
    Args:
        csv_path: Path to the CSV file containing paths
        column_name: Name of the column to read (default: None, will try to auto-detect)
        
    Returns:
        List of paths from the CSV
    """
    try:
        df = pd.read_csv(csv_path)
        
        # If column name is provided and exists, use it
        if column_name and column_name in df.columns:
            return df[column_name].tolist()
        
        # Try to detect column name from common patterns
        for name in ['path', 'folder_path', 'file_path', 'folder', 'file']:
            if name in df.columns:
                return df[name].tolist()
        
        # Default to first column if no matching column found
        column = df.columns[0]
        print(f"No path column detected, using first column: '{column}'")
        return df[column].tolist()
        
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        return []
