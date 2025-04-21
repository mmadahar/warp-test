from pathlib import Path
from typing import Union, List, Set
import argparse
import pandas as pd
from utils import create_directory_if_not_exists, save_to_csv
def get_folders(folders: Union[str, List[str]], glob_pattern: Union[str, List[str]], 
               save_csv: bool = False, csv_filename: str = "folders.csv") -> List[str]:
    """
    Find folders matching specified glob patterns in given directories (non-recursively).
    
    Args:
        folders: Single path or list of paths to search
        glob_pattern: Single glob pattern or list of patterns to match (e.g. "test*", "data-*")
        save_csv: Whether to save results to CSV (default: False)
        csv_filename: Name of CSV file to save results (default: "folders.csv")
    
    Returns:
        List of matching folder paths as strings, sorted
    """
    # Convert inputs to lists if they're strings
    folder_list = [folders] if isinstance(folders, str) else folders
    patterns = [glob_pattern] if isinstance(glob_pattern, str) else glob_pattern
    
    # Use a set to prevent duplicates
    matching_folders: Set[str] = set()
    
    # Search each folder
    for folder in folder_list:
        path = Path(folder)
        if not path.exists() or not path.is_dir():
            continue
            
        # For each pattern, find matching folders (non-recursively)
        for pattern in patterns:
            for item in path.glob(pattern):
                if item.is_dir():
                    matching_folders.add(str(item))
    
    result = sorted(matching_folders)
    
    # Save to CSV if requested
    # Save to CSV if requested
    if save_csv and result:
        save_to_csv(result, csv_filename, column_name="folder_path")
    return result


if __name__ == '__main__':
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description="Find folders matching glob patterns")
    parser.add_argument("--paths", nargs="+", help="One or more paths to search (overrides default test paths)")
    parser.add_argument("--patterns", nargs="+", help="One or more glob patterns to match")
    parser.add_argument("--output", default="folders.csv", help="Output CSV filename (default: folders.csv)")
    args = parser.parse_args()
    
    # Default test paths
    test_paths = ["/home/matthew/python", "/home/matthew/Documents"]
    
    # If paths are provided via CLI, use those instead
    if args.paths:
        print(f"Using custom paths: {args.paths}")
        search_paths = args.paths
    else:
        search_paths = test_paths
    
    # If patterns are provided via CLI, use those instead of running default tests
    if args.patterns:
        print(f"\nFinding folders matching patterns: {args.patterns}")
        found_folders = get_folders(search_paths, args.patterns, 
                                     save_csv=True, csv_filename=args.output)
        
        # Print results
        for folder in found_folders:
            print(f"  {folder}")
        # CSV already saved by the function
        
    else:
        # Run default test cases
        # Test case 1: Single pattern
        print("\nFinding folders matching 'agent*':")
        agent_folders = get_folders(search_paths, "agent*")
        for folder in agent_folders:
            print(f"  {folder}")
        
        # Test case 2: Multiple patterns
        print("\nFinding folders matching multiple patterns:")
        multi_pattern_folders = get_folders(search_paths, ["agent*", "y*"])
        for folder in multi_pattern_folders:
            print(f"  {folder}")
        
        # Test case 3: Find all folders
        print("\nFinding all folders with '*' pattern:")
        all_folders = get_folders(search_paths, "*", 
                                  save_csv=True, csv_filename=args.output)
        for folder in all_folders:
            print(f"  {folder}")
        # CSV already saved by the function
