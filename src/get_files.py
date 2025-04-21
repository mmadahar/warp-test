from pathlib import Path
from typing import Union, List, Set
import os
import pathspec
import concurrent.futures
import time
import argparse
import pandas as pd
from utils import create_directory_if_not_exists, save_to_csv, read_from_csv
def load_gitignore() -> pathspec.PathSpec:
    """
    Load the .gitignore file and create a PathSpec object for pattern matching.
    
    Returns:
        PathSpec object with patterns from .gitignore
    """
    gitignore_path = Path(".gitignore")
    patterns = []
    
    if gitignore_path.exists():
        with open(gitignore_path, "r") as f:
            for line in f:
                line = line.strip()
                # Skip empty lines and comments
                if line and not line.startswith("#"):
                    patterns.append(line)
    
    return pathspec.PathSpec.from_lines(pathspec.patterns.GitWildMatchPattern, patterns)

def process_folder(folder: str, extensions: List[str], gitignore_spec: pathspec.PathSpec) -> Set[str]:
    """Process a single folder and return matching files."""
    print(f"Processing folder: {folder}...")
    start_time = time.time()
    
    path = Path(folder)
    if not path.exists():
        print(f"  Folder not found: {folder}")
        return set()
        
    matching_files: Set[str] = set()
    
    # Recursively search for files
    for ext in extensions:
        for p in path.rglob(f"*{ext}"):
            file_path = str(p)
            
            # Get the relative path for gitignore matching
            try:
                # Convert to relative path if it's not already
                rel_path = os.path.relpath(file_path)
                
                # Only add the file if it's not ignored by gitignore patterns
                if not gitignore_spec.match_file(rel_path):
                    matching_files.add(file_path)
            except ValueError:
                # This can happen if files are on different drives
                matching_files.add(file_path)
    
    elapsed = time.time() - start_time
    print(f"  Completed {folder}: found {len(matching_files)} files in {elapsed:.2f} seconds")
    return matching_files


def get_files(folder: Union[str, List[str]], ext: Union[str, List[str]],
              save_csv: bool = False, csv_filename: str = "files.csv") -> List[str]:
    """
    Find files with specified extensions in given folders.
    
    Args:
        folder: Single path or list of paths to search
        ext: Single extension or list of extensions to match
        save_csv: Whether to save results to CSV (default: False)
        csv_filename: Name of CSV file to save results (default: "files.csv")
    
    Returns:
        List of matching file paths as strings
    """
    # Convert inputs to lists if they're strings
    folders = [folder] if isinstance(folder, str) else folder
    extensions = [ext] if isinstance(ext, str) else ext
    
    # Ensure extensions start with '.'
    extensions = [e if e.startswith(".") else f".{e}" for e in extensions]
    
    # Load gitignore patterns
    gitignore_spec = load_gitignore()
    
    # Use a set to prevent duplicates
    matching_files: Set[str] = set()
    
    # Handle based on number of folders
    if len(folders) == 1:
        # For a single folder, just process directly
        matching_files = process_folder(folders[0], extensions, gitignore_spec)
    else:
        # For multiple folders, use parallel processing
        print(f"Processing {len(folders)} folders in parallel...")
        start_time = time.time()
        
        with concurrent.futures.ThreadPoolExecutor() as executor:
            # Map the function across all folders
            results = executor.map(
                lambda f: process_folder(f, extensions, gitignore_spec),
                folders
            )
            
            # Collect all results
            for folder_files in results:
                matching_files.update(folder_files)
                
        elapsed = time.time() - start_time
        print(f"All folders processed: found {len(matching_files)} unique files in {elapsed:.2f} seconds")
    
    result = sorted(matching_files)
    
    # Save to CSV if requested
    if save_csv and result:
        save_to_csv(result, csv_filename, column_name="file_path")
    
    return result

if __name__ == '__main__':
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description="Find files with specified extensions in given folders")
    parser.add_argument("--paths", nargs="+", help="One or more paths to search")
    parser.add_argument("--input-csv", help="Path to CSV file containing folder paths to search")
    parser.add_argument("--extensions", nargs="+", help="One or more file extensions to match")
    parser.add_argument("--output", default="files.csv", help="Output CSV filename (default: files.csv)")
    args = parser.parse_args()
    
    # Default test paths and extensions
    search_paths = ["."]
    
    # Process input sources for folder paths
    if args.paths:
        search_paths = args.paths
        print(f"Using command line paths: {search_paths}")
    elif args.input_csv:
        csv_folders = read_from_csv(args.input_csv, column_name="folder_path")
        if csv_folders:
            search_paths = csv_folders
            print(f"Using paths from CSV file: {args.input_csv}")
    
    # If extensions are provided via CLI, use those
    if args.extensions:
        # Run with the specified parameters
        print(f"\nFinding files with extensions: {args.extensions}")
        found_files = get_files(search_paths, args.extensions,
                                save_csv=True, csv_filename=args.output)
        # Print results
        for file in found_files:
            print(f"  {file}")
        # CSV already saved by the function
        
    else:
        # Default test cases
        # Test case 1: Single extension
        print("\nFinding all Python files:")
        python_files = get_files(search_paths, "py")
        for file in python_files:
            print(f"  {file}")
        
        # Test case 2: Multiple extensions
        print("\nFinding Python and TOML files:")
        code_files = get_files(search_paths, ["py", "toml"])
        for file in code_files:
            print(f"  {file}")
        
        # If using default current directory, also do the multi-folder test
        if search_paths == ["."]:
            # Test case 3: Multiple folders
            print("\nFinding Python files in src and root:")
            multi_folder_files = get_files(["src", "."], "py",
                                           save_csv=True, csv_filename=args.output)
            for file in multi_folder_files:
                print(f"  {file}")
            # CSV already saved by the function
