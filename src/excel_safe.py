import subprocess
import pandas as pd
import os
import sys
import argparse
from pathlib import Path

# Get the directory where this script (excel_safe.py) is located
SCRIPT_DIR = Path(__file__).parent.resolve()
# Construct the path to read_excel.py relative to this script
# Assuming read_excel.py is in the same directory as this script
READ_EXCEL_SCRIPT_PATH = SCRIPT_DIR / "excel.py"

def process_excel_files_safely(file_list, errors_dir_path, delta_path):
    """
    Process Excel files safely with engine fallback:
    1. First attempt: engine="calamine"
    2. Second attempt: engine=None (default)
    3. If both fail: add to errors list
    
    Uses uv run for executing the excel.py script located in the same directory.
    Saves errors to a CSV in the specified errors directory.
    Delta table outputs will be saved to the specified delta path.
    """
    error_files = []

    # Ensure read_excel.py exists before processing any files
    if not READ_EXCEL_SCRIPT_PATH.is_file():
        print(f"Error: Required script '{READ_EXCEL_SCRIPT_PATH}' not found.", file=sys.stderr)
        # Add all potential files to error list as we cannot process them
        for filepath_str in file_list:
            filepath = Path(filepath_str)
            if not filepath.exists():
                 print(f"File not found: {filepath}. Logging as error.", file=sys.stderr)
            # Log the original string path provided
            error_files.append(filepath_str)
        print("Cannot proceed without excel.py. Exiting.", file=sys.stderr)
        # Skip writing CSV if read_excel.py is missing, just return the list
        return error_files # Early exit

    for filepath_str in file_list:
        # Resolve the path relative to the CWD *before* checking existence
        # This ensures relative paths passed on the command line work correctly
        filepath = Path(filepath_str).resolve()
        if not filepath.exists():
            print(f"File not found: {filepath}. Skipping.", file=sys.stderr)
            error_files.append(filepath_str) # Log the original string path
            continue

        print(f"Processing {filepath}...")

        # --- First attempt with calamine engine ---
        try:
            print(f"Attempting {filepath} with calamine engine...")
            # Pass the resolved, absolute path to the subprocess
            cmd_calamine = ['uv', 'run', str(READ_EXCEL_SCRIPT_PATH), str(filepath), 'calamine', str(delta_path)]
            print(f"  Executing: {' '.join(cmd_calamine)}")
            result_calamine = subprocess.run(
                cmd_calamine,
                capture_output=True,
                text=True,
                timeout=300,
                check=False # Don't raise error on non-zero exit
            )

            print(f"  calamine stdout:\n{result_calamine.stdout}")
            print(f"  calamine stderr:\n{result_calamine.stderr}")

            if result_calamine.returncode == 0:
                print(f"Successfully processed {filepath} with calamine engine")
                continue  # Success! Move to next file
            else:
                 print(f"Non-zero exit code ({result_calamine.returncode}) processing {filepath} with calamine engine.")

        except subprocess.TimeoutExpired:
            print(f"Timeout processing: {filepath} with calamine engine", file=sys.stderr)
        except Exception as e:
            print(f"Exception running subprocess for {filepath} with calamine engine: {e}", file=sys.stderr)

        # --- If first attempt failed - try with default engine ---
        print(f"Retrying {filepath} with default engine...")
        try:
             # Pass the resolved, absolute path to the subprocess
            cmd_default = ['uv', 'run', str(READ_EXCEL_SCRIPT_PATH), str(filepath), str(delta_path)] # No engine specified
            print(f"  Executing: {' '.join(cmd_default)}")
            result_default = subprocess.run(
                cmd_default,
                capture_output=True,
                text=True,
                timeout=300,
                check=False # Don't raise error on non-zero exit
            )

            print(f"  default stdout:\n{result_default.stdout}")
            print(f"  default stderr:\n{result_default.stderr}")

            if result_default.returncode == 0:
                print(f"Successfully processed {filepath} with default engine")
                continue  # Success! Move to next file
            else:
                print(f"Error in file with default engine: {filepath}. Exit code: {result_default.returncode}", file=sys.stderr)
                error_files.append(filepath_str) # Log the original string path

        except subprocess.TimeoutExpired:
            print(f"Timeout processing: {filepath} with default engine", file=sys.stderr)
            error_files.append(filepath_str) # Log the original string path
        except Exception as e:
            print(f"Exception running subprocess for {filepath} with default engine: {e}", file=sys.stderr)
            error_files.append(filepath_str) # Log the original string path

    # --- Create and save errors CSV ---
    if error_files:
        try:
            # Ensure the target directory exists (relative to CWD)
            errors_dir_path.mkdir(parents=True, exist_ok=True)
            errors_csv_path = errors_dir_path / 'errors.csv'
            # Use original file paths provided by user for the error report
            errors_df = pd.DataFrame({'path': error_files})
            errors_df.to_csv(errors_csv_path, index=False)
            print(f"Logged {len(error_files)} errors to {errors_csv_path.resolve()}")
        except Exception as e:
             # Use resolve() to show absolute path in error message
             print(f"Failed to write errors CSV to {errors_dir_path.resolve()}: {e}", file=sys.stderr)
    else:
        print("All files processed successfully!")

    return error_files

# --- Main execution block ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Process Excel files safely using 'excel.py' with engine fallback (calamine -> default).",
        formatter_class=argparse.RawTextHelpFormatter # Keep formatting in help text
        )
    parser.add_argument(
        "files",
        metavar="FILE",
        nargs='*', # Accept zero or more files
        help="Path(s) to Excel file(s) to process. Relative paths are resolved from the current working directory."
    )
    parser.add_argument(
        "--errors-dir",
        default="./data",
        help="Directory to save the 'errors.csv' file (default: ./data relative to current working directory)."
    )
    parser.add_argument(
        "--delta-path",
        default="./data/excel",
        help="Directory to save the Delta table (default: ./data/excel relative to current working directory)."
    )
    parser.add_argument(
        "--use-examples",
        action="store_true",
        help="If specified and no FILE arguments are given, use default example files (data1.xlsx, data2.xlsx, possibly_corrupt.xlsx) if they exist in the current working directory."
    )


    args = parser.parse_args()

    files_to_process = args.files
    # errors_directory is relative to CWD unless an absolute path is given
    errors_directory = Path(args.errors_dir) # Convert to Path object
    delta_path = args.delta_path

    # Use example files only if requested and no files were provided via arguments
    if not files_to_process and args.use_examples:
        print("No input files provided. Using example files as requested...")
        # Default example files (relative to CWD where script is run)
        example_files = [
            "data1.xlsx",
            "data2.xlsx",
            "possibly_corrupt.xlsx",
        ]
        # Check existence relative to CWD
        cwd = Path.cwd()
        for f_str in example_files:
            f_path = cwd / f_str
            if f_path.exists():
                files_to_process.append(f_str) # Keep relative path for processing list
                print(f"  Found example file: {f_str}")
            else:
                print(f"  Example file '{f_str}' not found in current directory ({cwd}), skipping.", file=sys.stderr)

    if not files_to_process:
         print("\nError: No input files to process.", file=sys.stderr)
         print(f"Usage: uv run {Path(__file__).relative_to(Path.cwd())} [--errors-dir <dir>] [--use-examples] [<file1.xlsx> ...]", file=sys.stderr)
         sys.exit(1)


    print(f"\nFiles to process: {files_to_process}")
    # Use resolve() to show the absolute path where errors will be saved
    print(f"Errors will be saved to: {(errors_directory / 'errors.csv').resolve()}")

    # Call the processing function with the list of original file path strings
    # the Path object for the errors directory, and the delta path
    print(f"Delta table will be saved to: {delta_path}")
    process_excel_files_safely(files_to_process, errors_directory, delta_path)

