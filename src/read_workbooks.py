    return success, dfs


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
    # First try with calamine engine
    try:
        result = subprocess.run(
            ['python', 'src/read_excel.py', filepath, 'calamine'],
            capture_output=True,
            timeout=300
        )
        
        if result.returncode == 0:
            print(f"Successfully processed {filepath} with calamine engine")
            # Now load the data and write to Delta
            dfs = read_excel(filepath, engine="calamine")
            success = write_dfs_to_delta(dfs, delta_path, partition_col)
            return success
    except subprocess.TimeoutExpired:
        print(f"Timeout processing: {filepath} with calamine engine")
    except Exception as e:
        print(f"Exception with calamine engine: {e}")
    
    # If first attempt failed - try with default engine
    print(f"Retrying {filepath} with default engine...")
    try:
        result = subprocess.run(
            ['python', 'src/read_excel.py', filepath],
            capture_output=True,
            timeout=300
        )
        
        if result.returncode == 0:
            print(f"Successfully processed {filepath} with default engine")
            # Process with default engine
            dfs = read_excel(filepath)
            success = write_dfs_to_delta(dfs, delta_path, partition_col)
            return success
        else:
            print(f"Error in file with default engine: {filepath}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"Timeout processing: {filepath} with default engine")
        return False
    except Exception as e:
        print(f"Exception with default engine: {e}")
        return False


if __name__ == "__main__":
    if args.write_delta:
        # Ensure directory exists
        os.makedirs(args.delta_path, exist_ok=True)
        
        if args.engine:
            # If engine is specified explicitly, use it directly
            dfs = read_excel(args.file, args.engine)
            success = write_dfs_to_delta(dfs, args.delta_path, args.partition_col)
            if success:
                print("Successfully wrote all Excel sheets to Delta tables")
        else:
            # Use the safe subprocess approach with fallback
            success = safe_process_workbook(args.file, args.delta_path, args.partition_col)
            if success:
                print("Successfully processed workbook")
            else:
                print("Failed to process workbook")
    else:
        # Just read the Excel file
        if args.engine:
            dfs = read_excel(args.file, args.engine)
        else:
            # Try with calamine first, then fall back to default if needed
            try:
                print("Attempting to read with calamine engine...")
                dfs = read_excel(args.file, engine="calamine")
                print("Successfully read with calamine engine")
            except Exception as e:
                print(f"Failed with calamine engine: {e}")
                print("Falling back to default engine...")
                dfs = read_excel(args.file)
        
