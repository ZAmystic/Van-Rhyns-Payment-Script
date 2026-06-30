#Libaries Imports
import pandas as pd
import os
import glob
import datetime

# Directories for imports and processed files
IMPORTS_DIR = "imports/"
PROCESSED_DIR = "processed/"

# Pick up for all excel files in the imports directory
files = glob.glob(os.path.join(IMPORTS_DIR, "**", "*.xls*"), recursive=True)

column_map = {
    "matter nr": "matter no",
    "matter no": "matter no",
    "matter number": "matter no",
    "matter to": "matter no"
}

unwanted_patterns = ["999999", "FULL BANK STATEMENT"]

# Define the required columns (normalized to lowercase)
required_cols = ["date", "description", "matter no", "amount", "balance"]

def fix_date_format(x):
    # Case 0: Already a datetime object
    if isinstance(x, (pd.Timestamp, datetime.date)):
        return pd.to_datetime(x)

    x = str(x).strip()

    # Case 1: Already has dividers (YYYY/MM/DD)
    if "/" in x:
        try:
            return pd.to_datetime(x)   # let pandas infer
        except:
            return pd.NaT

    # Case 2: Compact format (YYYYMMDD)
    elif len(x) == 8 and x.isdigit():
        try:
            formatted = f"{x[0:4]}/{x[4:6]}/{x[6:8]}"
            return pd.to_datetime(formatted)
        except:
            return pd.NaT

    else:
        return pd.NaT

for file in files:
    try:
        filename = os.path.basename(file)

        # 🔎 Delete unwanted files first
        if any(pat in filename.upper() for pat in unwanted_patterns):
            os.remove(file)
            print(f"Deleted unwanted file: {file}")
            continue  # skip to next file

        df = pd.read_excel(file)

        # Some columns have been giving issues due to capitalization and whitespace. 
        # Therefore, we will convert all column names to lowercase and strip whitespace from the column names.
        df.columns = df.columns.str.lower().str.strip()

        # Rename columns based on the column_map dictionary
        # Some columns have been spelled using different variations, therefore we will rename them to a standard name.
        df.rename(columns=column_map, inplace=True)

        # Keep only those columns
        df = df[required_cols]

        # Drops all rows with negative values in the 'Amount' column
        # This is due to the fact that negative values in the 'Amount' column are not valid amounts for allocations and unallocations. 
        # Therefore, we will drop all rows with negative values in the 'Amount' column.
        df = df[df['amount'] >= 0]

        # Allocated column is created to indicate whether the matter is allocated or not.
        # This coloumn will server for SQL query to filter out unallocated matters or to find them.
        df["allocated"] = df["matter no"].apply(lambda x: "1" if x != 999999 else "0")

        # Apply the fix to the date column
        df["date"] = df["date"].apply(fix_date_format)

        # Format consistently for Excel readability
        df["date"] = df["date"].dt.strftime("%Y-%m-%d")

        # Save cleaned file to processed folder
        rel_path = os.path.relpath(file, IMPORTS_DIR)
        processed_path = os.path.join(PROCESSED_DIR, rel_path)
        os.makedirs(os.path.dirname(processed_path), exist_ok=True)
        df.to_excel(processed_path, index=False)
        os.remove(file)

        # Try removing empty parent folder
        folder = os.path.dirname(file)
        try:
            os.removedirs(folder)
            print(f"Deleted empty folder: {folder}")
        except OSError:
            pass

        print(f"\033[92m{file} successfully processed to file: {processed_path}\033[0m")
    except Exception as e:
        print(f"\033[91mError processing file {file}: {e}\033[0m")

