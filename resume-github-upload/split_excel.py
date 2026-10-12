"""
Split one Excel (or CSV) file into many smaller Excel files.

Each output file keeps the header row and holds up to 100 data rows
(change with --rows).

Setup (once):
    pip install pandas openpyxl

Usage:
    python split_excel.py contacts.xlsx
    python split_excel.py contacts.xlsx --rows 100 --out my_folder   (optional: different output folder)

By default the split files are saved in the 'downloaded_hr_details' folder next to this script.
    python split_excel.py contacts.xlsx --sheet "Sheet2"
"""

import argparse
import math
import os
import sys

import pandas as pd

# Output files are saved in 'downloaded_hr_details' next to this script by default
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(SCRIPT_DIR, "downloaded_hr_details")


def main():
    parser = argparse.ArgumentParser(description="Split an Excel/CSV file into chunks.")
    parser.add_argument("input", help="Path to the .xlsx / .xls / .csv file")
    parser.add_argument("--rows", type=int, default=100, help="Data rows per file (default 100)")
    parser.add_argument("--out", default=DEFAULT_OUT, help="Output folder (default: downloaded_hr_details next to this script)")
    parser.add_argument("--sheet", default=0, help="Sheet name or index (default: first sheet)")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        sys.exit(f"File not found: {args.input}")
    if args.rows < 1:
        sys.exit("--rows must be at least 1")

    ext = os.path.splitext(args.input)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(args.input)
    else:
        sheet = args.sheet
        # allow a numeric sheet index passed on the command line
        if isinstance(sheet, str) and sheet.isdigit():
            sheet = int(sheet)
        df = pd.read_excel(args.input, sheet_name=sheet)

    total = len(df)
    if total == 0:
        sys.exit("The file has no data rows.")

    os.makedirs(args.out, exist_ok=True)
    base = os.path.splitext(os.path.basename(args.input))[0]
    parts = math.ceil(total / args.rows)
    width = max(3, len(str(parts)))

    for i in range(parts):
        chunk = df.iloc[i * args.rows : (i + 1) * args.rows]
        name = f"{base}_part_{str(i + 1).zfill(width)}.xlsx"
        chunk.to_excel(os.path.join(args.out, name), index=False)

    print(f"Done: {total} rows split into {parts} file(s) of up to {args.rows} rows in '{args.out}'.")


if __name__ == "__main__":
    main()
