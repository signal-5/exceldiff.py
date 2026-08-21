"""
exceldiff.py - Compare the first sheet of two or more Excel files.

WHY USE THIS SCRIPT
-------------------
Sometimes you have "the same" data spread across several Excel files and you
need to know:
  * Which rows exist in ALL files (the "same" rows)?
  * Which rows are unique to one file (the "different" rows)?

Instead of comparing entire rows (which is fragile, because column order or
extra columns differ), this script lets you compare only the KEY columns you
care about - for example an ID, an email, or a name. The comparison is
case-insensitive and ignores leading/trailing spaces.

HOW IT WORKS
------------
1. It reads the FIRST sheet of every Excel file you list.
2. For each file you tell it which column(s) form the "key" using -columnN,
   where N is the position of the file in the command (1 for the first file,
   2 for the second, and so on).
3. It builds a key from those columns (columnA is matched against columnC,
   columnB against columnD, etc., in the order you list them).
4. Rows whose key appears in EVERY file go to the "<file> same" sheets.
5. Rows whose key does NOT appear in every file go to the "<file> different"
   sheets.
6. Everything is written to a single output .xlsx file (-out).

RULES
-----
  * You must provide at least two Excel files.
  * You must provide one -columnN switch per file (e.g. -column1 -column2).
  * Each -columnN must list the SAME NUMBER of columns.
  * To match on several columns at once, separate them with commas:
        -column1 colA,colB
  * If a column NAME itself contains a comma, escape it with a backslash: \\,

EXAMPLE
-------
    python exceldiff.py -column1 columnA,columnB -column2 columnC,columnD ^
        -out diff.xlsx file1.xlsx file2.xlsx

This creates diff.xlsx with four sheets:
    "file1.xlsx same"      - full rows from file1 whose key is in both files
    "file2.xlsx same"      - full rows from file2 whose key is in both files
    "file1.xlsx different" - rows from file1 not found in file2
    "file2.xlsx different" - rows from file2 not found in file1
"""

import argparse
import os
import re
import sys

import pandas as pd


def split_columns(value):
    """Split on commas that are not escaped with backslash, then unescape."""
    parts = re.split(r'(?<!\\),', value)
    return [p.replace('\\,', ',').strip() for p in parts]


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description=(
            'Compare the first sheet of two or more Excel files and split the '
            'rows into "same" (present in all files) and "different" (unique '
            'to a file), based on the key columns you choose.'
        ),
        epilog=(
            'Example:\n'
            '  python exceldiff.py -column1 colA,colB -column2 colC,colD '
            '-out diff.xlsx file1.xlsx file2.xlsx\n\n'
            'Use -columnN (N = the file position, starting at 1) to pick the '
            'key column(s) for each file. Separate multiple columns with '
            'commas, and escape a comma inside a column name with \\,'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        '-out', required=True, help='Name of the output file (xlsx).'
    )

    # Parse known args to grab -out and files, and gather -columnN manually
    known, remaining = parser.parse_known_args(argv)

    files = []
    columns = {}
    i = 0
    while i < len(remaining):
        arg = remaining[i]
        m = re.match(r'^-column(\d+)$', arg)
        if m:
            idx = int(m.group(1))
            if i + 1 >= len(remaining):
                parser.error(f'Missing value for {arg}')
            columns[idx] = split_columns(remaining[i + 1])
            i += 2
        elif arg.startswith('-'):
            parser.error(f'Unknown switch: {arg}')
        else:
            files.append(arg)
            i += 1

    return known.out, files, columns


def validate(files, columns):
    if len(files) < 2:
        sys.exit('Error: Provide at least two Excel files.')

    for n in range(1, len(files) + 1):
        if n not in columns:
            sys.exit(f'Error: Missing switch -column{n} for file number {n}.')

    num_cols = len(columns[1])
    for n in range(1, len(files) + 1):
        if len(columns[n]) != num_cols:
            sys.exit(
                f'Error: -column{n} has {len(columns[n])} columns, '
                f'expected {num_cols}.'
            )

    for f in files:
        if not os.path.isfile(f):
            sys.exit(f'Error: File does not exist: {f}')


def build_key(df, cols):
    """Build a case-insensitive comparison key for the given columns."""
    for c in cols:
        if c not in df.columns:
            sys.exit(f'Error: The column "{c}" does not exist in the file.')
    key = df[cols].astype(str).apply(
        lambda row: tuple(v.strip().lower() for v in row), axis=1
    )
    return key


def main(argv):
    out, files, columns = parse_args(argv)
    validate(files, columns)

    # Read the first sheet of each file
    dfs = [pd.read_excel(f, sheet_name=0) for f in files]
    keys = [build_key(dfs[i], columns[i + 1]) for i in range(len(files))]

    # Determine the intersection of keys across all files
    key_sets = [set(k) for k in keys]
    common = set.intersection(*key_sets)

    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        # "same" sheets: rows whose key is in the common set
        for i, f in enumerate(files):
            name = os.path.basename(f)
            mask = keys[i].isin(common)
            sheet = f'{name} same'[:31]
            dfs[i][mask].to_excel(writer, sheet_name=sheet, index=False)

        # "different" sheets: rows whose key is not in the common set
        for i, f in enumerate(files):
            name = os.path.basename(f)
            mask = ~keys[i].isin(common)
            sheet = f'{name} different'[:31]
            dfs[i][mask].to_excel(writer, sheet_name=sheet, index=False)

    print(f'Done. The result was written to: {out}')


if __name__ == '__main__':
    main(sys.argv[1:])
