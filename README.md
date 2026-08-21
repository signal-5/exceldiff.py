# exceldiff.py
Compare the first sheet of two or more Excel files.

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
  * If a column NAME itself contains a comma, escape it with a backslash: \,

EXAMPLE
-------
    python exceldiff.py -column1 columnA,columnB -column2 columnC,columnD ^
        -out diff.xlsx file1.xlsx file2.xlsx

This creates diff.xlsx with four sheets:
    "file1.xlsx same"      - full rows from file1 whose key is in both files
    "file2.xlsx same"      - full rows from file2 whose key is in both files
    "file1.xlsx different" - rows from file1 not found in file2
    "file2.xlsx different" - rows from file2 not found in file1
