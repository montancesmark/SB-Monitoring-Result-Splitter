# SB-Monitoring-Result-Splitter
# Created by: Mark Montances
# Date: Dec 6, 2024


# Excel Macro: Split Columns to Files

This tool splits specified columns in an Excel sheet into separate files based on predefined logic.
It assumes a specific spreadsheet formating to work.


## Features
- Detects column ranges using specific criteria (e.g., "FT" or "PT" in row 3, "SCORE" in row 2).
- Creates new files with dynamically generated filenames.
- Includes value from column `E1` (faculty fullname) in the filename.

## Requirements
- Microsoft Excel 2010 or later.
- Macros must be enabled for the tool to work.

## How to Use
1. Download the `SplitColumnsTool.xlsm` file from this repository.
2. Open the file in Excel.
3. Enable macros when prompted.
4. Run the macro by:
   - Pressing `Alt + F8`, selecting `SplitColumnsToFiles`, and clicking `Run`.
   - Or clicking the assigned button (if applicable).
   - Choose the file you wish to automatically split
   - The output files will be generated in the same directory/folder as the original/input file



## License
This project is licensed under the MIT License. Feel free to use, modify, and share.
