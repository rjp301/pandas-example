# Pandas Excel Cleaning Example

Example module to parse over-complicated and poorly formatted progress tracking spreadsheet into a clean CSV. Meant to be used as a showcase of how Python can be very powerful when used in conjunction with Excel.

## Premise

Quantity progress tracking is captured in a large and convoluted workbook with several different formats and different crews on different pages. This workbook is not very intuitive for anyone but its creator to look at and is very difficult to use for data analysis. Additionally, the Excel file is quite heavy and consumes resources while open. We will convert this workbook into a homogenous, normalized, lightweight CSV file.

For the most part, the original data consists of tables plotting construction status against KP ranges on the Y-axis and crew types on the X-axis. We will follow that format but clean it up by using integer cost codes as the column names and keep the KP range as the index.

## File Management

As this progress report is distributed weekly, we will always want to use the latest version of the workbook. We can create a directory to drop the files in every week called `/raw`.

We will write a function called `latest_file` to find the most recent distribution in the folder based on common patterns in the filename. As you can see by looking in the `/raw` folder, the file naming scheme is not overly strict. It's alright though, because each filename has an internal pattern of "_...Report_YYYY MM DD_Macro..._" which can be used to extract the date.

## Progress File

The progress quantity tracking spreadsheet has several worksheets of interest showing status of KP chunks. They are in the format of KP chunk vs. crew type but have way too much extraneous information.

The worksheets of interest are as below and we will make a separate function for each.

- 2002-LineSweep
- 2003 - Clearing
- 2100-Pipeline DETAILS RoC

## Code

Have a look through the code in this order:

1. `latest_file.py`
1. `main.py`
1. `import_progress/__init__.py`
1. `import_progress/import_sweep.py`
1. `import_progress/import_clearing.py`
1. `import_progress/import_main.py`
1. `sterilize_KP.py`

## Output

The cleaned and combined CSV file is saved to `parsed.csv`.

## How to Run

1. Ensure Pandas is installed
1. Open and run `main.py`
