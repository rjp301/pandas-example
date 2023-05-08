# Pandas Examples

Example module to parse over-complicated and poorly formatted progress tracking
spreadsheet into a clean CSV. Meant to be used as a showcase of how Pandas can
be very powerfull when used in conjunction with Excel.

## Input Info

### /raw

- folder containing weekly progress report distribution
- each file has a some-what consistent naming scheme but not enough to be sorted by name
- internal pattern of "_...Report_YYYY MM DD_Macro..._" can be used to extract the date

### /raw/[file].xlsx

- progress quantity tracking spreadsheet
- has several worksheets of interest showing status of KP chunks
- vaguely in the below format but with a substantial amount of extraneous information

| KP Start | KP End   | Crew #1 | Crew #2 | ... |
| -------- | -------- | ------- | ------- | --- |
| 1050+000 | 1050+050 | 50m     | 0m      | ... |
| 1050+050 | 1050+100 | 50m     | 50m     | ... |
| 1050+100 | 1050+150 | 50m     | 30m     | ... |
| 1050+150 | 1050+200 | 0m      | 0m      | ... |
| ...      | ...      | ...     | ...     | ... |

- worksheets of interest

  - 2003 - Clearing
  - 2002-LineSweep
  - 2100-Pipeline DETAILS RoC

## Scripts

We will be making a seperate function for parsing each worksheet as they are
all in different formats and to keep our files a reasonable length

The desired output is below. The crew progress has been normalized, columns 
given sensible names, and all extraneous columns removed

| KP_beg  | KP_end  | crew_code_1 | crew_code_2 | ... |
| ------- | ------- | ----------- | ----------- | --- |
| 1050000 | 1050050 | 1           | 0           | ... |
| 1050050 | 1050100 | 1           | 1           | ... |
| 1050100 | 1050150 | 1           | 0.6         | ... |
| 1050150 | 1050200 | 0           | 0           | ... |
| ...     | ...     | ...         | ...         | ... |

`latest_file.py`
function to analyze all files in a directory with dates in their titles and return the latest

`import_clearing.py`
function to extract clearing information in desired format

`import_sweep.py`
function to extract sweep information in desired format

`import_main.py`
function to extract other crews information in desired format

`__init__.py`
function to execute the three import functions on the latest file and combine the results into
one dataframe and save to CSV

## How to Run

1. Ensure Pandas is installed
1. Open and run `main.py`
