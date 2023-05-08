# New way to think of scope
# Some areas do not have any scope and therefore should not be counted in progress
# NaN for not in scope, 0 for incomplete, 1 for complete

import pandas as pd
import re


def import_main(file: pd.ExcelFile):
    """
    Import all progress chunks from Macro progress tracking workbook
    :param file: Excel File containing progress information
    :return: Pandas DataFrame with index of [KP_beg,KP_end] and column name of crew cost codes

    Both cost codes and KPs should be as floats

    Values:
      NaN: No progress able to be claimed for chunk
      0: Progress has not been claimed
      1: Progress has been claimed
    """

    # this particular sheet is much more complicated than the others
    # there are many crews in this sheet
    # each crew has an arbitraty number of columns associated with it
    # there is a bunch of useless rows between the data and the header row

    # parse file using pandas
    # pandas has excellent docs
    # refer to the below for all possible arguments
    # https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
    data = file.parse(
        sheet_name="2100-Pipeline DETAILS RoC",
        usecols="N:O,Q:BY",
    )

    # print("original main")
    # print(data)  # move this print statement up and down to see what each stage does

    # create dictionary to use in renaming dataframe method
    col_rename = {
        data.columns[0]: "KP_beg",
        data.columns[1]: "KP_end",
    }

    data = (
        data.rename(col_rename, axis=1)  # rename KP_beg and KP_end columns
        .drop(range(30))  # drop the first 30 rows
        .dropna(subset=["KP_beg", "KP_end"])  # remove rows with not KP_beg or KP-end
    )

    # remove non-numeric KPs
    data = data[pd.to_numeric(data["KP_beg"], errors="coerce").notnull()]

    # remove columns with unnamed
    data = data.drop([i for i in data.columns if "Unnamed" in i], axis=1)

    # if scope is 0 null else percentage complete
    def get_status(r: pd.Series, scope: str, complete: str):
        if r[scope] == 0:
            return None
        return r[complete] / r[scope]

    result = data[["KP_beg", "KP_end"]].copy()
    for index, col in enumerate(data.columns):
        if "% Comp" not in col:
            continue

        prev_col = data.columns[index - 1]
        crew_num = re.search(r"[0-9]{4}", prev_col)  # use regex to find crew cost code
        if not crew_num:
            continue

        crew_num = float(crew_num.group(0))
        temp = data[[prev_col, col]].copy()
        result[crew_num] = temp.apply(lambda r: get_status(r, prev_col, col), axis=1)

    print("Macro main imported")
    return result.set_index(["KP_beg", "KP_end"])


# this chunk of code is run only if this file is run as a script itself
# not when the code is imported into another script file
# very useful for testing functions without having to run your entire program
if __name__ == "__main__":
    fname = "raw/Master Progress Report_2023 05 05_MacroNK - Copy.xlsx"
    file = pd.ExcelFile(fname)

    result = import_main(file)
    result.to_csv("test.csv")
    print(result)
