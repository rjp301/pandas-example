import pandas as pd


def import_clearing(file: pd.ExcelFile):
    """
    Import all progress chunks from Macro progress tracking workbook
    :param file: Excel File containing progress information
    :return: Pandas DataFrame with index of [KP_beg,KP_end] and column name of crew cost codes.

    Both cost codes and KPs should be as floats

    Values:
      NaN: No progress able to be claimed for chunk
      0: Progress has not been claimed
      1: Progress has been claimed
    """

    # parse file using pandas
    # pandas has excellent docs
    # refer to the below for all possible arguments
    # https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
    data = file.parse(
        sheet_name="2003 - Clearing",
        usecols="G:H,J,T:X",  # only pull in necessary rows. Not best practice to hard code (oh well)
        header=31,  # define header row. Ignore rows above this
    )

    # data is of the form KP_beg,KP_end,scope,crews...
    # where crews with a date assigned have completed their scope in that KP range
    # if scope is 0 there is no work to be done in that area
    # print("original clearing")
    # print(data)

    crews = [2003.1, 2003.2, 2003.3, 2003.4, 2003.5]  # custom crew names

    # define column names using unpacking operator and rename columns
    # https://towardsdatascience.com/unpacking-operators-in-python-306ae44cd480
    columns = ["KP_beg", "KP_end", "scope", *crews]
    data.columns = columns

    # remove non-numeric KPs
    # to_numeric returns series with numbers and null values
    # notnull returns series of true false
    # pass true false series to data to filter for only rows corresponding to true
    data = data[pd.to_numeric(data["KP_beg"], errors="coerce").notnull()]

    # make a copy of the data dataframe with only KP_beg and KP-end columns
    result = data[["KP_beg", "KP_end"]].copy()

    # define function used to map over the dataframe to define area status per crew
    def get_status(r: pd.Series, scope: str, crew: str) -> float | None:
        if r[scope] == 0:
            return None
        return 1 if r[crew] > 0 else 0

    for crew in crews:
        # make copy of only scope and crew column of data dataframe
        temp = data[["scope", crew]].copy()

        # use same pattern of coercing series to a certain type
        # this time convert true false column into 1 or 0
        # crew column is 1 if date else 0
        temp[crew] = pd.to_datetime(temp[crew], errors="coerce").notnull().astype(int)

        # now apply status function to determine in conjunction with scope the status
        result[crew] = temp.apply(lambda r: get_status(r, "scope", crew), axis=1)

    print("Macro clearing imported")
    # set index of dataframe to be KP range before returning
    return result.set_index(["KP_beg", "KP_end"])


# this chunk of code is run only if this file is run as a script itself
# not when the code is imported into another script file
# very useful for testing functions without having to run your entire program
if __name__ == "__main__":
    fname = "raw/Master Progress Report_2023 05 05_MacroNK - Copy.xlsx"
    file = pd.ExcelFile(fname)

    result = import_clearing(file)
    result.to_csv("test.csv")
    print(result)
