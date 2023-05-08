import pandas as pd


def import_sweep(file: pd.ExcelFile):
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

    # parse file using pandas
    # pandas has excellent docs
    # refer to the below for all possible arguments
    # https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
    data = file.parse(
        sheet_name="2002-LineSweep",
        usecols="G:H,U:V",  # only pull in necessary rows. Not best practice to hard code (oh well)
        header=35,  # define header row. Ignore rows above this
    )

    # data is of the form KP_beg,KP_end,2-way,4-way
    # where cells in the 2-way or 4-way columns having values are considered complete
    # cells that are blank are considered incomplete
    # there are some rows with the letter x in them that need to be cleaned
    # the columns will also be renamed
    print("original sweep")
    print(data)

    # rename columns
    columns = ["KP_beg", "KP_end", 2002.1, 2002.2]
    data.columns = columns

    # set the index of the dataframe to the KP columns
    # this removes these columns from the dataset so they aren't effected by global operations
    # drop all rows containing "x"
    # convert the dataset to true false using notnull based on whether a cell is blank
    # convert the type of the dataset to int so it is just 1 and 0
    data = data.set_index(["KP_beg", "KP_end"]).drop("x").notnull().astype("int")

    print("Macro sweep imported")
    return data


# this chunk of code is run only if this file is run as a script itself
# not when the code is imported into another script file
# very useful for testing functions without having to run your entire program
if __name__ == "__main__":
    fname = "raw/Master Progress Report_2023 05 05_MacroNK - Copy.xlsx"
    file = pd.ExcelFile(fname)

    result = import_sweep(file)
    result.to_csv("test.csv")
    print(result)
