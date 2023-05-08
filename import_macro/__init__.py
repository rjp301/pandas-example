# relative import our functions from same folder as current script
from .import_main import import_main
from .import_clearing import import_clearing
from .import_sweep import import_sweep

from latest_file import latest_file  # our script for getting latest file by filename
from sterilize_KP import sterilize_KP

import os
import pandas as pd  # import pandas library with alias of pd for breviety


def import_macro():
    fname, date = latest_file(
        path="raw",
        date_format="%Y %m %d",
        prefix="Report_",
        suffix="_Macro",
    )

    # using the pattern of reading file once and passing that file to each function
    # this improves performance because reading excel files takes a couple seconds
    # only have to do it once this way and each file takes care of parsing
    file = pd.ExcelFile(fname)

    # create empty array and fill with dataframes from each function
    # nice pattern because if more data is needed from progress sheet it's trival to add
    dfs = []
    dfs.append(import_main(file))
    dfs.append(import_clearing(file))
    dfs.append(import_sweep(file))

    # concatentate dataframes
    data = pd.concat(dfs).sort_index()
    data = sterilize_KP(data)
    print(data)
    
    # save result to CSV file
    fname = "parsed.csv"
    data.to_csv(fname)
    

    print(f"Macro data extracted for {date}\n")
