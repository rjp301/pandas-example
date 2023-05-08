
# 
# relative import o
from .import_main import import_main
from .import_clearing import import_clearing
from .import_sweep import import_sweep

from pathlib import Path # Path library to make system-independent filepaths
from latest_file import latest_file # our script for getting latest file by filename

import pandas as pd # import pandas with alias of pd for breviety


def import_macro():
  fname,date = latest_file(
    path="raw", # folder containing files
    date_format="%Y %m %d",
    prefix="Report_",
    suffix="_Macro",
  )

  file = pd.ExcelFile(fname)
  dfs = []
  dfs.append(import_main(file))
  dfs.append(import_clearing(file))
  dfs.append(import_sweep(file))

  fname = os.path.join(path,"..","crews.csv")
  pd.concat(dfs).sort_index().to_csv(fname)

  print(f"Macro data extracted for {date}\n")