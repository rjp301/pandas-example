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
  data = file.parse(
    sheet_name="2100-Pipeline DETAILS RoC",
    usecols="N:O,Q:BY",
  )

  col_rename = {
    data.columns[0]:"KP_beg",
    data.columns[1]:"KP_end",
  }

  data = (data
    .rename(col_rename,axis=1)
    .drop(range(30))
    .reset_index(drop=True)
    .dropna(subset=["KP_beg","KP_end"])  
  )
  
  data = data[pd.to_numeric(data["KP_beg"],errors="coerce").notnull()] # remove non-numeric KPs
  data = data.drop([i for i in data.columns if "Unnamed" in i],axis=1) # remove unnamed columns

  def get_status(r: pd.Series, scope: str, complete: str):
    if r[scope] == 0: return None
    return r[complete] / r[scope]

  result = data[["KP_beg","KP_end"]].copy()
  for index,col in enumerate(data.columns):
    if "% Comp" in col:
      prev_col = data.columns[index-1]
      crew_num = re.search(r"[0-9]{4}", prev_col)
      if not crew_num: continue

      crew_num = float(crew_num.group(0))
      temp = data[[prev_col,col]].copy()
      result[crew_num] = temp.apply(lambda r: get_status(r,prev_col,col), axis=1)

  print("Macro main imported")
  return result.set_index(["KP_beg","KP_end"])

if __name__ == "__main__":
  file = pd.ExcelFile("data/macro/Master Progress Report_2023 03 24_Macro.xlsx")
  result = import_main(file)
  result.to_csv("test.csv")
  print(result)