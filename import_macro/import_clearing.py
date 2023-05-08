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
  data = file.parse(
    sheet_name="2003 - Clearing",
    usecols="G:H,J,T:X",
    header=31,
  )

  crews = [2003.1, 2003.2, 2003.3, 2003.4, 2003.5]
  columns = ["KP_beg", "KP_end", "scope", *crews]
  data = data.rename({i:j for i,j in zip(data.columns,columns)},axis=1)
  data = data[pd.to_numeric(data["KP_beg"],errors="coerce").notnull()] # remove non-numeric KPs

  result = data[["KP_beg","KP_end"]].copy()

  def get_status(r: pd.Series, scope: str, complete: str):
    if r[scope] == 0: return None
    return 1 if r[complete] > 0 else 0

  for crew in crews:
    temp = data[["scope",crew]].copy()
    temp[crew] = pd.to_datetime(temp[crew],errors="coerce").notnull().astype(int)
    result[crew] = temp.apply(lambda r: get_status(r,"scope",crew), axis=1)

  print("Macro clearing imported")
  return result.set_index(["KP_beg","KP_end"])

if __name__ == "__main__":
  file = pd.ExcelFile("data/macro/raw/Macro Master Progress Report_2023 04_14_MacroNK.xlsx")
  result = import_clearing(file)
  result.to_csv("test.csv")
  print(result)