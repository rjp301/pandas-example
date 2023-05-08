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
  data = file.parse(
    sheet_name="2002-LineSweep",
    usecols="G:H,U:V",
    header=35,
  )

  columns = ["KP_beg","KP_end",2002.1,2002.2]
  data = (data
    .rename({i:j for i,j in zip(data.columns,columns)},axis=1)
    .set_index(["KP_beg","KP_end"])
    .drop("x")
    .notnull()
    .astype('int')
  )

  print("Macro sweep imported")
  return data


if __name__ == "__main__":
  file = pd.ExcelFile("data/macro/Master Progress Report_2023 03 24_Macro.xlsx")
  result = import_sweep(file)
  result.to_csv("test.csv")
  print(result)