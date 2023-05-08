import pandas as pd
import re
from mapping.data import convert_chainage_string

def import_crossings(file: pd.ExcelFile):
  """
  Import all completed crossings from progress workbook
  :param file: Excel File containing progress information
  :return: Pandas DataFrame with columns of [WC_id,complete]
  """
  data = file.parse(
    sheet_name="2100-Pipeline DETAILS RoC",
    header=30,
    usecols="CA,CI"
  )

  def get_IDs(s: str):
    try:
      strings = re.findall(r"[TWPJ][0-9]+\.[0-9]",s)
      return strings
    except: return None

  col_rename = {
    data.columns[0]: "string",
    data.columns[1]: "date"
  }

  data = (data
    .rename(col_rename,axis=1)
    .dropna(how="all")
  )
  data["complete"] = ~data.date.isnull()
  dfs = []

  # Add ID to junk yard as it doesn't have one
  junk_yard_id = "J0000.0"
  data["string"] = data["string"].apply(lambda i: f"{i} {junk_yard_id}" if "Junk" in str(i) else i )

  for _,row in data.iterrows():
    temp = pd.DataFrame()

    temp["WC_id"] = get_IDs(row["string"])

    for col in data.columns:
      temp[col] = row[col]
    dfs.append(temp)

  result = pd.concat(dfs,ignore_index=True)

  print("Macro crossings imported")
  return result

if __name__ == "__main__":
  file = pd.ExcelFile("data/macro/raw/Macro Master Progress Report_2023 04_14_MacroNK.xlsx")
  result = import_crossings(file)
  result.to_csv("macro_crossings.csv",index=False)
  print(result)