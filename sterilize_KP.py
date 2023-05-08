import pandas as pd


def sterilize_KP(data: pd.DataFrame) -> pd.DataFrame:
    data = data.reset_index()
    length_beg = len(data)

    assert "KP_beg" in data.columns, "KP_beg column must be included"
    assert "KP_end" in data.columns, "KP_end column must be included"

    # get all chainage values, remove duplicates, remove blanks and sort
    chainages = (
        pd.Series(data["KP_beg"].tolist() + data["KP_end"].tolist())
        .sort_values()
        .drop_duplicates()
        .dropna()
        .tolist()
    )

    result = pd.DataFrame()  # instantiate empty dataframe
    result["KP_beg"] = chainages[:-1]  # KP_beg = all chainages except last
    result["KP_end"] = chainages[1:]  # KP_end = all chainages except first

    temps = []  # instantiate empty array to fill with series
    for _, row in result.iterrows():
        # filter dataset for all rows that overlap with current chainage range
        temp = data.loc[(data["KP_beg"] <= row.KP_beg) & (data["KP_end"] >= row.KP_end)]

        # remove KP_beg and KP_end columns
        temp = temp.drop(["KP_beg", "KP_end"], axis=1)

        # append current row KP_beg and KP_end to the max of each column to get a row
        temp = pd.concat([row, temp.max()])

        # append that row to the list of rows
        temps.append(temp)

    # combine all rows and set index
    result = pd.concat(temps, axis=1).T.set_index(["KP_beg", "KP_end"])
    length_end = len(result)

    print(f"Dataframe sterilized from {length_beg} to {length_end} rows")
    return result
