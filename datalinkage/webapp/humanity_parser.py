import pandas as pd 
file_name = "humanity_in"

data = pd.read_csv(file_name + ".csv")
data.dropna(inplace = True)

new = data["shift_title"].str.split("~", n = 1, expand = True)
data.drop(columns =["shift_title"], inplace = True)
data["shift_title"] = new[0]
data["event_id"] = new[1]

data.to_csv("humanity_out.csv")