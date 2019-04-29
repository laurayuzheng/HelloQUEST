import pandas as pd 

file_name = "2019-03 Humanity"
file_out = "humanity_eventid"

def parse_eventid():
	data = pd.read_excel("input/" + file_name + ".xlsx", sheet_name=0)
	data.dropna(subset=['shift_title'], inplace = True)

	new = data["shift_title"].str.split("~", n = 1, expand = True)
	data.drop(columns =["shift_title"], inplace = True)
	#print(new)
	data["shift_title"] = new[0] 
	data["event_id"] = new[1]
	data = data[['eid','location','start_day','start_time','end_time','total_time','shift_title','event_id']]

	data.to_excel("input/" + file_out + ".xlsx", index=None)

if __name__ == "__main__":
	parse_eventid()
