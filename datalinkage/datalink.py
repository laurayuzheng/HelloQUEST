import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from datetime import datetime, timedelta
from pytz import timezone
import tkinter as tk
from tkinter import filedialog

# current month
#month = '2019-03'

# average wage
#avg_wage = 13.72

# file names
bob_name = ''
humanity_name = ''
sel_name = ''

# output files
#bob_out = 'bob_out.xlsx'
#humanity_out = 'humanity_out.xlsx'
#sel_out = 'sel_out.xlsx'
#joined_out = month + '_' + 'joined.xlsx'


standard_tz = timezone('US/Eastern')

timezones = {
	'Atlanta, GA' : timezone('US/Eastern'),
	'Austin, TX' : timezone('US/Central'),
	'Boston, MA' : timezone('US/Eastern'),
	'Chicago, IL' : timezone('US/Central'),
	'Cleveland, OH' : timezone('US/Eastern'),
	'Columbus, OH' : timezone('US/Eastern'),
	'Dallas, TX' : timezone('US/Central'),
	'Denver, CO' : timezone('US/Mountain'),
	'Houston, TX' : timezone('US/Central'),
	'Indianapolis, IN' : timezone('US/Eastern'),
	'Jacksonville, FL' : timezone('US/Eastern'),
	'Nashville, TN' : timezone('US/Eastern'),
	'New York, NY' : timezone('US/Eastern'),
	'Philadelphia, PA' : timezone('US/Eastern'),
	'Phoenix, AZ' : timezone('US/Arizona'),
	'Pittsburgh, PA' : timezone('US/Eastern'),
	'Pleasanton, CA' : timezone('US/Pacific'),
	'Portland, OR' : timezone('US/Pacific'),
	'Santa Ana, CA' : timezone('US/Pacific'),
	'Seattle, WA' : timezone('US/Pacific'),
	'Tampa, FL' : timezone('US/Eastern')}

def write_slogan():
    print(bob_name + '\n' + humanity_name + '\n' + sel_name)

def open_bob():
	global bob_name
	bob_name = filedialog.askopenfilename(title = "Select BOB file",
		filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))

def open_humanity():
	global humanity_name
	humanity_name = filedialog.askopenfilename(title = "Select Humanity file",
		filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))

def open_SEL():
	global sel_name
	sel_name = filedialog.askopenfilename(title = "Select SEL file",
		filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))

def set_num_sales(row):
    if row['customer_id'] == 0:
        val = 0
    else:
        val = 1
    return val

def twotwo(row):
	if row['box_type'] == '2x2':
		return 1
	else:
		return 0

def threetwo(row):
	if row['box_type'] == '3x2':
		return 1
	else:
		return 0

def twofour(row):
	if row['box_type'] == '2x4':
		return 1
	else:
		return 0

def fourtwo(row):
	if row['box_type'] == '4x2':
		return 1
	else:
		return 0

def threefour(row):
	if row['box_type'] == '3x4':
		return 1
	else:
		return 0

def parse_eventid():
	data = pd.read_excel(humanity_name, sheet_name=0)
	data.dropna(subset=['shift_title'], inplace = True)

	new = data["shift_title"].str.split("~", n = 1, expand = True)
	data.drop(columns =["shift_title"], inplace = True)
	#print(new)
	data["shift_title"] = new[0] 
	data["event_id"] = new[1]
	data = data[['eid','location','start_day','start_time','end_time','total_time','shift_title','event_id']]

	return data
	#data.to_excel("input/" + file_out + ".xlsx", index=None)

def combine(): 

	free_events_out = month.get() + '_' + 'free_events.xlsx'
	nosale_out = month.get() + '_' + 'no_sale.xlsx'
	paid_events_out = month.get() + '_' + 'paid_events.xlsx'
	per_shift_out = month.get() + '_' + 'per_shift.xlsx'
	per_event_out = month.get() + '_' + 'per_event.xlsx'
	combined_out = month.get() + '_' + 'combined.xlsx'
	# load data files
	bob = pd.read_excel(bob_name, sheet_name=0)
	#humanity = pd.read_excel('input/' + humanity_name)
	sel = pd.read_excel(sel_name, sheet_name=0)
	humanity = parse_eventid()
	# no need to filter hellofresh from bob anymore

	# formatting dates in bob

	sel.rename(columns={'Event ID':'event_id', 'Category of Event':'category',
		'Market':'market','Amortized Cost':'amortized_cost'}, inplace=True)

	bob['date_sign_up'] = pd.to_datetime(bob['date_sign_up'], format='%m/%d/%y %-H:%M')
	bob['date_sign_up'] = bob['date_sign_up'].dt.strftime('%m/%d/%y')
	bob.rename(columns={'Employee ID' : 'eid', 'Office':'office',
	                          'Box Type':'box_type',
	                          'date_sign_up':'date'}, 
	                 inplace=True)
	# drop unnecessary columns
	# bob = bob.drop(['initial_sales_order_nr','customer_name'], axis=1)
	 #, sheet_name='Sheet1')

	# cleaning humanity sheet
	humanity['start_day'] = pd.to_datetime(humanity['start_day'], format='%m/%d/%y')
	humanity['start_day'] = humanity['start_day'].dt.strftime('%m/%d/%y')
	humanity.rename(columns={'location':'office','start_day':'date'}, 
	                 inplace=True)

	#print(humanity['start_time'])
	#print(humanity['start_time'])
	humanity['start_time'] = pd.to_datetime(humanity['start_time'],format='%H:%M:%S')
	#humanity['start_time'] = humanity['start_time'].dt.strftime('%H:%M')
	humanity['end_time'] = pd.to_datetime(humanity['end_time'],format='%H:%M:%S')
	#humanity['end_time'] = humanity['end_time'].dt.strftime('%H:%M')
	humanity = humanity.dropna(subset=['eid', 'office','shift_title'])

	# localize the time zones, then set time zones to eastern
	# then, convert the new time zone to a string
	for i, row in humanity.iterrows():
		tz = timezones[row['office']]
		#print("before: " + humanity.at[i,'start_time'].strftime('%H:%M'))
		
		humanity.at[i,'start_time'] = tz.localize(row['start_time']) \
		    .astimezone(standard_tz).strftime('%H:%M')

		#print("after: " + humanity.at[i,'start_time'].strftime('%H:%M'))
		humanity.at[i,'end_time'] = tz.localize(row['end_time']) \
		    .astimezone(standard_tz).strftime('%H:%M')

	humanity['start_time'] = humanity['start_time'].dt.strftime('%H:%M')
	humanity['end_time'] = humanity['end_time'].dt.strftime('%H:%M')

	# merging humanity and bob
	# replacing NaN values with 0
	joined = pd.merge(bob, humanity, on=['eid', 'office', 'date'], how='outer')
	#joined = joined.drop(columns=['location', 'start_day', 'Employee ID'])
	#joined.rename(columns={'Office':'office',
	#                          'Box Type':'box_type',
	#                          'date_sign_up':'date'}, 
	#                 inplace=True)

	joined = joined.fillna(0)
	joined = joined[joined['eid'] > 0]
	joined = joined[joined['office'] != 0]
	joined = joined[joined['date'] != 0]
	joined = joined[joined['shift_title'] != 0]
	joined['num_sales'] = joined.apply(set_num_sales,axis=1)
	joined['2x2'] = joined.apply(twotwo,axis=1)
	joined['3x2'] = joined.apply(threetwo,axis=1)
	joined['2x4'] = joined.apply(twofour,axis=1)
	joined['4x2'] = joined.apply(fourtwo,axis=1)
	joined['3x4'] = joined.apply(threefour,axis=1)

	# standardize the column order
	joined = joined[['eid','customer_id','event_id', 'num_sales',
					'office','box_type','date','start_time','end_time',
					'total_time','shift_title','2x2','3x2','2x4','4x2','3x4']]

	nosales = joined[joined['customer_id'] == 0]
	free_events = joined[joined['event_id'] == 0]
	paid_events = joined[joined['event_id'] != 0]

	# start here
	sel = sel[['event_id','category','market','amortized_cost']]
	per_shift = joined.groupby(['eid','shift_title','office','start_time','end_time',
		'total_time','event_id'], as_index=False).agg(
		{'num_sales' : 'sum', '2x2' : 'sum', '3x2' : 'sum', '2x4' : 'sum', '4x2' : 'sum',
		'3x4' : 'sum'})
	# [['num_sales','total_time','box_type']].sum()
	#print(per_shift)
	avg_wage_num = float(avg_wage.get())
	per_shift['wage_cost'] = (avg_wage_num * per_shift['total_time']).round(decimals=2)

	per_event = per_shift.groupby(['shift_title','office','event_id'], as_index=False).agg(
		{'num_sales' : 'sum', 'total_time':'sum','wage_cost':'sum',
		'2x2' : 'sum', '3x2' : 'sum', '2x4' : 'sum', '4x2' : 'sum','3x4' : 'sum'})

	combined = pd.merge(per_event, sel, on='event_id', how='outer')
	combined = combined.fillna(0)
	combined['total_cost'] = combined['wage_cost'] + combined['amortized_cost']

	# standardize column order
	combined = combined[['event_id','shift_title','office','category','num_sales',
	'total_time','2x2','3x2','2x4','4x2','3x4','market','wage_cost',
	'amortized_cost','total_cost']]

	#print(per_shift_boolean.get())
	#print(per_event_boolean.get())
	#print(nosale_boolean.get())
	#print(free_events_boolean.get())
	#print(paid_events_boolean.get())
	#print(combined_boolean.get())

	# output all to csv files
	if per_shift_boolean.get():
		per_shift.to_excel('output/' + per_shift_out, index=None)
	
	if per_event_boolean.get():
		per_event.to_excel('output/' + per_event_out, index=None)
	#bob.to_excel('output/'+ bob_out)
	#humanity.to_excel('output/' + humanity_out)
	#joined.to_excel('output/' + joined_out, index=None)
	if nosale_boolean.get(): 
		nosales.to_excel('output/' + nosale_out, index=None)
	
	if free_events_boolean.get():
		free_events.to_excel('output/' + free_events_out, index=None)
	
	if paid_events_boolean.get():
		paid_events.to_excel('output/' + paid_events_out, index=None)
	
	if combined_boolean.get():
		combined.to_excel('output/' + combined_out, index=None)

if __name__ == "__main__":

#	combine(humanity)
	root = tk.Tk()
	root.title('hellofresh data combinatron 3000')
	frame = tk.Frame(root)
	#root.geometry("500x200")
	root.configure(background='#F2F5DC')
	frame.grid()
	frame.configure(background='#F2F5DC')

	combined_boolean = tk.IntVar()
	nosale_boolean = tk.IntVar()
	free_events_boolean = tk.IntVar()
	paid_events_boolean = tk.IntVar()
	per_shift_boolean = tk.IntVar()
	per_event_boolean = tk.IntVar()
	avg_wage = tk.StringVar(frame, value='13.72');
	month = tk.StringVar()

	combined_cb = tk.Checkbutton(frame,
						text="combined",
						variable=combined_boolean,
						#onvalue=True, offvalue=False
				)
	#combined_cb.pack(side=tk.TOP)
	combined_cb.grid(row=1, column=0)

	nosale_cb = tk.Checkbutton(frame,
						text="no sale events",
						variable=nosale_boolean,
						#onvalue=True, offvalue=False
				)
	nosale_cb.grid(row=2, column=0, padx=10, pady=10)

	free_cb = tk.Checkbutton(frame,
						text="free events",
						variable=free_events_boolean,
						#onvalue=True, offvalue=False
				)
	free_cb.grid(row=3, column=0, padx=10, pady=10)

	paid_cb = tk.Checkbutton(frame,
						text="paid events",
						variable=paid_events_boolean,
						#onvalue=True, offvalue=False
				)
	paid_cb.grid(row=4, column=0, padx=10, pady=10)

	shift_cb = tk.Checkbutton(frame,
						text="per shift",
						variable=per_shift_boolean,
						#onvalue=True, offvalue=False
				)
	shift_cb.grid(row=5, column=0, padx=10, pady=10)

	event_cb = tk.Checkbutton(frame,
						text="per event",
						variable=per_event_boolean,
						#onvalue=True, offvalue=False
				)
	event_cb.grid(row=6, column=0, padx=10, pady=10)

	L1 = tk.Label(frame, text="average wage rate")
	L1.grid(row=1, column=1)
	wages_entry = tk.Entry(frame,
						textvariable=avg_wage)
	wages_entry.grid(row=2, column=1, padx=10, pady=10)

	L2 = tk.Label(frame, text="month [YYYY-MM]")
	L2.grid(row=3, column=1)
	month_entry = tk.Entry(frame,
						textvariable=month)
	month_entry.grid(row=4, column=1, padx=10, pady=10)

	bob_button = tk.Button(frame, 
	                   text="Upload BOB File", 
	                   command=open_bob)

	bob_button.grid(row=1, column=2, padx=10, pady=10)

	humanity_button = tk.Button(frame, 
	                   text="Upload Humanity File", 
	                   command=open_humanity)
	humanity_button.grid(row=2, column=2, padx=10, pady=10)

	sel_button = tk.Button(frame, 
	                   text="Upload SEL File", 
	                   command=open_SEL)
	sel_button.grid(row=3, column=2, padx=10, pady=10)

	combine_button = tk.Button(frame,
	                   text="COMBINE",
	                   command=combine,
	                   bg= '#5A9D3A',
	                   highlightbackground='#5A9D3A')
	combine_button.grid(row=4, column=2, padx=10, pady=10)

	title = tk.Label(frame, text="HELLOFRESH DATA COMBINATRON 3000")
	title.grid(row=0, column=1, padx=10, pady=10)

	root.mainloop()
	