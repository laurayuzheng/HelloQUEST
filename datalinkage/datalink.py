import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

# file names
bob_name = '2019-03 BOB.xlsx'
humanity_name = '2019-03 Humanity.csv'
sel_name = '2019-03 SEL.xlsx'

# output files
bob_out = 'bob_out.csv'
humanity_out = 'humanity_out.csv'
sel_out = 'sel_out.csv'

# load data files
bob = pd.read_excel(bob_name, sheet_name=0)
humanity = pd.read_csv(humanity_name)
sel = pd.read_excel(sel_name, sheet_name=0)

# no need to filter hellofresh from bob anymore

# formatting dates in bob

bob['date_sign_up'] = pd.to_datetime(bob['date_sign_up'], format='%m/%d/%y %-H:%M')
bob['date_sign_up'] = bob['date_sign_up'].dt.strftime('%m/%d/%y')

# drop unnecessary columns
bob = bob.drop(['initial_sales_order_nr','customer_name'], axis=1)
 #, sheet_name='Sheet1')

# cleaning humanity sheet
humanity['start_day'] = pd.to_datetime(humanity['start_day'], format='%m/%d/%y')
humanity['start_day'] = humanity['start_day'].dt.strftime('%m/%d/%y')

#print(humanity['start_time'])
humanity['start_time'] = pd.to_datetime(humanity['start_time'],format='%I:%M %p')
humanity['start_time'] = humanity['start_time'].dt.strftime('%H:%M')
humanity['end_time'] = pd.to_datetime(humanity['end_time'],format='%I:%M %p')
humanity['end_time'] = humanity['end_time'].dt.strftime('%H:%M')
humanity = humanity.dropna(subset=['eid', 'location','shift_title'])

# merging humanity and bob
# replacing NaN values with 0
joined = pd.merge(bob, humanity, left_on=['Master Sales',
	'Office','Date'], right_on=['eid','location','start_day'])
joined = joined.fillna(0)

# TO DO: format time zones, get processed bob sheet from james

# output all to csv files
bob.to_csv('output/'+bob_out)
humanity.to_csv('output/'+humanity_out, index=False)
