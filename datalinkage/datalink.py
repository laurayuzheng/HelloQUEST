import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from datetime import datetime, timedelta
from pytz import timezone

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

# file names
bob_name = 'mastersales.xlsx'
humanity_name = 'humanity_eventid.xlsx'
sel_name = '2019-03 SEL.xlsx'

# output files
bob_out = 'bob_out.xlsx'
humanity_out = 'humanity_out.xlsx'
sel_out = 'sel_out.xlsx'

# load data files
bob = pd.read_excel('input/' + bob_name, sheet_name=0)
humanity = pd.read_excel('input/' + humanity_name)
sel = pd.read_excel('input/' + sel_name, sheet_name=0)

# no need to filter hellofresh from bob anymore

# formatting dates in bob

bob['date_sign_up'] = pd.to_datetime(bob['date_sign_up'], format='%m/%d/%y %-H:%M')
bob['date_sign_up'] = bob['date_sign_up'].dt.strftime('%m/%d/%y')

# drop unnecessary columns
# bob = bob.drop(['initial_sales_order_nr','customer_name'], axis=1)
 #, sheet_name='Sheet1')

# cleaning humanity sheet
humanity['start_day'] = pd.to_datetime(humanity['start_day'], format='%m/%d/%y')
humanity['start_day'] = humanity['start_day'].dt.strftime('%m/%d/%y')

#print(humanity['start_time'])
#print(humanity['start_time'])
humanity['start_time'] = pd.to_datetime(humanity['start_time'],format='%H:%M:%S')
#humanity['start_time'] = humanity['start_time'].dt.strftime('%H:%M')
humanity['end_time'] = pd.to_datetime(humanity['end_time'],format='%H:%M:%S')
#humanity['end_time'] = humanity['end_time'].dt.strftime('%H:%M')
humanity = humanity.dropna(subset=['eid', 'location','shift_title'])

# localize the time zones, then set time zones to eastern
# then, convert the new time zone to a string
for i, row in humanity.iterrows():
	tz = timezones[row['location']]
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
joined = pd.merge(bob, humanity, left_on=['Employee ID',
	'Office','date_sign_up'], right_on=['eid','location','start_day'])
joined = joined.fillna(0)

# TO DO: format time zones


# output all to csv files
bob.to_excel('output/'+ bob_out)
humanity.to_excel('output/' + humanity_out)
joined.to_excel('output/joined.xlsx', index=None)
