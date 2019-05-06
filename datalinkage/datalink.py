import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from datetime import datetime, timedelta
from pytz import timezone

def set_num_sales(row):
    if row['customer_id'] == 0:
        val = 0
    else:
        val = 1
    return val

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

# current month
month = '2019-03'

# average wage
avg_wage = 13.72

# file names
bob_name = 'mastersales.xlsx'
humanity_name = 'humanity_eventid.xlsx'
sel_name = '2019-03 SEL.xlsx'

# output files
bob_out = 'bob_out.xlsx'
humanity_out = 'humanity_out.xlsx'
sel_out = 'sel_out.xlsx'
joined_out = month + '_' + 'joined.xlsx'
free_events_out = month + '_' + 'free_events.xlsx'
nosale_out = month + '_' + 'no_sale.xlsx'
paid_events_out = month + '_' + 'paid_events.xlsx'

# load data files
bob = pd.read_excel('input/' + bob_name, sheet_name=0)
humanity = pd.read_excel('input/' + humanity_name)
sel = pd.read_excel('input/' + sel_name, sheet_name=0)

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

# standardize the column order
joined = joined[['eid','customer_id','event_id', 'num_sales',
				'office','box_type','date','start_time','end_time',
				'total_time','shift_title']]

nosales = joined[joined['customer_id'] == 0]
free_events = joined[joined['event_id'] == 0]
paid_events = joined[joined['event_id'] != 0]

# start here
sel = sel[['event_id','category','market','amortized_cost']]
per_shift = joined.groupby(['eid','shift_title','office','start_time','end_time',
	'date','event_id'], as_index=False).agg({'num_sales' : 'sum', 'total_time':'sum'})
# [['num_sales','total_time','box_type']].sum()
#print(per_shift)
per_shift['wage_cost'] = (avg_wage * per_shift['total_time']).round(decimals=2)

per_event = per_shift.groupby(['shift_title','office','event_id'], as_index=False).agg({'num_sales' : 'sum', 'total_time':'sum','wage_cost':'sum'})

combined = pd.merge(per_event, sel, on='event_id', how='outer')
combined = combined.fillna(0)
combined['total_cost'] = combined['wage_cost'] + combined['amortized_cost']

# output all to csv files
per_shift.to_excel('output/' + 'per_shift.xlsx', index=None)
per_event.to_excel('output/' + 'per_event.xlsx', index=None)
bob.to_excel('output/'+ bob_out)
humanity.to_excel('output/' + humanity_out)
joined.to_excel('output/' + joined_out, index=None)
nosales.to_excel('output/' + nosale_out, index=None)
free_events.to_excel('output/' + free_events_out, index=None)
paid_events.to_excel('output/' + paid_events_out, index=None)
combined.to_excel('output/' + 'combined.xlsx', index=None)