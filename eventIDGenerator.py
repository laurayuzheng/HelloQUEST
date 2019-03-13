import xlrd
import datetime
from random import randint

filename = "sales-expense-sample.xlsx"

wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0,0)

def createEventID(row_num, office):
	date = xlrd.xldate.xldate_as_datetime(sheet.cell(row_num-1,11).value,0).date().strftime('%m%d%y')
	pin = ''.join(str(randint(0,9)) for _ in range(4))
	output = office + date + pin
	return output

def main():
	print("main executed")
	row = input("Which row needs an event ID? ")
	office = input("What is the abbreviation of your office? ")
	row_num = int(row)
	output = createEventID(row_num, office)
	print(output)
	

if __name__ == '__main__':
	main()