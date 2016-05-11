#!usr/bin/python
import xlrd
import xlwt
import sys
import os
import re
import datetime
from datetime import timedelta
import pandas

def main():
	incidentID = 0

	prodcat2 = 13
	ticketVector = 14

	buildings = 11
	rooms = 12
	closedDates = 21


	inputfile = ''

	for i in range(0, len(sys.argv)):
		arg = sys.argv[i]
		if(arg == '-i'):
			inputfile = sys.argv[i+1]
		elif(arg == "-h"):
			print("analyze.py -i <inputtext>")

	book = xlrd.open_workbook(inputfile)
	sheet1 = book.sheet_by_index(0)

	print("\nNumber of worksheets detected: " + str(book.nsheets) + "\n")
	print(str(book.sheet_names()))
	headers = sheet1.row(0)
	#print(headers)
	headers_cleaned = []
	for header in headers:
		string = str(header)
		rx = re.compile('\W+')
		string1 = string.split(":")
		res = rx.sub('', string1[1]).split()
		#print(res)
		headers_cleaned.append(res)
	
	ticket_data = {}
	clean_ticket_data = {}

	r1 = re.compile('\W+')
	for i in range(0, len(headers_cleaned)):
		field = str(r1.sub('', str(headers_cleaned[i])))
		ticket_data[field] = sheet1.col(i)
		clean_ticket_data[field] = []
		#print(str(ticket_data[str(headers_cleaned[i])][12]).split(":"))
		if(i < len(headers_cleaned) - 1):

			for data in ticket_data[field]:
				data_header = r1.sub('', str(data).split(":")[1])
				header_cleaned = r1.sub('', str(headers_cleaned[i]))
				#print(data_header)
				#print(header_cleaned)
				if(r1.sub('', str(data).split(":")[1]) in r1.sub('', str(headers_cleaned[i]))):
					next
				else:
					string = str(data)
					rx = re.compile('\W+')
					string1 = string.split(":")
					res = rx.sub('', string1[1]).split()
					#print(res)
					clean_ticket_data[field].append(res)
		elif(i == len(headers_cleaned) - 1):
			for data in ticket_data[field]:
				#print(data)
				data_header = r1.sub('', str(data).split(":")[1])
				header_cleaned = r1.sub('', str(headers_cleaned[i]))
				#print(data_header)
				#print(header_cleaned)
				if(r1.sub('', str(data).split(":")[1]) in r1.sub('', str(headers_cleaned[i]))):
					next
				else:
					string = str(data)
					clean_ticket_data[field].append(string)

	for i in range(0, len(headers_cleaned) -1):
		string = str(headers_cleaned[i])
		headers_cleaned[i] = r1.sub('', string)
	headers_cleaned[len(headers_cleaned) - 1] = r1.sub('', str(headers_cleaned[len(headers_cleaned) - 1]))

	#print(clean_ticket_data[headers_cleaned[closedDates]])

	for date in clean_ticket_data[headers_cleaned[closedDates]]:
		string = str(date)
		rx = re.compile('/W+')
		string1 = string.split(":")
		string2 = string1[1].split()
		#print(string2[0])
		res = rx.sub('', string2[0]).split()
		#print(res)
		clean_ticket_data[headers_cleaned[closedDates]][clean_ticket_data[headers_cleaned[closedDates]].index(date)] = res

	#print(clean_ticket_data[headers_cleaned[closedDates]])

	Months = []
	Days = []
	Years = []
	for date in clean_ticket_data[headers_cleaned[closedDates]]:
		date_for_usage = r1.sub(' ', str(date))
		MM = ''
		DD = ''
		YYYY = ''
		counter = 0

		stripped_date = date_for_usage.split()
		Months.append(str(stripped_date[0]))
		Days.append(str(stripped_date[1]))
		Years.append(str(stripped_date[2]))

	for i in range(0, len(Months)):
		date = datetime.date(int(Years[i]), int(Months[i]), int(Days[i]))
		clean_ticket_data[headers_cleaned[closedDates]][i] = date

	#print(clean_ticket_data[headers_cleaned[closedDates]])

	today = datetime.date.today()
	print("\nToday's date is " + str(today.month) + "/" + str(today.day) + "/" + str(today.year))
	days_of_week_non_iso = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
	print("Today's day of the week is " + str(days_of_week_non_iso[today.weekday()]))

	end_period = today - timedelta(today.isoweekday() + 1)
	start_period = today - timedelta(today.isoweekday() + 7)

	print("The reporting period is " + str(start_period.month)+ "/" + str(start_period.day) + "/" + str(start_period.year) + " - " + str(end_period.month) + "/" + str(end_period.day) + "/" + str(end_period.year))

	#print(clean_ticket_data[headers_cleaned[closedDates]][len(clean_ticket_data[headers_cleaned[closedDates]]) - 1])
	total_obs = len(clean_ticket_data[headers_cleaned[incidentID]])
	print(str(len(clean_ticket_data[headers_cleaned[incidentID]])) + " incidents were observed in this dataset")
	
	#cleaning function
	'''
	for i in range(0, len(headers_cleaned) - 1):
		counter_var = 0
		for data in clean_ticket_data[headers_cleaned[i]]:
			
			if(counter_var < total_obs - 1):

				print(clean_ticket_data[headers_cleaned[i]].index(data))
				print("\n" + str(clean_ticket_data[headers_cleaned[i+1]]))

				if(clean_ticket_data[headers_cleaned[i+ 1]][counter_var]):
					next
				else:
					next
					clean_ticket_data[headers_cleaned[i +1].insert(counter_var, "NA")
			else: next
	'''
	count_below =  0
	count_above = 0
	count = 0
	for date in clean_ticket_data[headers_cleaned[closedDates]]:
		if(date > end_period):
			count_above += 1
			for field in headers_cleaned:
				try:
					clean_ticket_data[field].pop(clean_ticket_data[headers_cleaned[closedDates]].index(date))
				except:
					pass
		elif(date < start_period):
			count_below += 1
			for field in headers_cleaned:
				try:
					clean_ticket_data[field].pop(clean_ticket_data[headers_cleaned[closedDates]].index(date))
				except:
					pass
		else:
			count += 1

	print("\nOf " + str(total_obs) + " incidents observed, " + str(count) + " were within the reporting period")
	print(str(count_below + count_above) + " incidents were outside the reporting period")
	for field in headers_cleaned:
		print(str(field) + " had " + str(len(clean_ticket_data[field])) + " entries")
	df = pandas.DataFrame(clean_ticket_data)
	#for field in headers_cleaned:
	#	try:
	#		clean_ticket_data[field].pop()
	#	except:
	#		pass


main()