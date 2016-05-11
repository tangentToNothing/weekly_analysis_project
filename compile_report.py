#!usr/bin/python
import xlrd
import xlwt
import sys
import os
import re
import datetime
from datetime import timedelta
import pandas
import pprint

def main():

	incidentID = 0				#Field Definitions for ticket categorizations' headers
	prodcat2 = 13				#
	ticketVector = 14			#
	buildings = 11				#
	rooms = 12					#
	closedDates = 21			#

	inputfile = ''				#input file holder

	for i in range(0, len(sys.argv)):					#standard argument processing
		arg = sys.argv[i]
		if(arg == '-i'):
			inputfile = sys.argv[i+1]
		elif(arg == "-h"):
			print("analyze.py -i <inputtext>")

	book = xlrd.open_workbook(inputfile)										#open workbook
	sheet1 = book.sheet_by_index(0)												#parse first sheet

	print("\nNumber of worksheets detected: " + str(book.nsheets) + "\n")		#output number of sheets
	print(str(book.sheet_names()))												#further sheet details
	
	headers = sheet1.row(0)											#header processing
	headers_cleaned = []											#similar to csv work
	for header in headers:											#
		string = str(header)										#
		rx = re.compile('\W+')										#
		string1 = string.split(":")									#
		res = rx.sub('', string1[1]).split()						#
		headers_cleaned.append(res)									#
	
	ticket_data = {}									#placeholder
	clean_ticket_data = {}								#true data frame equivalent

	r1 = re.compile('\W+')																				#pull all of the data using
	for i in range(0, len(headers_cleaned)):															#double nested for loops
		field = str(r1.sub('', str(headers_cleaned[i])))												#
		ticket_data[field] = sheet1.col(i)																#
		clean_ticket_data[field] = []																	#
		#print(str(ticket_data[str(headers_cleaned[i])][12]).split(":"))								#
		if(i < len(headers_cleaned) - 1):																#
			for data in ticket_data[field]:																#
				data_header = r1.sub('', str(data).split(":")[1])										#
				header_cleaned = r1.sub('', str(headers_cleaned[i]))									#
				if(r1.sub('', str(data).split(":")[1]) in r1.sub('', str(headers_cleaned[i]))):			#
					next																				#
				else:																					#
					string = str(data)																	#
					rx = re.compile('\W+')																#
					string1 = string.split(":")															#
					res = rx.sub('', string1[1]).split()												#
					#print(res)																			#
					clean_ticket_data[field].append(res)												#
		elif(i == len(headers_cleaned) - 1):															#
			for data in ticket_data[field]:																#
				#print(data)																			#
				data_header = r1.sub('', str(data).split(":")[1])										#
				header_cleaned = r1.sub('', str(headers_cleaned[i]))									#
				if(r1.sub('', str(data).split(":")[1]) in r1.sub('', str(headers_cleaned[i]))):			#
					next																				#
				else:																					#
					string = str(data)																	#
					clean_ticket_data[field].append(string)												#

	for i in range(0, len(headers_cleaned) -1):																#Further header processing
		string = str(headers_cleaned[i])																	#
		headers_cleaned[i] = r1.sub('', string)																#
	headers_cleaned[len(headers_cleaned) - 1] = r1.sub('', str(headers_cleaned[len(headers_cleaned) - 1]))	#

	#print(clean_ticket_data[headers_cleaned[closedDates]])

	for date in clean_ticket_data[headers_cleaned[closedDates]]: 											#Regular expression date cleaning
		string = str(date)																					#
		rx = re.compile('/W+')																				#
		string1 = string.split(":")																			#
		string2 = string1[1].split()																		#	
		res = rx.sub('', string2[0]).split()																#
		#print(res)
		clean_ticket_data[headers_cleaned[closedDates]][clean_ticket_data[headers_cleaned[closedDates]].index(date)] = res

	Months = []																			#extract time info
	Days = []																			#-using regular expressions
	Years = []																			#
	for date in clean_ticket_data[headers_cleaned[closedDates]]:						#
		date_for_usage = r1.sub(' ', str(date))											#
		MM = ''																			#
		DD = ''																			#
		YYYY = ''																		#
		counter = 0																		#
		stripped_date = date_for_usage.split()											#
		Months.append(str(stripped_date[0]))											#
		Days.append(str(stripped_date[1]))												#
		Years.append(str(stripped_date[2]))												#

	for i in range(0, len(Months)):														#convert time info into datetime objects
		date = datetime.date(int(Years[i]), int(Months[i]), int(Days[i]))				#
		clean_ticket_data[headers_cleaned[closedDates]][i] = date 						#

	#print(clean_ticket_data[headers_cleaned[closedDates]])

	today = datetime.date.today()																				#Extract today's date
	print("\nToday's date is " + str(today.month) + "/" + str(today.day) + "/" + str(today.year))				#and alert the user
	days_of_week_non_iso = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")		#Along with day of the week
	print("Today's day of the week is " + str(days_of_week_non_iso[today.weekday()]))							#

	end_period = today - timedelta(today.isoweekday() + 1)														#also establish reporting period datetime objects
	start_period = today - timedelta(today.isoweekday() + 7)													#and alert user \/

	print("The reporting period is " + str(start_period.month)+ "/" + str(start_period.day) + "/" + str(start_period.year) + " - " + str(end_period.month) + "/" + str(end_period.day) + "/" + str(end_period.year))

	#print(clean_ticket_data[headers_cleaned[closedDates]][len(clean_ticket_data[headers_cleaned[closedDates]]) - 1])
	total_obs = len(clean_ticket_data[headers_cleaned[incidentID]])
	print(str(len(clean_ticket_data[headers_cleaned[incidentID]])) + " incidents were observed in this dataset")	#total # of incidents
	
	'''
	pp = pprint.PrettyPrinter(indent=4)
	for i in range(0, len(headers_cleaned) - 1):
		
		for j in range(0, total_obs - 10):
			
			#print(clean_ticket_data[headers_cleaned[i]].index(data))
			#pp.pprint("\n" + str(clean_ticket_data[headers_cleaned[i+1]]))
			print(j)
			print(i)
			if(clean_ticket_data[headers_cleaned[i+ 1]][j]):
				next
			else:
				clean_ticket_data[headers_cleaned[i +1]].insert(j, "NA")
	'''	
				
	count_below = 0																								#
	count_above = 0																								#truncate data
	count = 0																									#
	for date in clean_ticket_data[headers_cleaned[closedDates]]:												#
		if(date > end_period):																					#
			count_above += 1																					#
			for field in headers_cleaned:																		#
				try:																							#
					clean_ticket_data[field].pop(clean_ticket_data[headers_cleaned[closedDates]].index(date))	#using pop
				except:																							#
					pass																						#
		elif(date < start_period):																				#
			count_below += 1																					#
			for field in headers_cleaned:																		#
				try:																							#
					clean_ticket_data[field].pop(clean_ticket_data[headers_cleaned[closedDates]].index(date))	#
				except:																							#
					pass																						#
		else:																									#
			count += 1																							#

	print("\nOf " + str(total_obs) + " incidents observed, " + str(count) + " were within the reporting period")	#Alert user of clipping
	print(str(count_below + count_above) + " incidents were outside the reporting period")							#
	
	for field in headers_cleaned:																#investigate size of fields
		print(str(field) + " had " + str(len(clean_ticket_data[field])) + " entries")			#
	df = pandas.DataFrame(clean_ticket_data)
	#for field in headers_cleaned:
	#	try:
	#		clean_ticket_data[field].pop()
	#	except:
	#		pass


main()