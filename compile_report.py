#!usr/bin/python
import xlrd
import xlwt
import sys
import os
import re

def main():

	inputfile = ''

	for i in range(0, len(sys.argv)):
		arg = sys.argv[i]
		if(arg == '-i'):
			inputfile = sys.argv[i+1]
		elif(arg == "-h"):
			print("analyze.py -i <inputtext>")

	book = xlrd.open_workbook(inputfile)
	sheet1 = book.sheet_by_index(0)

	print("The number of worksheets is " + str(book.nsheets))
	print(str(book.sheet_names()))
	headers = sheet1.row(0)
	print(headers)
	headers_cleaned = []
	for header in headers:
		string = str(header)
		rx = re.compile('\W+')
		string1 = string.split(":")
		res = rx.sub('', string1[1]).split()
		print(res)
		headers_cleaned.append(res)

	#this needs some work.  Dates need to be dropped in to datetime objects somehow
	dates = sheet1.col(len(headers_cleaned) - 1)
	for date in dates:
		string = str(date)
		rx = re.compile('/W+')
		string1 = string.split(":")
		string2 = string1[1].split()
		print(string2[0])
		res = rx.sub('', string2[0]).split()
		print(res)
main()