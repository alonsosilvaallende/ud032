# -*- coding: utf-8 -*-
# Find the time and value of max load for each of the regions
# COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
# and write the result out in a csv file, using pipe character | as the delimiter.
# An example output can be seen in the "example.csv" file.
import xlrd
import os
import csv
# from zipfile import ZipFile
datafile = "2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


# def open_zip(datafile):
#     with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
#         myzip.extractall()


def parse_file(datafile):
	workbook = xlrd.open_workbook(datafile)
	sheet = workbook.sheet_by_index(0)
	# YOUR CODE HERE
	# Create an empty dictionary
	data = {}
	for col in range(1,9):
		# Create a list of the values for each column (excluding the header)
		column = sheet.col_values(col, start_rowx=1, end_rowx=range(sheet.nrows))
		# Get the maximum value and its index
		maxvalue = max(column)
		maxindex = column.index(maxvalue) + 1
		# Get the maximum time as a tuple
		maxtime = sheet.cell_value(maxindex, 0)
		realtime = xlrd.xldate_as_tuple(maxtime, 0)
		# Exclude minutes and seconds and add the maximum value
		realtime = realtime[0:4]
		realtime = realtime+(maxvalue,)
		data[sheet.cell_value(0,col)] = realtime
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)
	return data

def save_file(data, filename):
    # YOUR CODE HERE
	with open(filename, 'wb') as csvfile:
		linewriter = csv.writer(csvfile, delimiter='|')
		# Add the header
		header = ('Station', 'Year', 'Month', 'Day', 'Hour', 'Max Load')
		linewriter.writerow(header)
		for keys,values in data.items():
			# Write the name of the stations and the other values
			values = (keys,)+values
			linewriter.writerow(values)
	return filename
    
def test():
#    open_zip(datafile)
	data = parse_file(datafile)
	save_file(data, outfile)
	
	ans = {'FAR_WEST': {'Max Load': "2281.2722140000024", 'Year': "2013", "Month": "6", "Day": "26", "Hour": "17"}}
	
	fields = ["Year", "Month", "Day", "Hour", "Max Load"]
	with open(outfile) as of:
		csvfile = csv.DictReader(of, delimiter="|")
		for line in csvfile:
			s = line["Station"]
			if s == 'FAR_WEST':
				for field in fields:
					assert ans[s][field] == line[field]
        
test()
