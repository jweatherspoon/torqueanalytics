#!/usr/bin/python
import sys
import time
import os
from datetime import datetime
import xlsxwriter

from optparse import OptionParser

from csvfile import *
from mail import *

DATE_FORMAT = "%d-%b-%Y %H:%M:%S.%f"

def main():
	
	parser = OptionParser(usage="""\
	Analyze a Torque logfile from either email or local directory
	Usage: %prog [options]
	""")
	parser.add_option('-f', '--file', type='string', action='store', metavar="FILE",
					  help="""The filename of the Torque logfile on the \
					  local machine. The email option is ignored if this argument \
					  is given""")
	parser.add_option('-e', '--email', action='store_true', metavar="EMAIL",
					  help="""Get Torque logfile from the most recent email \
					  residing in the default email inbox""")
	parser.add_option('-p', '--profile', type='string', action='store', metavar="PROFILE",
					  help="""Use a predefined profile for configuring the \
					  analytics tool""")
	
	opts, args = parser.parse_args()
	datafile = None
	
	if opts.file:
		datafile = CSVFile(opts.file)
	elif opts.email:
		GetMostRecentEmail()
		ExtractCSV()
		name = glob.glob("data/*.csv")
		datafile = CSVFile(name[0])
		os.system("rm -rf data")
	else:
		parser.print_help()
		sys.exit(1)
	
	profile = None
	if opts.profile:
		profile = LoadProfile(opts.profile)
	
	filename = "datafiles/" + os.path.basename(datafile.getFilename())[:-4] #Remove .csv from filename. Will be replaced by .xlsx
	Analyze(filename, datafile, profile)
		
def LoadProfile(profileName):
	if not profileName.startswith("profiles/"):
		profileName = "profiles/" + profileName
	fin = open(profileName)
	text = fin.read()
	text = text.split('\n')
	#Anything ending with a : is a column header
	#Anything beginning with a tab is a command
	profile = {}
	for i in range(len(text)):
		if text[i].endswith(":"):
			text[i] = text[i][:-1]
			data = []
			for j in range(i + 1, len(text)):
				if not text[j].startswith("\t"):
					break
				if text[j] != '\t':
					data.append(text[j][1:])
			profile[text[i]] = data
	return profile
				

def Analyze(filename, csvfile, profile=None):
	
	#Create the workbook
	if not filename.endswith(".xlsx"):
		filename += ".xlsx"
	workbook = xlsxwriter.Workbook(filename)
	print "Workbook filename is", filename
	
	for heading in csvfile.headers:
		if "Device Time" in heading:
			csvfile.convertAllToType(heading, datetime, DATE_FORMAT)
		else:
			csvfile.convertAllToType(heading, float)
			
	datasheet, columns = WriteDatasheet(workbook, csvfile)
	
	if profile is None:
		AnalyzeDefault(workbook, datasheet, csvfile)
	else:
		AnalyzeProfile(workbook, datasheet, csvfile, columns, profile)
	
	workbook.close()
	
def AnalyzeDefault(workbook, datasheet, csvfile):
	#Write the initial data as well as all the charts / graphs
	col = 0
	end = len(csvfile.getAllData("Device Time")) - 1
	
	chartsheets = []
	charts = {}

	for heading in csvfile.headers:
		columnData = csvfile.getAllData(heading)
		if "Device Time" not in heading:
			tab = cleanup(heading)
			chartsheets.append(workbook.add_chartsheet(tab.decode("utf-8")))
			charts[col - 1] = workbook.add_chart({
				"type": "scatter",
				"subtype": "smooth",
				"name": "Time v. " + heading
			})
			charts[col - 1].add_series({
				"categories": ["Data", 0, 0, end, 0],
				"values": ["Data", 0, col, end, col]
			})
			charts[col - 1].set_legend({"none": True})
			charts[col - 1].set_y_axis({"name": heading.decode("utf-8")})
			charts[col - 1].set_x_axis({"name": "Device Time", "date_axis": True, "num_format": "MM/DD/YY HH:MM:SS"})
			chartsheets[col - 1].set_chart(charts[col - 1])
		col += 1
		
def AnalyzeProfile(workbook, datasheet, csvfile, columns, profile):
	cIndex = 0
	end = len(csvfile.getAllData("Device Time")) - 1
	
	chartsheets = []
	charts = {}
	
	calcrow = 1
	calcsheet = workbook.add_worksheet("Calc")
	calcsheet.set_column(0,1,16)
	
	calcsheet.write(0,0,"Parameter")
	calcsheet.write(0,1,"Maximum")
	
	for heading in profile.keys():
		data = csvfile.getAllData(heading)
		for command in profile[heading]:
			if command == "Max":
				calcsheet.write(calcrow, 0, heading)
				calcsheet.write(calcrow, 1, max(data))
				calcrow += 1
			if command == "Graph":
				if heading in columns.keys():
					tab = cleanup(heading)
					chartsheets.append(workbook.add_chartsheet(tab.decode("utf-8")))
					charts[cIndex] = workbook.add_chart({
						"type": "scatter",
						"subtype": "smooth",
						"name": "Time v. " + heading
					})
					charts[cIndex].add_series({
						"categories": ["Data", 0, 0, end, 0],
						"values": ["Data", 0, columns[heading], end, columns[heading]]
					})
					charts[cIndex].set_legend({"none": True})
					charts[cIndex].set_y_axis({"name": heading.decode("utf-8")})
					charts[cIndex].set_x_axis({"name": "Device Time", "date_axis": True,
											   "num_format": "MM/DD/YY HH:MM:SS"})
					chartsheets[cIndex].set_chart(charts[cIndex])
					cIndex += 1
				
def WriteDatasheet(workbook, csvfile):
	columns = {}
	datasheet = workbook.add_worksheet("Data")
	datasheet.set_column(0,0,18)
	datasheet.set_column(1,len(csvfile.headers),16)
	dateForm = workbook.add_format({"num_format": "MM/DD/YY HH:MM:SS"})
	col = 0
	for heading in csvfile.headers:
		columns[heading] = col
		datasheet.write(0,col,heading.decode("utf-8"))
		columnData = csvfile.getAllData(heading)
		if "Device Time" in heading:
			datasheet.write_column(1,col,columnData,dateForm)
		else:
			datasheet.write_column(1,col,columnData)
		col += 1
	return datasheet, columns
	
def cleanup(item):
	item = item.replace('[','').replace(']','').replace('/','').replace('\\','').replace('*','').replace('?','').replace(':','')
	if len(item) > 31:
		item = item[0:30]
	return item
	
if __name__ == '__main__':
	main()