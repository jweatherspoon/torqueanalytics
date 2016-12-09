#!/usr/bin/python
import sys
import time
import os
import shutil
import zipfile
from datetime import datetime
import xlsxwriter

from optparse import OptionParser

from csvfile import *
from mail import *

DATE_FORMAT = "%d-%b-%Y %H:%M:%S.%f"

def main():
	
	#Create possible options and add help messages 
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
	
	#Get the used arguments
	opts, args = parser.parse_args()
	datafile = None
	
	#Get the corresponding file from the argument. local file has priority over email.
	#If neither file nor email, print help message. One is required.
	if opts.file:
		datafile = CSVFile(opts.file)
	elif opts.email:
		GetMostRecentEmail()
		ExtractCSV()
		name = glob.glob("data/*.csv")
		datafile = CSVFile(name[0])
		shutil.rmtree("data")
		#os.system("rm -rf data")
	else:
		parser.print_help()
		sys.exit(1)
	
	#Attempt to load a profile if one is given
	profile = None
	if opts.profile:
		profile = LoadProfile(opts.profile)
	
	#Use the same filename as the .csv just change to .xlsx
	#Then analyze the file. (graph / calculate)
	filename = "datafiles/" + os.path.basename(datafile.getFilename())[:-4] #Remove .csv from filename. Will be replaced by .xlsx
	Analyze(filename, datafile, profile)
		
		
def LoadProfile(profileName):
	#Get the profile from the profiles directory
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
	
	#Convert all data to their corresponding value
	#Device time -> datetime
	#All else -> float
	for heading in csvfile.headers:
		if "Device Time" in heading:
			csvfile.convertAllToType(heading, datetime, DATE_FORMAT)
		else:
			csvfile.convertAllToType(heading, float)
		
	#Create the datasheet and get a list of headings with their respective column number
	datasheet, columns = WriteDatasheet(workbook, csvfile)
	
	#Graph all headings vs. time. No max / min
	if profile is None:
		AnalyzeDefault(workbook, datasheet, csvfile)
	else: #Analyze based on the loaded profile
		AnalyzeProfile(workbook, datasheet, csvfile, columns, profile)
	
	workbook.close()
	
def AnalyzeDefault(workbook, datasheet, csvfile):
	#Write the initial data as well as all the charts / graphs
	col = 0
	end = len(csvfile.getAllData("Device Time")) - 1
	
	chartsheets = []
	charts = {}

	#Graph each heading vs. device time.
	for heading in csvfile.headers:
		columnData = csvfile.getAllData(heading)
		if "Device Time" not in heading:
			#Create a new chart for this column
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
			#Get rid of the legend, set axis titles. Set x axis to date format.
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
	
	#Create a sheet for calculations like min / max / avg
	calcrow = 1
	calcsheet = workbook.add_worksheet("Calc")
	calcsheet.set_column(0,0,maxLen(profile.keys()))
	calcsheet.set_column(1,3,16)
	
	#Write table for calculations
	data = ["Parameter","Maximum","Minimum","Average"]
	calcsheet.write_row(0,0,data)
	
	parameters = {}
	
	
	#For each heading in the profile, do the given command.
	for heading in profile.keys():
		data = csvfile.getAllData(heading)
		if heading in columns:
			parameters[heading] = calcrow
			calcrow += 1
			calcsheet.write(parameters[heading], 0, heading.decode("utf-8"))
			
			if "Max" in profile[heading]:
				calcsheet.write(parameters[heading], 1, max(data))				
			if "Min" in profile[heading]:
				temp = [val for val in data if val != 0] #Remove all 0s from the list
				if len(temp):
					calcsheet.write(parameters[heading], 2, min(temp))
			if "Avg" in profile[heading] or "Average" in profile[heading]:
				calcsheet.write(parameters[heading], 3, mean(data))
			if "Graph" in profile[heading]:
				#Create a new chart for this column. Graph it against the device time
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
				#Remove legend, set axis titles, format x axis as date.
				charts[cIndex].set_legend({"none": True})
				charts[cIndex].set_y_axis({"name": heading.decode("utf-8")})
				charts[cIndex].set_x_axis({"name": "Device Time", "date_axis": True,
										   "num_format": "MM/DD/YY HH:MM:SS"})
				chartsheets[cIndex].set_chart(charts[cIndex])
				cIndex += 1
		'''
		for command in profile[heading]:
			
			if heading in columns:
				#Add the max of this column to the calcsheet
				if command == "Max":
					calcsheet.write(calcrow, 0, heading)
					calcsheet.write(calcrow, 1, max(data))
					calcrow += 1
				if command == "Min":
					
				if command == "Graph":
					#Create a new chart for this column. Graph it against the device time
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
					#Remove legend, set axis titles, format x axis as date.
					charts[cIndex].set_legend({"none": True})
					charts[cIndex].set_y_axis({"name": heading.decode("utf-8")})
					charts[cIndex].set_x_axis({"name": "Device Time", "date_axis": True,
											   "num_format": "MM/DD/YY HH:MM:SS"})
					chartsheets[cIndex].set_chart(charts[cIndex])
					cIndex += 1
			'''
				
def WriteDatasheet(workbook, csvfile):
	columns = {}
	#Create a new worksheet for the entire dataset.
	datasheet = workbook.add_worksheet("Data")
	#Set column widths.
	datasheet.set_column(0,0,18)
	datasheet.set_column(1,len(csvfile.headers),16)
	dateForm = workbook.add_format({"num_format": "MM/DD/YY HH:MM:SS"})
	col = 0
	#For every header from the csv file, write the column.
	#Write device time as date format MM/DD/YY HH:MM:SS
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
	
#Clean up some text for use as a spreadsheet tab
def cleanup(item):
	item = item.replace('[','').replace(']','').replace('/','').replace('\\','').replace('*','').replace('?','').replace(':','')
	if len(item) > 31:
		item = item[0:30]
	return item
								
def mean(list):
	return float(sum(list) / max(len(list), 1))
	
def maxLen(list):
	maxLen = 0
	for item in list:
		if len(item) > maxLen:
			maxLen = len(item)
	return maxLen
	
if __name__ == '__main__':
	main()