#!/usr/bin/python
import sys
import time
from datetime import datetime
import xlsxwriter

from csvfile import *

def main():
	if len(sys.argv) < 2:
		print "USAGE:", sys.argv[0], "<filename> [,--<profile>]"
		sys.exit(1)
	
	datafile = CSVFile(sys.argv[1])
	
	profile = None
	
	if len(sys.argv) > 2:
		profile = LoadProfile(sys.argv[2])
		
	filename = "datafiles/" + str(datetime.now()).split(' ')[1] + ".xlsx"
	Analyze(filename, datafile, profile)
		
def LoadProfile(profileName):
	pass

def Analyze(filename, csvfile, profile=None):	
	#Create the workbook
	if not filename.endswith(".xlsx"):
		filename += ".xlsx"
	workbook = xlsxwriter.Workbook(filename)
	datasheet = workbook.add_worksheet("Data")
	
	#Write the initial data as well as all the charts / graphs
	chartsheets = []
	col = 0
	for heading in csvfile.headers:
		datasheet.write(0,col, heading.decode("utf-8"))
		columnData = csvfile.getAllData(heading)
		row = 1
		for item in columnData:
			datasheet.write(row, col, item.encode("utf-8"))
			row += 1
		
		if "Device Time" not in heading:
			#Create a chart vs device time
			timeData = csvfile.getAllData("Device Time")
			
			tab = cleanup(heading)
			
			chartsheet = workbook.add_chartsheet(tab.decode("utf-8"))
			chart = workbook.add_chart({"type": "line"})
			
			end = len(timeData) - 1;
			chart.add_series({
				"categories": ["Data", 1, 1, end, end],
				"values": ["Data", col, col, end, end]
			})				
			
			chartsheet.set_chart(chart)
		
		col += 1
	
	
	if profile is not None:
		pass
	
	workbook.close()
	
def cleanup(item):
	item = item.replace('[','').replace(']','').replace('/','').replace('\\','').replace('*','').replace('?','').replace(':','')
	if len(item) > 31:
		item = item[0:30]
	return item
	
if __name__ == '__main__':
	main()