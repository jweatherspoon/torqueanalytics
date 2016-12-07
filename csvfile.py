import csv
from datetime import datetime

class CSVFile:
	def __init__(self, filename):
		self.csv = open(filename, "r")
		self.filename = filename
		self.headers = []
		self.rows = []
		
		self.__readFile()
		self.csv.close()
		
	def getFilename(self):
		return self.filename
		
	def getIndexOfKey(self, key):
		if key in self.headers:
			return self.rows[0].getIndexOfKey(key)
		else:
			return -1
		
	def printHeaders(self):
		print self.headers
		
	def printRow(self, row):
		if row < len(self.rows):
			print self.rows[row]
			
	def getAllData(self, key):
		try:
			data = []
			for row in self.rows:
				data.append(row.get(key))
			return data
		except:
			return None
		
	def convertAllToType(self, key, typ, form=None):
		for i in range(len(self.rows)):
			self.rows[i].convert(key, typ, form)
				
				
	
	def __readFile(self):
		#Row 1 contains headers
		self.__getHeaders()
		
		csvreader = csv.reader(self.csv, delimiter=',')
		for values in csvreader:
			row = {}
			for i in range(len(values)):
				if values[i] == '-':
					values[i] = '0'
				row[self.headers[i]] = values[i]
			self.rows.append(Row(row))
			
		
	def __getHeaders(self):
		headers = self.csv.readline()[:-1]
		headers = headers.split(',')
		for i in range(len(headers)):
			headers[i] = headers[i].lstrip()
		self.headers = headers
			
	def __readLine(self):
		try:
			line = self.csv.readline()[:-1]
			line = line.split(',')
			for i in range(len(line)):
				line[i] = line[i].lstrip()
			return line
		except:
			return None
		
class Row:
	def __init__(self, data):
		self.__data = data
		
	def set(self, key, val):
		self.__data[key] = val
	
	def get(self, key):
		try:
			return self.__data[key]
		except:
			return None
	
	def remove(self, key):
		try:
			self.__data.remove(key)
		except:
			pass
	
	def convert(self, key, typ, form=None):
		if typ == datetime:
			self.__data[key] = datetime.strptime(self.__data[key], form)
		elif typ == int:
			self.__data[key] = int(self.__data[key])
		elif typ == float:
			self.__data[key] = float(self.__data[key])
			
	def getIndexOfKey(self, key):
		keys = self.__data.keys()
		for i in range(len(keys)):
			if key == keys[i]:
				return i
		return -1
		
	def __str__(self):
		s = ' '.join(self.__data)
		return s
	
	def __getitem__(self, key):
		return self.get(key)
	