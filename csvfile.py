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
		#Return -1 if the given key is not in the list of headers
		if key in self.headers:
			return self.rows[0].getIndexOfKey(key)
		else:
			return -1
		
	def printHeaders(self):
		print self.headers
		
	def printRow(self, row):
		#Only print the row if it is in the list of rows
		if row < len(self.rows):
			print self.rows[row]
			
	def getAllData(self, key):
		#Return a list of all data in the csvfile with the given header
		try:
			data = []
			for row in self.rows:
				data.append(row.get(key))
			return data
		except: #If not able, return Nonetype object
			return None
		
	def convertAllToType(self, key, typ, form=None):
		#Convert all values of the given key to given type (with given format)
		for i in range(len(self.rows)):
			self.rows[i].convert(key, typ, form)
				
				
	
	def __readFile(self):
		#Row 1 contains headers
		self.__getHeaders()
		
		#Read the rest and store in row objects
		csvreader = csv.reader(self.csv, delimiter=',')
		for values in csvreader:
			row = {}
			for i in range(len(values)):
				if values[i] == '-':
					values[i] = '0'
				row[self.headers[i]] = values[i]
			self.rows.append(Row(row))
			
		
	def __getHeaders(self):
		#Headers should be on first line. Get all characters until the \n
		headers = self.csv.readline()[:-1]
		headers = headers.split(',')
		for i in range(len(headers)):
			headers[i] = headers[i].lstrip()
		self.headers = headers
		
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
		'''
		Convert a value in the row to the given type. Currently supported
		types are datetime.datetime, int, and float. form parameter is
		only used for the datetime.datetime conversion.
		'''
		try:
			if typ == datetime:
				self.__data[key] = datetime.strptime(self.__data[key], form)
			elif typ == int:
				self.__data[key] = int(self.__data[key])
			elif typ == float:
				self.__data[key] = float(self.__data[key])
		except:
			pass
		
	def getIndexOfKey(self, key):
		'''
		Find the index of a given key. Return -1 if not found.
		'''
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
	