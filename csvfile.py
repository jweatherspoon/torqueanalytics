import csv

class CSVFile:
	def __init__(self, filename):
		self.csv = open(filename, "r")
		self.headers = []
		self.rows = []
		
		self.__readFile()
		self.csv.close()
		
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
	
	def __readFile(self):
		#Row 1 contains headers
		self.__getHeaders()
		
		csvreader = csv.reader(self.csv, delimiter=',')
		for values in csvreader:
			row = {}
			for i in range(len(values)):
				row[self.headers[i]] = values[i]
			self.rows.append(row)
			
		
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
		
	def __str__(self):
		s = ' '.join(self.__data)
		return s
	