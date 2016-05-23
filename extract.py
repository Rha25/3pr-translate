import os
import argparse
import sys
import math
import xml.etree.ElementTree as ET
import xlwt, xlrd, xlutils
from xlutils.copy import *

languages = ['italian', 'japanese', 'chinese', 'german', 'spanish', 'polish', 'french'] # add more
donottranslate = []
loadedpages = list()

#=====================
# STYLE for the header
#=====================
headerstyle = xlwt.XFStyle()
font = xlwt.Formatting.Font()
font.bold = True
pattern = xlwt.Formatting.Pattern()
#pattern.pattern = xlwt.Formatting.Pattern.SOLID_PATTERN
#pattern.pattern_back_colour = 11
#headerstyle.pattern = pattern
headerstyle.font = font


#=====================
# A Page in xls file
#=====================
class XLSPage:
	def __init__(self, title):
		self.strings = list()
		self.title = title
		
	def __str__(self):
		return "Page: %s {%d elements}"%(self.title, len(self.strings))
	
	def write_to(self, page):
		pagechars = 0
		page.write(0, 0, "Original String", headerstyle)
		page.write(0, 1, "Max. Length", headerstyle)
		page.write(0, 2, "Description", headerstyle)
		row = 1
		maxchars = 0
		for line in self.strings:
			page.write(row, 0, line)
			pagechars += len(line)
			page.write(row, 1, len(line))
			h, w = dimensions(line)
			if w > maxchars:
				maxchars = w
			# adjust row height for multiple lines
			page.row(row).height = page.row(row).height * h
			row += 1
		column = 4 # translations start
		for header in languages:
			page.write(0, column, header.upper(), headerstyle)
			column += 1
		for col in range(3, column):
			page.col(col).set_width(maxchars * 255)
		page.col(0).set_width(maxchars * 255)
		page.col(2).set_width(5000)
		page.col(1).set_width(5000)
		return pagechars
			
	def add_source_string(self, string):
		self.strings.append(string.rstrip('\n')) # remove the last newline char
		
def dimensions(string):
	# count lines + max length
	lines = string.count('\n') + 1
	maxlen = 0
	for s in string.split('\n'):
		if len(s) > maxlen:
			maxlen = len(s)
	return lines, maxlen
		
# =======
# METHODS
# =======

def loadQtTS(tsfile):
	print "[INFO] Loading TS input file", tsfile
	try:
		pages = list()
		tree = ET.parse(tsfile)
		for context in tree.iter('context'):
			# for each context
			page = None
			for el in context:
				if el.tag == 'name':
					page = XLSPage(el.text)
				elif el.tag == 'message':
					page.add_source_string(el.find('source').text)
			loadedpages.append(page)
	except ET.ParseError as err:
		print "[ERROR] Unable to parse input file", tsfile
		sys.exit()
	except IOError:
		print "[ERROR] File", tsfile, "unavailable"
		sys.exit()
		
def loadXML(xmlfile):
	print "[INFO] Loading XML input file", xmlfile
	try:
		tree = ET.parse(xmlfile)
		page = XLSPage(os.path.basename(xmlfile))
		for alarm in tree.iter('ALARM'):
			# for each context
			# MESSAGE, DESCRIPTION, HELP, NAME
			message = alarm.find('MESSAGE')
			description = alarm.find('DESCRIPTION')
			name = alarm.find('NAME')
			help = alarm.find('HELP')
			if message is not None:
				page.add_source_string(message.text)
			if name is not None:
				page.add_source_string(name.text)
			if description is not None:
				page.add_source_string(description.text)
			if help is not None:
				page.add_source_string(help.text)
		loadedpages.append(page)
	except ET.ParseError as err:
		print "[ERROR] Unable to parse input file", tsfile
		sys.exit()
	except IOError:
		print "[ERROR] File", tsfile, "unavailable"
		sys.exit()
		
def loadCSV(csvfile):
	# TODO
	pass
	
def loadTXT(txtfile):
	# make just one page
	print "[INFO] Loading TXT input file", txtfile
	p = XLSPage(os.path.basename(txtfile))
	with open(txtfile, 'r') as myfile:
		l = myfile.readline()
		while l != '':
			p.add_source_string(l.split('=')[1].rstrip('\n'))
			l = myfile.readline() # go on next line
	myfile.close()
	loadedpages.append(p)

def writeXLS(xlsfile):
	book = xlwt.Workbook()
	characters = 0
	if os.path.exists(xlsfile):
		readbook = xlrd.open_workbook(xlsfile)
		# FIXME: here it loses the format
		book = copy(readbook) # xlutils.copy
	for page in loadedpages:
		try:
			sheet = book.add_sheet(page.title)
		except Exception as ex:
			print "[WARNING]", ex.message
			print "[INFO] Updating strings on the page..."
			rb = xlrd.open_workbook(xlsfile)
			index = 0
			for s in rb.sheets():
				if s.name == page.title:
					break
				index += 1
			sheet = book.get_sheet(index)
		finally:
			characters += page.write_to(sheet)
	book.save(xlsfile)
	print "[INFO] Export done. Check output file %s (%d characters, %d cartels)"%(xlsfile, characters, math.ceil(characters/1375.0))
	
def get_file_extension(pathname):
	l = os.path.basename(pathname).lower().split('.')
	return l[len(l) - 1] # the last part
	
def load_file(pathname):
	ext = get_file_extension(pathname)
	# read from the input file
	if ext == "txt":
		loadTXT(pathname)
	elif ext == "ts":
		loadQtTS(pathname)
	elif ext == "xml":
		loadXML(pathname)
	elif ext == "csv":
		loadCSV(pathname)
		
		
# =====
# MAIN
# =====
if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Convert xlsm/xls into menu xml file')
	parser.add_argument('source', type=str, help='source file, to be translated')
	parser.add_argument('output', type=str, help='xls target file (existing or not)')
	args = parser.parse_args()
	if not os.path.isfile(args.source):
		print "[ERROR] You must input a valid source file with extension xml, ts (xml), csv, txt"
		sys.exit()
	if get_file_extension(args.output) != "xls":
		print "[ERROR] You must input a valid output file with extension xls"
		sys.exit()
	load_file(args.source)
	# write the output
	writeXLS(args.output)
