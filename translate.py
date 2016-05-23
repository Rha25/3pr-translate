import os
import argparse
import sys
import xml.etree.ElementTree as ET
import xlrd

global translations

# =========================
# operations on EXCEL file
# =========================
def openBook(filename):
	if not os.path.isfile(filename):
		print "[CRITICAL] file %s does not exist"%filename
		return None
	else:
		return xlrd.open_workbook(filename)

def openSheetName(sheetname):
	return translations.sheet_by_name(sheetname)

def getTranslation(original, from_sheet, lang_col):
	for row in range(1, from_sheet.nrows):
		if from_sheet.cell_value(row, 0) == original:
			return from_sheet.cell_value(row, lang_col)
	print "[WARNING] translation not found: '%s'"%original
	return ""
	
def get_column_index(language, sheet):
	for col in range(0, sheet.ncols):
		s = str(sheet.cell_value(0, col))
		if s.lower() == language.lower():
			return col
	print "[ERROR] missing language %s in translation file"%language
	return -1



# =========================
# operations on TS file
# =========================
def translate_ts(source, lang):
	print "[INFO] Loading TS input file", source
	try:
		tree = ET.parse(source)
		for context in tree.iter('context'):
			# for each context
			name = context.find('name').text
			sheet = openSheetName(name)
			col = get_column_index(lang, sheet)
			for message in context.iter('message'):
				original = message.find('source').text
				translation = getTranslation(original, sheet, col)
				if len(translation) > 0:
					#import pdb; pdb.set_trace()
					tr_tag = message.find('translation')
					tr_tag.set('type', 'finished')
					tr_tag.text = translation
		tree.write(get_output_filename(source, lang), 'UTF-8')
		print "[INFO] done."
	except ET.ParseError as err:
		print "[ERROR] Unable to parse input file", source
		sys.exit()
	except IOError:
		print "[ERROR] File", source, "unavailable"
		sys.exit()
	


# =========================
# operations on TXT file
# =========================
def translate_txt(source, lang):
	# in case of TXT file, the page is named after the filename
	the_sheet = openSheetName(os.path.basename(source))
	lang_column = get_column_index(lang, the_sheet)
	if lang_column != -1:
		print "[INFO] translating TXT input file...", source
		output = open(get_output_filename(source, lang), 'w')
		with open(source, 'r') as myfile:
			l = myfile.readline()
			while l != '':
				the_id = l.split('=')[0]
				the_string = l.split('=')[1].rstrip('\n')
				the_tran = getTranslation(the_string, the_sheet, lang_column)
				output.write(the_id + '=')
				if the_tran != "":
					output.write(the_tran.encode('UTF-8') + '\n')
				else:
					output.write(the_string + '\n')
				l = myfile.readline() # go on next line
		myfile.close()
		output.close()
		print "[INFO] done."



# =========================
# general operations
# =========================
def get_file_extension(pathname):
	l = os.path.basename(pathname).lower().split('.')
	return l[len(l) - 1] # the last part
	
def get_output_filename(source, lang):
	folder = os.path.dirname(source)
	fname = os.path.basename(source)
	return "%s/%s.%s"%(folder, lang, fname)

# =====
# MAIN
# =====
if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Translate a .ts file using .xlsx from Google Drive')
	parser.add_argument('source', type=str, help='source TS/TXT file, to be modified/translated')
	parser.add_argument('translations', type=str, help='xlsx containing translations')
	parser.add_argument('language', type=str, help='target output language')
	args = parser.parse_args()
	if not os.path.isfile(args.source) or not get_file_extension(args.source) in ['txt', 'ts']:
		print "[ERROR] You must input a valid source file with extension ts or txt"
		sys.exit()
	if get_file_extension(args.translations) != "xlsx":
		print "[ERROR] You must input a valid translations file with extension xlsx"
		sys.exit()
	
	global translations
	translations = openBook(args.translations)
	if get_file_extension(args.source) == 'txt':
		translate_txt(args.source, args.language)
	if get_file_extension(args.source) == 'ts':
		translate_ts(args.source, args.language)
