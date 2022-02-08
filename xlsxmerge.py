import sys
import argparse
import pandas as pd

# imports file f starting from a data array map d, by setting up new new_key key and by using key_string as key identifier, key_prepend 
def importfile(f, k, key_string, key_prepend, d={}):
	df = pd.read_excel(f)
	headers = df.columns
	data = df.to_numpy().tolist()
	key_index = None
	for i in range(0, len(headers)):
		if headers[i] == key_string: key_index = i
	if key_index is None: return None
	for e in data:
		o = {}
		new_key = None
		for i in range(0, len(headers)):
			if i == key_index:
				new_key = str(e[i])
				o[k] = new_key
				continue
			kk = str(headers[i])
			if not key_prepend is None: kk = '{} - {}'.format(key_prepend, str(headers[i]))
			o[kk] = str(e[i])
		if d.get(new_key) is None: d[new_key] = {}
		for kk in o: d[new_key][str(kk)] = o[kk]
	return d

# gets all headers found in d
def getheaders(d):
	r = []
	for k in d:
		for k2 in d.get(k):
			if k2 in r: continue
			r.append(k2)
	return r

# checks that all inputs are compliant
def checkinputs(i1, i2, i3):
	l = len(i1)
	return (len(i2) == l and len(i3) == l)

# management of input arguments
parser = argparse.ArgumentParser()
parser.add_argument('-f','--files-list', nargs='+', help='<Required> Set flag', required=True)
parser.add_argument('-k','--keys-list', nargs='+', help='<Required> Set flag', required=True)
parser.add_argument('-t','--tags-list', nargs='+', help='<Not required> Set flag', required=False)
parser.add_argument('-n','--new-key', help='<Required> Set flag', required=True)
parser.add_argument('-o','--output-file', help='<Required> Set the output file name', default='merged.xlsx', required=False)
args = vars(parser.parse_args())
filenames = args.get('files_list')
keys = args.get('keys_list')
tags = args.get('tags_list')
new_key = args.get('new_key')
output_file = args.get('output_file')
if tags is None: tags = [ None ] * len(filenames)

# additional input checking
if not checkinputs(filenames, keys, tags):
	print('All variables (filenames, keys, tags) has to be of the same length (use None values to ignore)')
	sys.exit(0)

# importing all files and generating the overall d array map
d = {}
for i in range(0, len(filenames)):
	d = importfile(filenames[i], new_key, keys[i], tags[i], d)

# getting all headers
headers = getheaders(d)

# writing to excel

import xlwt
from xlwt import Workbook

wb = Workbook()
s = wb.add_sheet('Merged data')

# writes a single row in the excel sheet
def writerow(sheet, row, headers, element=None):
	for i in range(0, len(headers)):
		if element is None:
			sheet.write(row, i, headers[i])
			continue
		sheet.write(row, i, element.get(headers[i]))

# writing headers
row = 0
writerow(s, row, headers)
# writing contents
for e in d:
	row += 1
	writerow(s, row, headers, d.get(e))

# generating the output file
wb.save(output_file)
