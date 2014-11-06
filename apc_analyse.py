#!/usr/bin/env python

import os
import getopt
import sys
import xlrd

usage_msg = """\

USAGE: python apc_analyse.py [-t template file path] [-f sample bin file path]

  -h	print this help message and exit

Examples:

python apc_analyse.py -f APC_IN1_170_20141106112023.bin

default template file path is 'apc_template.xls'

python apc_analyse.py -t apc_template.xls -f APC_IN1_170_20141106112023.bin

"""

separator = "---"

def prepare_output(template_filepath, output_path):
try:
		workbook = xlrd.open_workbook(in_path)
		sheet_number = range(workbook.nsheets)
		sheet = workbook.sheet_by_index(0)

		for i in range(1, sheet.nrows):
			key = ""
			empty = True

			for j in range(col_start - 1, col_end):
				cell = str(sheet.cell_value(i, j)).strip()
				if cell != "":
					empty = False

				key += cell + separator

			pos = key.rfind(separator)
			key = key[:pos]

			if empty == False:
				if key in index_table:
					index_table[key].append([fname, i + 1])
				else:
					index_table[key] = [[fname, i + 1]]

				print ".",

		print ""
	except:
	    print "no sheet in %s named Sheet1" % fname

def do(template_filepath, in_path, output_path):
	print "doing", in_path,

	# create output xls and copy figure sheet from template
	output_workbook = prepare_output(template_filepath, output_path)

	try:
		fd = open(in_path, 'rb')

		while True:
			sample = fd.read(52)
			if sample:
				for i in range(0, 8):
					temp = sample[i+1] << 8 + sample[i]
				print ".",
			else:
				break		

		print ""
	finally:
		fd.close

def main(argv):
	try:
		opts, args = getopt.getopt(argv[1:], "t:f:h:")

		template_filepath = "apc_template.xls"
		in_path = ""
		output_path = ""

		for opt, arg in opts:
			if opt == '-t':
				template_filepath = arg
			elif opt == '-f':
				in_path = arg
			elif opt =='-h':
				print(usage_msg)
				sys.exit(1)

		if template_filepath == "" or in_path == "":
			print(usage_msg)
			sys.exit(1)
		else:
			output_path = os.path.basename(in_path)
			do(template_filepath, in_path, output_path)
	except getopt.error, msg:
		sys.stderr.write("error: %s\n" % str(msg))
		sys.stderr.write(usage_msg)
		sys.exit(2)

if __name__ == '__main__':
    main(sys.argv)