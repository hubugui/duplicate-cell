#!/usr/bin/env python

import os
import getopt
import sys
import xlrd

usage_msg = """\

USAGE: duplicate [-f excel file list] [-d excel directory] [-c column range] [-p print row when greater than duplicate count]

  -h	print this help message and exit

Examples:

duplicate -f "a.xls b.xls c.xls" -c 1,4 -p 1
        print duplicate rows from a.xls & b.xls & c.xls, and comparison of the 1~4 column and print row when duplicate count >= 1

duplicate -d /opt/xls -c 2,2 -p 2
        print duplicate rows from "/opt/xls", and comparison of the 2~2 column and only print row when duplicate count >= 2

"""

separator = "---"

def duplicate_stat(index_table, duplicate_count):
	id = 1

	for k, v in sorted(index_table.iteritems()):
		if len(v) >= duplicate_count:
			id_str = str(id)
			print id_str+"." + k
			id += 1

			prefix = ""
			for i in range(0, len(id_str) + 1):
				prefix += " "

			print prefix + str(len(v)) + " lines:"
			for obj in v:
				print prefix + obj[0] + ", line " + str(obj[1])

def duplicate_file(fname, col_start, col_end, index_table):
	print "reading", fname,

	try:
		workbook = xlrd.open_workbook(fname)
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

def duplicate_dir(directory, col_start, col_end, index_table):
	files = os.listdir(directory)

	for fname in files:
		if os.path.isfile(fname) == True and fname.endswith(".xls") == True:
			duplicate_file(fname, col_start, col_end, index_table)

def duplicate_file_list(file_list, col_start, col_end, index_table):
	files = file_list.split(" ")

	for fname in files[0:]:
		duplicate_file(fname, col_start, col_end, index_table)

def get_column(column):
	try:
		column_list = column.split(",")
		return int(column_list[0]), int(column_list[1])
	except:
		print(usage_msg)
		sys.exit(2)

def duplicate(file_list, directory, column, duplicate_count):
	if file_list == "" and directory == "":
		print(usage_msg)
	else:
		col_start, col_end = get_column(column)
		index_table = {}

		if file_list == "":
			duplicate_dir(directory, col_start, col_end, index_table)
		else:
			duplicate_file_list(file_list, col_start, col_end, index_table)

		duplicate_stat(index_table, duplicate_count)

def main(argv):
	try:
		opts, args = getopt.getopt(argv[1:], "f:d:c:p:h:")

		file_list = ""
		directory = ""
		column = "1,4"
		duplicate_count = 1		

		for opt, arg in opts:
			if opt == '-f':
				file_list = arg
			elif opt == '-d':
				directory = arg
			elif opt == '-c':
				column = arg
			elif opt == '-p':
				duplicate_count = int(arg)
			elif opt =='-h':
				print(usage_msg)
				sys.exit(2)

		duplicate(file_list, directory, column, duplicate_count)
	except getopt.error, msg:
		sys.stderr.write("Error: %s\n" % str(msg))
		sys.stderr.write(usage_msg)
		sys.exit(2)

if __name__ == '__main__':
    main(sys.argv)
