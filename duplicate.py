#!/usr/bin/python

import os
import getopt
import sys
import xlrd

usage_msg = """\

USAGE: duplicate [-f excel file list] [-d excel directory] [-c column range] [-p only print row when duplicate count]

  -h	print this help message and exit

Examples:

duplicate -f "a.xls b.xls c.xls" -c 1,4 -p 1
        find duplicate rows from a.xls & b.xls & c.xls, and order by the column of 1~4 and only print row when duplicate count > 1

duplicate -d /opt/xls -c 1,4 -p 2
        find duplicate rows from "/opt/xls", and order by the column of 1~4 and only print row when duplicate count > 2
"""

separator = "---"

def duplicate_stat(index_table, duplicate_count):
	id = 1

	for k, v in sorted(index_table.iteritems()):
		if len(v) >= duplicate_count:
			print str(id)+". " + k
			id += 1

			print '  total ' + str(len(v)) + " lines:"
			for obj in v:
				print "  " + obj[0] + "'s " + str(obj[1]) + " line"

def duplicate_directory(directory, column_begin, column_end, index_table):
	print("sorry, temporary unrealized")
	return True

def duplicate_file_list(file_list, column_begin, column_end, index_table):
	files = file_list.split(" ")

	for fname in files[0:]:
		print "reading", fname,

		workbook = xlrd.open_workbook(fname)
		sheet_number = range(workbook.nsheets)
		try:
		    sheet = workbook.sheet_by_index(0)
		except:
		    print "no sheet in %s named Sheet1" % fname
		    return False

		for i in range(1, sheet.nrows):
			row_key = ""
			for j in range(column_begin - 1, column_end):
				row_key += str(sheet.cell_value(i, j)).strip() + separator

			if row_key != "":
				if row_key in index_table:
					index_table[row_key].append([fname, i + 1])
				else:
					index_table[row_key] = [[fname, i + 1]]

				print ".",

		print ""

	return True

def get_column(column):
	try:
		column_list = column.split(",")
		return int(column_list[0]), int(column_list[1])
	except:
		print(usage_msg)
		sys.exit(2)

def duplicate(file_list, directory, column, show_count):
	if file_list == "" and directory == "":
		print(usage_msg)
	else:
		column_begin, column_end = get_column(column)

		index_table = {}

		if file_list == "":
			duplicate_directory(directory, column_begin, column_end, index_table)
		else:
			duplicate_file_list(file_list, column_begin, column_end, index_table)

		duplicate_stat(index_table, show_count)

def main(argv):
	try:
		opts, args = getopt.getopt(argv[1:], "f:d:c:p:h:")

		file_list = ""
		directory = ""
		column = "1,4"
		duplicate_count = 0		

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
