#!/usr/bin/env python

import math
import struct
import shutil
import os
import getopt
import sys
import xlrd

usage_msg = """\

USAGE: duplicate [-f excel file list] [-d excel directory] [-c column range]

  -h	print this help message and exit

Examples:

duplicate -f "a.xls b.xls c.xls" -c 1,4
        compare a.xls & b.xls & c.xls the column of 1~4

duplicate -d /opt/xls -c 1,4
        compare all xls files of "/opt/xls" the column of 1~4
"""

def duplicate_directory(directory, column_begin, column_end):
	return True

def duplicate_file_list(file_list, column_begin, column_end):
	files = file_list.split(" ")

	index_table = []

	for fname in files[0:]:
		print(fname)

		workbook = xlrd.open_workbook(fname)
		sheet_number = range(workbook.nsheets)
		try:
		    sheet = workbook.sheet_by_index(0)
		except:
		    print "no sheet in %s named Sheet1" % fname
		    return False

		for i in range(1, sheet.nrows):
			cel_list = []
			for j in range(column_begin - 1, column_end):
				cel_list.append(sheet.cell_value(i, j))

			print(cel_list)
			# index_table[cel_list] = {fname, i}
			# print(index_table[cel_list])

	return True

def get_column(column):
	try:
		column_list = column.split(",")
		return int(column_list[0]), int(column_list[1])
	except:
		print(usage_msg)
		sys.exit(2)

def valid_parameter(file_list, directory, column):
	if file_list == "" and directory == "":
		return False
	elif column == "":
		return False

	return True

def duplicate(file_list, directory, column):
	if file_list == "" and directory == "":
		print(usage_msg)
	else:
		column_begin, column_end = get_column(column)

		if file_list == "":
			duplicate_directory(directory, column_begin, column_end)
		else:
			duplicate_file_list(file_list, column_begin, column_end)

def main(argv):
	try:
		opts, args = getopt.getopt(argv[1:], "f:d:c:h:")

		file_list = ""
		directory = ""
		column = "1,4"

		for opt, arg in opts:
			if opt == '-f':
				file_list = arg
			elif opt == '-d':
				directory = arg
			elif opt == '-c':
				column = arg
			elif opt =='-h':
				print(usage_msg)
				sys.exit(2)

		duplicate(file_list, directory, column)
	except getopt.error, msg:
		sys.stderr.write("Error: %s\n" % str(msg))
		sys.stderr.write(usage_msg)
		sys.exit(2)

if __name__ == '__main__':
	main(sys.argv)
