#!/bin/env python
#***************************************************************************************************
# usage: python csv2Chart.py (required inputfilenamd) (optional outputfilename)
# Change Log
#
# Date     Person       Description
# -------- ------------ -----------------------------------------------------------------------------
# 03/6/14  HGA   v1.0 This script takes in csv files and converts it to xlsx
#***************************************************************************************************

import xlsxwriter
import os
import os.path
import sys
import re
import gzip
import zipfile
import csv

#check to see if var is passed
if len(sys.argv) == 1:
    print "Please input tab - file name to covert to csv"
    print "usage: python csv2Chart.py (required inputfilenamd) (optional outputfilename)"
    sys.exit(0)

# input files
inputfile = sys.argv[1]

#check the extension of the input file
def check_ext(input):
    """Returns the extension of the file"""
    extension = os.path.splitext(input)[1]
    return extension

#check to see if the output file name is defined - if not, define
if len(sys.argv) == 2:
    if (check_ext(inputfile) == '.gz'):
        outputfile = inputfile.replace('.gz', '.xlsx')
    elif (check_ext(inputfile) == '.zip'):
        outputfile = inputfile.replace('.zip', '.xlsx')
    elif (check_ext(inputfile) == '.csv'):
        outputfile = inputfile.replace('.csv', '.xlsx')
    else:
        outputfile = inputfile + '.xlsx'
#    output = re.sub('gz','xlsx', inputfile)
else:                                       # otherwise, use files specified
    outputfile = sys.argv[2]

#open the file to read
def open_file(input):
    """Returns opened input file"""
    #checks to see if the file extension is gzip/zip
    extension = check_ext(input)
    #print extension
    if extension == '.gz':
        datafile = gzip.open(input, 'rb')
    elif extension == '.zip':
        datafile = zipfile.open(input, 'rb')
    else:
        datafile = open(input, 'rb')
    return datafile
# check to see if the record is float or integer
# chart need these number to be int or float instead of str
def is_number(s):
    """Returns true if float or integer"""
    try:
        float(s) or int(s)
        return True
    except ValueError:
        return False
def createXls(inputfile, outputfile):
    """Returns a file with chart """
    #Creating Excel files with Python and XlsxWriter
    workbook = xlsxwriter.Workbook(outputfile)
    worksheet = workbook.add_worksheet()
    style =  workbook.add_format()
    style.set_num_format('mm/dd/yyyy')
    row_list = []
    #open the inputfile to read
#    with open(input, 'r' ) as csvfile:
#        ifile = csv.reader(csvfile, delimiter=',', quotechar='"')
    ifile = open_file(inputfile)
    rows = csv.reader(ifile, delimiter=',', quotechar='"')
    #Read through the file and append the row to row_list
    for row in rows:
        row_list.append(row)

    # read through the rows and columns and write the records
    i = 0
    column_list = row_list
    for column in column_list:
        for item in range(len(column)):
            value = column[item].strip()
            if is_number(value):
                worksheet.write(i,item, float(value))
            else:
                worksheet.write(i, item, value)
        i += 1

    # Save and close workbook
    workbook.close()

# run function creatChart with two var input and output
createXls(inputfile, outputfile)
