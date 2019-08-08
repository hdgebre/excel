#!/bin/env python
#***************************************************************************************************
# usage: python tab2csv.py (required inputfilenamd) (optional outputfilename)
# Change Log
#
# Date     Person       Description
# -------- ------------ -----------------------------------------------------------------------------
# 03/6/14  HGA   v1.0 This script takes in tab delimited files and converts it to csv
#***************************************************************************************************

import os
import os.path
import sys
import csv
import re
import gzip
import zipfile

#check to see if var is passed
if len(sys.argv) == 1:
    print "Please input tab - file name to covert to csv"
    print "usage: python tab2csv.py (required inputfilename) (optional outputfilename)"
    sys.exit(0)

#iput file name
input = sys.argv[1]
#two var input and output files
def check_ext(input):
    """Returns the extension of the file"""
    extension = os.path.splitext(input)[1]
    return extension

#check to see if the output file name is defined - if not, define
if len(sys.argv) == 2:
    if (check_ext(input) == '.gz'):
        output = input.replace('.gz', '.csv')
    elif (check_ext(input) == '.zip'):
        output = input.replace('.zip', '.csv')
    else:
        output = input + '.csv'
#    output = re.sub('gz','csv', input)
else:                                       # otherwise, use files specified
    output = sys.argv[2]

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

# looks for pattern to ignor
def ignore_record(line):
    """Returns True if string found in line"""
    try:
        if found == re.search('rows', line).group():
            return 1
        else:
            return 0
    except AttributeError:
        found = '' # apply your error handling

datafile = open_file(input)
print datafile
data = []
count = 0
#insert the row information to dataa list
for row in datafile:
    data.append(row.strip().split('\t'))
# remove the last line because it contails the row count only
# (NOTE: comment this if you need all rows to output
data = data[:-1]
#QUOTE OPTIONS
#quote all output
#out_csv = csv.writer(open(output, 'wb'), quoting=csv.QUOTE_ALL)
#quote on the special character
out_csv = csv.writer(open(output, 'wb'), delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
#quote non numeric
#out_csv = csv.writer(open(output, 'wb'), quoting=csv.QUOTE_NONNUMERIC)
#quote none
#out_csv = csv.writer(open(output, 'wb'), quoting=csv.QUOTE_NONE)
out_csv.writerows(data)

#print  data[:], ignore_record(lastrow)
