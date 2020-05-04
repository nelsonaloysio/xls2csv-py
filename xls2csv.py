#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
Converts file to and from CSV/XLS/XLX.

usage: xls2csv.py [-h] [-o OUTPUT] [-q {0,1,2,3}] [-e ENCODING]
                  input

positional arguments:
  input                 input file name

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        output file or folder name
  -q {0,1,2,3}, --quoting {0,1,2,3}
                        text quoting {0: 'minimal', 1: 'all',
                        2: 'non-numeric', 3: 'none'}
  -e ENCODING, --encoding ENCODING
                        file encoding (default: utf-8)
'''

from argparse import ArgumentParser
from csv import reader, writer, QUOTE_MINIMAL
from os import listdir, mkdir
from os.path import basename, splitext
from os.path import exists, isdir, isfile

from xlrd import open_workbook
from xlwt import Workbook

ENCODING = 'utf-8'

QUOTING = {0: 'minimal',
           1: 'all',
           2: 'non-numeric',
           3: 'none'}

def convert_file(input_name, output_name=None,
    quoting=QUOTE_MINIMAL, encoding=ENCODING):
    '''
    Converts file based on extension format.
    '''
    name, ext = splitext(input_name)

    if ext in ('.xls', '.xlsx'):
        # convert to CSV file
        xls2csv(input_name, output_name, quoting, encoding)
    else: # convert to Excel file
        csv2xls(input_name, output_name, quoting, encoding)

def csv2xls(input_name, output_file=None,
    quoting=QUOTE_MINIMAL, encoding=ENCODING):
    '''
    Converts CSV format file to Excel (XLS).
    '''
    if isinstance(input_name, str):
        input_files = []
        if isdir(input_name):
            for f in sorted(listdir(input_name)):
                input_files.append(input_name+'/'+f)
        elif isfile(input_name):
            input_files = [input_name]
        else: # error
            print('Error: invalid input; neither a folder nor a file.')
            raise SystemExit
    else: # as a list of files
        input_files = input_name

    if not output_file:
        name, ext = splitext(input_name if isinstance(input_name,str) else input_files[0])
        output_file = basename(name)+'.xlsx'

    if exists(output_file):
        print("Error: file '%s' already exists." % output_file)
        raise SystemExit

    book = Workbook()

    for s,n in enumerate(input_files):
        delimiter = get_file_delimiter(n, encoding)
        sheet = book.add_sheet(basename(n))

        with open(n, 'r', encoding=encoding) as f:
            file_reader = reader(f, delimiter=delimiter, quoting=quoting)
            header = next(file_reader)
            row = sheet.row(0)
            for i,v in enumerate(header):
                row.write(i, v)
            for r,line in enumerate(file_reader):
                row = sheet.row(r+1)
                for i,v in enumerate(line):
                    row.write(i, v)

    book.save(output_file)

def xls2csv(input_file, output_folder='.',
    delimiter=None, quoting=QUOTE_MINIMAL, encoding=ENCODING):
    '''
    Converts Excel (xls/xlsx) format files to CSV.
    '''
    name, ext = splitext(input_file)
    name = basename(name)

    input_xls = open_workbook(input_file)
    sheets = input_xls.sheet_names()

    if not output_folder:
        output_folder = '.'

    if not exists(output_folder):
        mkdir(output_folder)
    elif output_folder != '.':
        print("Error: folder '%s' already exists." % output_folder)
        raise SystemExit

    if not delimiter:
        delimiter = ','

    lst = []
    for i in sheets:
        s = input_xls.sheet_by_name(i)
        o = output_folder+'/'+name+'_'+str(i)+'.csv'
        with open(o, 'w', encoding=encoding) as f:
            w = writer(f, delimiter=delimiter, quoting=quoting)
            for line in range(s.nrows):
                row = s.row_values(line)
                w.writerow(row)
        lst.append(o)

    return lst

def get_file_delimiter(input_name, encoding=ENCODING):
    '''
    Returns character delimiter from file.
    '''
    with open(input_name, 'rt', encoding=encoding) as input_file:
        file_reader = reader(input_file)
        header = str(next(file_reader))

    for i in ['|', '\\t', ';', ',']:
        if i in header: # \\t != \t
            return i.replace('\\t', '\t')

    return '\n'

if __name__ == "__main__":

    parser = ArgumentParser()

    parser.add_argument('input', action='store', help='input file name')
    parser.add_argument('-o', '--output', action='store', help='output file or folder name')
    parser.add_argument('-q', '--quoting', action='store', type=int, choices=QUOTING.keys(), default=0, help='text quoting %s' % QUOTING)
    parser.add_argument('-e', '--encoding', action='store', help='file encoding (default: %s)' % ENCODING)

    args = parser.parse_args()

    convert_file(args.input,
                 args.output,
                 args.quoting,
                 args.encoding)
