#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
Converts file to and from CSV/XLS/XLX.

usage: xls2csv [-h] [-o OUTPUT_NAME] [-d DELIMITER] [-e ENCODING]
               [-E ESCAPECHAR] [-q {0,1,2,3}] [-Q QUOTECHAR]
               input_name [input_name ...]

positional arguments:
  input_name            input file name (one or multiple)

options:
  -h, --help            show this help message and exit
  -o OUTPUT_NAME, --output-name OUTPUT_NAME
                        output file or folder name
  -d DELIMITER, --delimiter DELIMITER
                        column field delimiter (default: comma)
  -e ENCODING, --encoding ENCODING
                        file encoding (default: 'utf-8')
  -E ESCAPECHAR, --escapechar ESCAPECHAR
                        escape character (default: '\')
  -q {0,1,2,3}, --quoting {0,1,2,3}
                        text quoting {0: 'minimal', 1: 'all', 2: 'non-
                        numeric', 3: 'none'}
  -Q QUOTECHAR, --quotechar QUOTECHAR
                        quote character (default: '"')
'''

from argparse import ArgumentParser
from csv import reader, writer
from os import listdir, mkdir
from os.path import basename, splitext
from os.path import exists, isdir, isfile
from sys import stderr

from xlrd import open_workbook
from xlwt import Workbook

ENCODING = 'utf-8'
ESCAPECHAR = '\\'
QUOTECHAR = '"'

QUOTING = {0: 'minimal',
           1: 'all',
           2: 'non-numeric',
           3: 'none'}

def convert_file(files, **kwargs):
    '''
    Converts file based on extension format.
    '''
    for input_name in ([files] if type(files) == str else files):
        name, ext = splitext(input_name)
        (xls2csv if ext in ('.xls', '.xlsx') else csv2xls)(input_name, **kwargs)

def csv2xls(
    input_name,
    output_name=None,
    delimiter=',',
    encoding=ENCODING,
    escapechar=ESCAPECHAR,
    quotechar=QUOTECHAR,
    quoting=0,
) -> str:
    '''
    Converts CSV format file to Excel (XLS).
    '''
    book = Workbook()

    if quoting == 3:
        quotechar = ''

    if isinstance(input_name, str):
        input_files = []
        if isdir(input_name):
            for f in sorted(listdir(input_name)):
                input_files.append(f'{input_name}/{f}')
        elif isfile(input_name):
            input_files = [input_name]
        else: # error
            print("Error: '%s' is neither a folder nor a file.", file=stderr)
            raise SystemExit
    else: # as a list of files
        input_files = input_name

    if not output_name:
        name, ext = splitext(input_name if isinstance(input_name, str) else input_files[0])
        output_name = basename(name)+'.xlsx'

    if exists(output_name):
        print("Error: file '%s' already exists." % output_name, file=stderr)
        raise SystemExit

    for s, n in enumerate(input_files):
        sheet = book.add_sheet(basename(n))

        if not delimiter:
            delimiter = get_file_delimiter(n, encoding)

        with open(n, 'r', encoding=encoding) as f:
            file_reader = reader(f, delimiter=delimiter, escapechar=escapechar, quotechar=quotechar, quoting=quoting)
            header = next(file_reader)
            row = sheet.row(0)
            for i,v in enumerate(header):
                row.write(i, v)
            for r,line in enumerate(file_reader):
                row = sheet.row(r+1)
                for i,v in enumerate(line):
                    row.write(i, v)

    book.save(output_name)
    return output_name

def xls2csv(
    input_file,
    output_name='.',
    delimiter=',',
    encoding=ENCODING,
    escapechar=ESCAPECHAR,
    quotechar=QUOTECHAR,
    quoting=0,
) -> list:
    '''
    Converts Excel (xls/xlsx) format files to CSV.
    '''
    if quoting == 3:
        quotechar = ''

    name, ext = splitext(input_file)
    name = basename(name)

    input_xls = open_workbook(input_file)
    sheets = input_xls.sheet_names()

    if not output_name:
        output_name = '.'

    if not exists(output_name):
        mkdir(output_name)
    elif output_name != '.':
        print("Error: folder '%s' already exists." % output_name, file=stderr)
        raise SystemExit

    lst = []
    for i in sheets:
        s = input_xls.sheet_by_name(i)
        o = f'{output_name}/{name}_{str(i)}.csv'
        with open(o, 'w', encoding=encoding) as f:
            file_writer = writer(f, delimiter=delimiter, escapechar=escapechar, quoting=quoting, quotechar=quotechar)
            for line in range(s.nrows):
                row = s.row_values(line)
                file_writer.writerow(row)
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

    parser.add_argument('input_name', action='store', help='input file name (one or multiple)', nargs='+')
    parser.add_argument('-o', '--output-name', action='store', help='output file or folder name')
    parser.add_argument('-d', '--delimiter', action='store', default=',', help='column field delimiter (default: comma)')
    parser.add_argument('-e', '--encoding', action='store', default=ENCODING, help='file encoding (default: \'%s\')' % ENCODING)
    parser.add_argument('-E', '--escapechar', action='store', default=ESCAPECHAR, help='escape character (default: \'%s\')' % ESCAPECHAR)
    parser.add_argument('-q', '--quoting', action='store', type=int, choices=QUOTING.keys(), default=0, help='text quoting %s' % QUOTING)
    parser.add_argument('-Q', '--quotechar', action='store', default=QUOTECHAR, help='quote character (default: \'%s\')' % QUOTECHAR)

    args = parser.parse_args()

    convert_file(
        args.input_name,
        delimiter=args.delimiter,
        encoding=args.encoding,
        escapechar=args.escapechar,
        output_name=args.output_name,
        quotechar=args.quotechar,
        quoting=args.quoting,
    )
