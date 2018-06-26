# -*- coding: utf-8 -*-

import os
import sys

import xlrd

def file2list(filepath):
    ret = []
    with open(filepath, encoding='utf8', mode='r') as f:
        ret = [line.rstrip('\n') for line in f.readlines()]
    return ret

def list2file(filepath, ls):
    with open(filepath, encoding='utf8', mode='w') as f:
        f.writelines(['{:}\n'.format(line) for line in ls] )

def column2alphabet(column_number):
    div = column_number
    s = ''
    temp=0
    while div>0:
        module = (div-1)%26
        s = chr(65 + module)+s
        div = int((div-module)/26)
        return s

# row is Y, col is X
def convert_sheet_contents(bookobj, sheet_index, outfilename):
    sheet = bookobj.sheet_by_index(sheet_index)
    xsize = sheet.ncols
    ysize = sheet.nrows

    outlines = []
    outfilepath = os.path.join(selfdir, outfilename)

    for y in range(ysize):

        # Y is 0-origin, but excel rows is 1-origin.
        realy = y+1
        section_y = '# Line {:}'.format(realy)
        outlines.append(section_y)
        outlines.append('')

        for x in range(xsize):
            cell = sheet.cell(y, x)
            # float or Unicode
            v = cell.value
            if isinstance(v, float):
                content = str(v)
            else:
                content = v
            # X is 0-origin, but excel rows is 1-origin.
            alphabet = column2alphabet(x+1)
            section_x = '## {:} - {:}'.format(realy, alphabet)
            outlines.append(section_x)
            outlines.append(content)
            outlines.append('')

    list2file(outfilepath, outlines)

def parse_arguments():
    import argparse

    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        description='',
    )

    parser.add_argument('-t', '--target', default=None, required=True,
        help='A xlsx file you want to read.')

    parser.add_argument('-s', '--summary', default=False, action='store_true',
        help='If true, display the summary information.')

    parser.add_argument('targetindexes', nargs='*',
        help='Numbers of sheet you want to read. 0-origin.')

    parser.add_argument('--header', default='', type=str,
        help='If given, header of output files is overwritten to it.')
    parser.add_argument('--footer', default='', type=str,
        help='If given, the footer of output files is overwritten to it.')

    parsed_args = parser.parse_args()
    return parsed_args

# parsing arguments.
args = parse_arguments()
use_summary = args.summary
target_indexes = args.targetindexes
header_str = args.header
footer_str = args.footer
targetfilename = args.target

# fix xls filepath.
selfdir = os.path.abspath(os.path.dirname(__file__))
targetfilepath = os.path.join(selfdir, targetfilename)

# raed xls file.
book = xlrd.open_workbook(targetfilepath)
sheet_count = book.nsheets
sheet_names = book.sheet_names()

if use_summary:
    print('Total {:} sheets exist.'.format(sheet_count))
    for i,name in enumerate(sheet_names):
        print('{:<3}: {:}'.format(i, name))
    exit(0)

# @todo 桁数も引数指定できるようにする.
for i,idx_str in enumerate(target_indexes):
    idx = int(idx_str)
    print('[{:}/{:}]...'.format(i+1, len(target_indexes)))
    outfilename = '{:}{:03d}{:}.md'.format(header_str, idx, footer_str)
    convert_sheet_contents(book, idx, outfilename)

print('fin.')
