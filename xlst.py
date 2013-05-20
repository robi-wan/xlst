#!/usr/bin/env python
# vim: set fileencoding=utf-8 :

import argparse
import os
import xlrd

class MPS3(object):

    sheet_index = 3
    # rows and columns are zero based
    column = 0
    start_row = 9
    values = []
    output_name = 'mps3.ini'

    def __init__(self, book, path):
        self.book = book
        self.path = path
        self._collect_data()
        self._write()

    def _collect_data(self):
        col = self.column
        sheet = self.book.sheet_by_index(self.sheet_index)
        #values = []
        for row in xrange(self.start_row, sheet.nrows):
            text = sheet.cell(row,col).value
            if len(text) > 0:
                self.values.append(text)

        #print len(self.values)
        #print ','.join(self.values)

    def _write(self):
        f = os.path.join(self.path, self.output_name)
        with open(f, 'wb') as mps_ini:
            for value in self.values:
                #if value.startswith('['):
                    #mps_ini.write('\n')
                mps_ini.write(value.encode('cp1252'))        
                mps_ini.write('\n')

class SetupExtractor(object):

    def __init__(self, workbook):
        self.path = workbook
        self.book = xlrd.open_workbook(self.path)
        self._main_config()

    def _main_config(self):
        MPS3(self.book, os.path.dirname(self.path))


def main():
    parser = argparse.ArgumentParser()
    #TODO optional argument: output path (default: path of workbook)
    parser.add_argument("workbook", help='path to workbook with setup data to extract')
    args = parser.parse_args()
    workbook = os.path.abspath(args.workbook)
    extractor = SetupExtractor(workbook)

if __name__ == '__main__':
    main()
