#!/usr/bin/env python
# vim: set fileencoding=utf-8 :

import argparse
import os
import xlrd


class LanguageConfig(object):

    sheet_index = 1  # 'lng903'
    # rows and columns are zero based
    start_row = 9

    def __init__(self, book):
        self.book = book
        self.values = []
        self._collect_data()

    def _collect_data(self):
        col = 0  # 'A'
        sheet = self.book.sheet_by_index(self.sheet_index)
        for row in xrange(self.start_row, sheet.nrows):
            text = sheet.cell(row, col).value
            if len(text) > 0:
                self.values.append(text)

        col = 8  # 'I'
        for row in xrange(self.start_row, sheet.nrows):
            text = sheet.cell(row, col).value
            if len(text) > 0:
                self.values.append(text)
            else:
                # relevant data in this column ends with first blank cell
                break


class LanguageConfigGenerator(object):

    output_name = 'lng.ini'
    encoding = 'cp1252'

    def __init__(self, values, path):
        self.values = values
        self.path = path
        self._write()

    def _write(self):
        f = os.path.join(self.path, self.output_name)
        with open(f, 'wb') as mps_ini:
            for value in self.values:
                mps_ini.write("{}\n".format(value.encode(self.encoding)))


class TranslationExtractor(object):

    def __init__(self, workbook):
        self.path = workbook
        self.book = xlrd.open_workbook(self.path)
        print self.book.biff_version, self.book.codepage, self.book.encoding
        self.output_path = os.path.dirname(self.path)
        self._language_config()

    def _language_config(self):
        lng = LanguageConfig(self.book)
        LanguageConfigGenerator(lng.values, self.output_path)


def main():
    parser = argparse.ArgumentParser()
    #TODO optional argument: output path (default: path of workbook)
    parser.add_argument("workbook", help='path to workbook with setup data to extract')
    args = parser.parse_args()
    workbook = os.path.abspath(args.workbook)
    TranslationExtractor(workbook)

if __name__ == '__main__':
    main()
