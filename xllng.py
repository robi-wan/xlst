#!/usr/bin/env python
# vim: set fileencoding=utf-8 :

import argparse
import os
import codecs
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
            if text:
                self.values.append(text)

        col = 8  # 'I'
        for row in xrange(self.start_row, sheet.nrows):
            text = sheet.cell(row, col).value
            if text:
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
        with codecs.open(f, mode='w', encoding=self.encoding) as outfile:
            for value in self.values:
                outfile.write(u"{}\n".format(value))


class Translation(object):

    languages = ['de', 'en', 'fr', 'es', 'it', 'nl', 'no', 'ja', 'pt', 'fi', 'hu', 'sk', 'cs', 'sv', 'pl', 'ro', 'da',
                 'sl', 'tr', 'et', 'hr', 'ru', 'el', 'lt', 'bg', 'zh']
    start_row = 9
    section_column = 1
    key_column = 2

    def __init__(self, book, path):
        self.book = book
        self.path = path

        # sections and keys are defined reliable only for german - so use these data for all languages
        self.section_sheet = self.book.sheet_by_index(3)
        self.key_sheet = self.section_sheet

        self._current_section = None
        self._current_key_prefix = None
        self._current_section_row = None
        self.values = {}

        self._collect_data()

    def __filename(self, lang):
        return 'touch{0:02d}.ini'.format(self.__lang_index(lang))

    def __sheet_index(self, lang):
        return 3 + self.__lang_index(lang)

    def __io_column(self, lang):
        return 1 + self.__lang_index(lang)

    def __io_messages_column(self, lang):
        return 3 + self.__lang_index(lang)

    def __lang_index(self, lang):
        return self.languages.index(lang)

    def _collect_data(self):
        for lang in self.languages:
            sheet = self.book.sheet_by_index(self.__sheet_index(lang))
            # reset current section for every new language
            self._current_section = None

            with codecs.open(os.path.join(self.path, self.__filename(lang)), mode='w', encoding='utf-16') as outfile:
                self._key_in_section = False
                for row in xrange(self.start_row, sheet.nrows):
                    if row not in xrange(3899, 4499):  # skip these rows - we don't need them
                        self._extract_data(sheet, row, outfile)

                self._io_names(lang, outfile)
                self._io_messages(lang, outfile)

    def _extract_data(self, sheet, row, outfile):
        sec = self.section_sheet.cell(row, self.section_column).value
        if sec and sec.startswith('['):
            # write empty line after finished section
            if self._current_section:
                outfile.write(u'\n')

            self._current_section = sec
            self._current_section_row = row
            self._key_in_section = True
            # just when a new section starts a new prefix is given
            key_prefix = self.key_sheet.cell(row, self.key_column).value
            if key_prefix:
                self._current_key_prefix = key_prefix
            # write new section
            outfile.write(u"{}\n".format(sec))

        text = sheet.cell(row, 0).value
        if text and self._key_in_section:
            index = row - self._current_section_row
            outfile.write(u"{}{}={}\n".format(self._current_key_prefix, index, text))
        else:
            # empty text - current "block" with data ends
            # wait with text until new section starts
            self._key_in_section = False

    def _io_names(self, lang, outfile):
        sheet = self.book.sheet_by_index(2)  # 'Seitendefinitionen'
        # extract start and end row from excel file - beware of start index (1 vs 0)!
        start_row = cell_content_as_integer(sheet.cell(9, 0))
        end_row = cell_content_as_integer(sheet.cell(10, 0))

        for row in xrange(start_row - 1, end_row):
            text = sheet.cell(row, self.__io_column(lang)).value
            outfile.write(u"{}\n".format(text))

    def _io_messages(self, lang, outfile):
        sheet = self.book.sheet_by_index(0)  # 'EATexte'

        outfile.write(u"[IO_TEXTE]\n")
        start_row = 39

        for row in xrange(start_row, sheet.nrows):
            text = sheet.cell(row, self.__io_messages_column(lang)).value
            index = 1 + row - start_row
            outfile.write(u"IO_{}={}\n".format(index, text))


def cell_content_as_integer(cell):
    # extract param number as integer (Excel just knows floats)
    number = cell.value
    if cell.ctype in (2, 3) and int(number) == number:
        number = int(number)

    return number


class TranslationExtractor(object):

    def __init__(self, workbook):
        self.path = workbook
        self.book = xlrd.open_workbook(self.path)
        self.output_path = os.path.dirname(self.path)
        self._language_config()
        self._translation()

    def _language_config(self):
        lng = LanguageConfig(self.book)
        LanguageConfigGenerator(lng.values, self.output_path)

    def _translation(self):
        Translation(self.book, self.output_path)


def main():
    parser = argparse.ArgumentParser()
    #TODO optional argument: output path (default: path of workbook)
    parser.add_argument("workbook", help='path to workbook with setup data to extract')
    args = parser.parse_args()
    workbook = os.path.abspath(args.workbook)
    TranslationExtractor(workbook)

if __name__ == '__main__':
    main()
