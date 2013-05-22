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

    def __init__(self, book):
        self.book = book
        self.values = []
        self._collect_data()

    def _collect_data(self):
        col = self.column
        sheet = self.book.sheet_by_index(self.sheet_index)
        for row in xrange(self.start_row, sheet.nrows):
            text = sheet.cell(row, col).value
            if len(text) > 0:
                self.values.append(text)


class HMI(MPS3):

    sheet_index = 4

    def __init__(self, book):
        if book.sheet_by_index(self.sheet_index):
            super(HMI, self).__init__(book)

class Generator(object):

    encoding = 'cp1252'


class MPS3Generator(Generator):

    output_name = 'mps3.ini'

    def __init__(self, values, path):
        self.values = values
        self.path = path
        self._write()

    def _write(self):
        f = os.path.join(self.path, self.output_name)
        with open(f, 'wb') as mps_ini:
            for value in self.values:
                #if value.startswith('['):
                    #mps_ini.write('\n')
                mps_ini.write(value.encode(self.encoding))
                mps_ini.write('\n')


class HMIGenerator(MPS3Generator):

    output_name = 'HMISetup.ini'


class OutputObject(object):

    comment = None
    ini_key = None

    def __init__(self, number=None, name=None):
        self.number = number
        self.name = name

    def __str__(self):
        return "{}{}={}".format(self.ini_key, self.number, self.name.encode(Generator.encoding))

    def __repr__(self):
        return "{}{}={}".format(self.ini_key, self.number, self.name.encode(Generator.encoding))


class Parameter(OutputObject):

    comment = u'//Parametertexte'
    ini_key = 'PARAM'

    def __init__(self, note=None, **kwargs):
        self.note = note
        super(Parameter, self).__init__(**kwargs)


class Category(OutputObject):

    comment = u'//Texte Tabelle/Registerkarte'
    ini_key = 'TAB'


class Header(OutputObject):

    comment = u'//Überschriften Spalten'
    ini_key = 'COL'


class Menu(OutputObject):

    comment = u'//MenüTexte'
    ini_key = 'MENU'


class SystemMessages(OutputObject):

    comment = u'//Systemtexte(Beschriftungen, Überschriften, usw.)'
    ini_key = 'SYSTEM'


class ErrorMessages(OutputObject):

    comment = u'//Fehlertexte'
    ini_key = 'ERROR'


class HMICategory(OutputObject):

    comment = u'//Texte Registerkarte HMI'
    ini_key = 'TABHMI'


class Translation(object):

    languages = (('deutsch', 1), ('english', 2))
    start_row = 9
    params = 1300
    param_name_col = 0
    param_number_col = 1

    def __init__(self, book):
        self.book = book
        self.values = {}
        self._collect_data()
        #print self.values

    def _collect_data(self):
        for lang, sheet in self.languages:
            sheet = self.book.sheet_by_index(sheet)

            self._parameters(sheet, lang)
            self._categories(sheet, lang)
            self._column_header(sheet, lang)
            self._menues(sheet, lang)
            self._system_messages(sheet, lang)
            self._error_messages(sheet, lang)
            self._hmi_categories(sheet, lang)

    def _parameters(self, sheet, lang):
        for row in xrange(self.start_row, self.start_row + self.params):
            name = sheet.cell(row, self.param_name_col).value

            # extract param number as integer (Excel just knows floats)
            cell = sheet.cell(row, self.param_number_col)
            number = cell.value
            if cell.ctype in (2, 3) and int(number) == number:
                number = int(number)

            note = sheet.cell_note_map.get((row, self.param_name_col), None)

            self.values.setdefault((lang, 'params'), []).append(Parameter(number=number, name=name, note=note))

    def _categories(self, sheet, lang):
        start_row = 1349
        for row in xrange(start_row, start_row + 20):
            cat_name = sheet.cell(row, self.param_name_col).value
            if len(cat_name) > 0:
                index = row - start_row
                self.values.setdefault((lang, 'categories'), []).append(Category(name=cat_name, number=index))
            else:
                break

    def _column_header(self, sheet, lang):
        start_row = 1369
        for row in xrange(start_row, start_row + 10):
            name = sheet.cell(row, self.param_name_col).value
            if len(name) > 0:
                index = row - start_row
                self.values.setdefault((lang, 'header'), []).append(Header(name=name, number=index))
            else:
                break

    def _menues(self, sheet, lang):
        start_row = 1379
        for row in xrange(start_row, start_row + 30):
            name = sheet.cell(row, self.param_name_col).value
            if len(name) > 0:
                index = row - start_row
                self.values.setdefault((lang, 'menues'), []).append(Menu(name=name, number=index))
            else:
                break

    def _system_messages(self, sheet, lang):
        start_row = 1409
        for row in xrange(start_row, start_row + 50):
            name = sheet.cell(row, self.param_name_col).value
            if len(name) > 0:
                index = row - start_row
                self.values.setdefault((lang, 'system_messages'), []).append(SystemMessages(name=name, number=index))
            else:
                break

    def _error_messages(self, sheet, lang):
        start_row = 1459
        for row in xrange(start_row, self.__rows(sheet, start_row + 20)):
            name = sheet.cell(row, self.param_name_col).value
            if len(name) > 0:
                index = row - start_row
                self.values.setdefault((lang, 'error_messages'), []).append(ErrorMessages(name=name, number=index))
            else:
                break

    def _hmi_categories(self, sheet, lang):
        start_row = 1319
        for row in xrange(start_row, start_row + 30):
            name = sheet.cell(row, self.param_name_col).value
            if len(name) > 0:
                index = row - start_row
                self.values.setdefault((lang, 'hmi_categories'), []).append(HMICategory(name=name, number=index))
            else:
                break

    def __rows(self, sheet, rows):
        return min(rows, sheet.nrows)


def description_ranges():
    # xrange: end value is exclusive
    return xrange(200), xrange(200, 600), xrange(600, 1300)


class TranslationGenerator(Generator):

    suffix = '.lng'

    def __init__(self, lang, values, path):
        self.languages = lang
        self.values = values
        self.path = path
        self._write()

    def _write(self):
        for lang in self.languages:
            f = os.path.join(self.path, "{}{}".format(lang, self.suffix))
            with open(f, 'wb') as lang_file:
                lang_file.write("[{}]".format(lang))
                lang_file.write('\n')
                data = ('params', 'categories', 'header', 'menues',
                        'system_messages', 'error_messages', 'hmi_categories')
                for values in [self.values.get((lang, d)) for d in data]:
                    if values:  # values for 'hmi_categories' may be None
                        lang_file.write(values[0].comment.encode(self.encoding))
                        lang_file.write('\n')
                        for value in values:
                            lang_file.write(str(value))
                            lang_file.write('\n')

                        lang_file.write('\n')

                self._write_notes(lang)

    def _write_notes(self, lang):
        for i in range(len(description_ranges())):
            ran = description_ranges()[i]
            f = os.path.join(self.path, "{}{}{}".format(lang, i+1, self.suffix))
            with open(f, 'wb') as desc_file:
                desc_file.write("[{}]".format(lang.upper()))
                desc_file.write('\n')
                for n in ran:
                    desc_file.write("HILFEPARAM{}={}".format(n, self.__note(lang, n)))
                    desc_file.write('\n')

    def __delimiter(self):
        return u'§§'.encode(self.encoding)

    def __note(self, lang, number):
        params = self.values.get((lang, 'params'))
        for p in params:
            if p.number == number and p.note:
                note = p.note.text.encode(self.encoding)
                note = self.__delimiter().join(note.splitlines())
                return note
        return None


class SetupExtractor(object):

    def __init__(self, workbook):
        self.path = workbook
        self.book = xlrd.open_workbook(self.path)
        self.output_path = os.path.dirname(self.path)
        self._main_config()
        self._hmi_config()
        self._translation()

    def _main_config(self):
        mps3 = MPS3(self.book)
        MPS3Generator(mps3.values, self.output_path)

    def _hmi_config(self):
        hmi = HMI(self.book)
        if hmi.values:
            HMIGenerator(hmi.values, self.output_path)

    def _translation(self):
        t = Translation(self.book)
        TranslationGenerator([lang for lang, index in t.languages], t.values, self.output_path)


def main():
    parser = argparse.ArgumentParser()
    #TODO optional argument: output path (default: path of workbook)
    parser.add_argument("workbook", help='path to workbook with setup data to extract')
    args = parser.parse_args()
    workbook = os.path.abspath(args.workbook)
    SetupExtractor(workbook)

if __name__ == '__main__':
    main()
