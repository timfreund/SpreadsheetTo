import mimetypes
import os
import openpyxl
import sys
import xlrd

class Spreadsheet():
    mimetype = None

    def __init__(self, filename):
        self.filename = filename

    def __iter__(self):
        return self

    # def __getitem__(self, key):
    #     if isinstance(key, int):
    #         return self.worksheets[key]
    #     for s in self.worksheets:
    #         if s.name == key:
    #             return s
    #     return KeyError(key)

    def next(self):
        raise StopIteration

class XlsSpreadsheet(Spreadsheet):
    mimetype = 'application/vnd.ms-excel'

    def __init__(self, *args):
        Spreadsheet.__init__(self, *args)
        self.workbook = xlrd.open_workbook(self.filename)
        self.current_sheet = 0

    def __getitem__(self, key):
        try:
            if isinstance(key, int):
                return self.workbook.sheet_by_index(key)
            return self.workbook.sheet_by_name(key)
        except (IndexError, xlrd.biffh.XLRDError):
            pass

        raise KeyError(key)

    def __next__(self):
        i = self.current_sheet

        if self.workbook.nsheets <= i:
            self.current_sheet = 0
            raise StopIteration

        self.current_sheet = i + 1
        return self.workbook.sheet_by_index(i)

class XlsxSpreadsheet(Spreadsheet):
    mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    def __init__(self, *args):
        Spreadsheet.__init__(self, *args)
        self.workbook = openpyxl.reader.excel.load_workbook(self.filename)
        self.current_sheet = 0

    def __getitem__(self, key):
        try:
            if isinstance(key, int):
                return self.workbook.worksheets[key]
            return self.workbook.get_sheet_by_name(key)
        except (IndexError, xlrd.biffh.XLRDError):
            pass

        raise KeyError(key)

    def __next__(self):
        i = self.current_sheet

        if len(self.workbook.worksheets) <= i:
            self.current_sheet = 0
            raise StopIteration

        self.current_sheet = i + 1
        return self.workbook.worksheets[i]

def cli():
    # xls_filename = "/home/tim/src/energy_information_administration/eia-923-mirror/data/utility/multisheet.xls"
    # s = XlsSpreadsheet(xls_filename)
    # for sheet in s:
    #     print (sheet)

    # wb = s
    # import ipdb; ipdb.set_trace()
    # print(s['Sheet2'])
    # print(s[0])
    # print(s[2])
    # # raise RuntimeError("Unimplemented")

    xlsx_filename = "/home/tim/src/energy_information_administration/SpreadsheetTo/test_data/EIA923_Schedules_6_7_NU_SourceNDisposition_2013_Final.xlsx"
    wb = XlsxSpreadsheet(xlsx_filename)
    for sheet in wb:
        print(sheet)
    print(wb[0])
    # print(wb[3])
    print(wb['File Layout'])
    print(wb['blah'])
