import argparse
import mimetypes
import os
import openpyxl
import sys
import xlrd

implementation_names = [
    'XlsSpreadsheet',
    'XlsxSpreadsheet',
]

def open_spreadsheet(filename):
    mimeguess = mimetypes.guess_type(filename)[0]
    for impl_name in implementation_names:
        impl = getattr(sys.modules[__name__], impl_name)
        if impl.mimetype == mimeguess:
            return impl(filename)
    raise Exception("No implementation available for %s files" % mimeguess)

class Spreadsheet():
    mimetype = None

    def __init__(self, filename):
        self.filename = filename
        self.worksheets = []
        self.sheet_map = {}
        self._current_sheet = 0

    def __iter__(self):
        return self

    def __getitem__(self, key):
        if isinstance(key, int):
            return self.worksheets[key]
        return self.sheet_map(key)

    def __next__(self):
        i = self._current_sheet

        if len(self.worksheets) <= i:
            self._current_sheet = 0
            raise StopIteration

        self._current_sheet = i + 1
        return self.worksheets[i]

class Worksheet():
    def __getitem__(self, key):
        raise RuntimeError("Not implemented")

    def __iter__(self):
        return self

    def __next__(self):
        raise RuntimeError("Not implemented")

    def get_column(self, idx):
        raise RuntimeError("Not implemented")

    def get_column_count(self):
        raise RuntimeError("Not implemented")

    def get_row(self, idx):
        raise RuntimeError("Not implemented")

    def get_row_count(self):
        raise RuntimeError("Not implemented")


class XlsSpreadsheet(Spreadsheet):
    mimetype = 'application/vnd.ms-excel'

    def __init__(self, *args):
        Spreadsheet.__init__(self, *args)
        self.workbook = xlrd.open_workbook(self.filename)
        for ws in self.workbook.sheets():
            worksheet = XlsWorksheet(ws)
            self.worksheets.append(worksheet)
            self.sheet_map[ws.name] = worksheet

class XlsWorksheet(Worksheet):
    def __init__(self, xlrd_worksheet):
        self.worksheet = xlrd_worksheet

    def get_column(self, idx):
        return self.worksheet.col(idx)

    def get_column_count(self):
        return self.worksheet.ncols

    def get_row_count(self):
        return self.worksheet.nrows

    def get_row(self, idx):
        # We're using a 1 based index because that's what Excel and other
        # spreadsheets use, but xlrd uses 0 based indexing, thus the -1 below
        if idx == 0:
            raise RuntimeError('Spreadsheet rows start at 1')
        xlr_row = self.worksheet.row(idx - 1)
        row = []
        for cell in xlr_row:
            row.append(cell.value)
        return row

class XlsxSpreadsheet(Spreadsheet):
    mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    def __init__(self, *args):
        Spreadsheet.__init__(self, *args)
        self.workbook = openpyxl.reader.excel.load_workbook(self.filename)
        for ws in self.workbook.worksheets:
            worksheet = XlsxWorksheet(ws)
            self.worksheets.append(worksheet)
            self.sheet_map[ws.title] = worksheet

class XlsxWorksheet(Worksheet):
    def __init__(self, openpyxl_worksheet):
        self.worksheet = openpyxl_worksheet

    def get_column_count(self):
        return self.worksheet.max_column

    def get_row_count(self):
        return self.worksheet.max_row

    def get_row(self, idx):
        if idx == 0:
            raise RuntimeError('Spreadsheet rows start at 1')
        rangeid = 'A%d:%s%d' % (idx,
                                openpyxl.utils.get_column_letter(self.get_column_count() + 1),
                                idx)
        print(rangeid)
        row = []
        for row_impl in self.worksheet.iter_rows(rangeid):
            for cell in row_impl:
                row.append(cell.value)
        return row

def cli():
    parser = argparse.ArgumentParser(description="Convert spreadsheets to other formats")
    parser.add_argument('--source', dest='source',
                        required=True,
                        help='the spreadsheet we will convert')

    args = parser.parse_args()
    workbook = open_spreadsheet(args.source)

    import IPython
    IPython.embed()

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

    # xlsx_filename = "/home/tim/src/energy_information_administration/SpreadsheetTo/test_data/EIA923_Schedules_6_7_NU_SourceNDisposition_2013_Final.xlsx"
    # wb = XlsxSpreadsheet(xlsx_filename)
    # for sheet in wb:
    #     print(sheet)
    # print(wb[0])
    # # print(wb[3])
    # print(wb['File Layout'])
    # print(wb['blah'])
