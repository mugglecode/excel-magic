import xlrd
import xlsxwriter
import datetime
from excel_magic.dataset import Dataset
from typing import List, Union, Dict

__all__ = ['Pointer', 'MagicSheet', 'ExcelDocument']

XL_CELL_EMPTY = 0
XL_CELL_TEXT = 1
XL_CELL_NUMBER = 2
XL_CELL_DATE = 3
XL_CELL_BOOLEAN = 4
XL_CELL_ERROR = 5
XL_CELL_BLANK = 6


class Pointer:
    def __init__(self, row: int, col: int):
        self.row = row
        self.col = col

    def next_row(self, current_col=False):
        if not current_col:
            self.col = 0
        self.row += 1

    def next_col(self):
        self.col += 1


class MagicSheet:
    """
    Do not initialize this class yourself
    """
    def __init__(self, name, sheet):
        self.name = name
        self.cols = 0
        self.rows = 0
        self.raw_sheet: xlsxwriter.workbook.Worksheet = sheet
        self.headers = []

    def _add_col(self):
        self.cols += 1

    def _add_row(self):
        self.rows += 1

    def append_row(self, content: Union[Dict, List[str]]):
        """
        append data to current sheet
        :param content: a dict{header: content} or a list [content]
        :return: None
        """
        if isinstance(content, dict):
            # is a dict
            for i in range(self.headers.__len__()):
                if self.headers[i] in content:
                    self.raw_sheet.write(self.rows, i, content[self.headers[i]])
            self._add_row()
        elif isinstance(content, list):
            for i in range(content.__len__()):
                self.raw_sheet.write(self.rows, i, content[i])
            self._add_row()
        else:
            raise TypeError(f'Expected dict or list, got {str(type(content))} instead')

    def __str__(self):
        return self.name


def create_document(path: str, template: Union[str, Dict]):
    """
    Create a new xlsx document
    :param path: path of the new document
    :param template: a file template or a header
    :return: an ExcelDocument object

    """
    return ExcelDocument(path, template)


def open(path: str):
    return Dataset(path)


class ExcelDocument:

    def __init__(self, path: str, template: Union[str, Dict]):
        """
        Create a new xlsx document
        :param path: path of the new document
        """
        self.file: str = path
        self.pointer: Pointer = Pointer(0, 0)
        self.xlsxDocument: xlsxwriter.Workbook = xlsxwriter.Workbook(path)
        self.magicSheets: List[MagicSheet] = []

        # init sheet
        if isinstance(template, str):
            # using file template
            # get headers
            headers = self._get_header(template)
            for key in headers.keys():
                m_sheet = MagicSheet(key.name, key)
                m_sheet.headers = headers[key]
            # write to the empty book
            self.append(template)

        elif isinstance(template, dict):
            # create sheets
            for key in template.keys():
                m_sheet = self.add_sheet(key, template[key])
                self.magicSheets.append(m_sheet)
        else:
            raise TypeError(f'Expected str or dict, got {str(type(template))} instead')

    def close(self):
        self.xlsxDocument.close()

    def get_raw_sheets(self):
        result = []
        for s in self.magicSheets:
            result.append(s.raw_sheet)
        return result

    def append(self, path: str) -> None:
        """
        add another document to the current this document
        :param path: path of the document
        :return: None
        """
        workbook = xlrd.open_workbook(path)
        sheet: xlrd.sheet.Sheet
        for sheet in workbook.sheets():
            sheet_info = self._get_sheet(sheet.name)
            if sheet_info is not None:
                # add to sheet
                self.pointer = Pointer(sheet_info.rows, 0)
                self._write_sheet(sheet, workbook, sheet_info, sheet_info.raw_sheet)
            else:
                # create new sheet
                self.pointer = Pointer(0, 0)
                new_sheet = self.xlsxDocument.add_worksheet(sheet.name)
                sheet_info = MagicSheet(sheet.name, new_sheet)
                self.magicSheets.append(sheet_info)
                self._write_sheet(sheet, workbook, sheet_info, new_sheet)

    def add_sheet(self, name, header: List[str]):
        """
        add a new sheet
        :param name: name of the sheet
        :param header: headers of the sheet
        :return:
        """
        if self._get_sheet(name) is None:
            sheet = self.xlsxDocument.add_worksheet(name)
            m_sheet = MagicSheet(name, sheet)
            for i in range(header.__len__()):
                sheet.write_string(0, i, header[i])
                m_sheet.headers.append(header[i])
            return m_sheet

    def _get_sheet(self, sheet_name) -> Union[MagicSheet, None]:
        """
        Get SheetInfo instance if it exists in self.sheets
        :param sheet_name: name of a sheet
        :return: SheetInfo if exists, None if not
        """
        for s in self.magicSheets:
            if s.name == sheet_name:
                return s
        else:
            return None

    def _get_header(self, path: str) -> Dict:
        result = {}
        workbook = xlrd.open_workbook(path)
        sheet: xlrd.sheet.Sheet
        for sheet in workbook.sheets():
            headers = []
            for cell in sheet.row(0):
                headers.append(cell.value)
            result[sheet] = headers
        return result

    def _write_sheet(self, sheet, workbook, sheet_info, this_sheet):
        for row in sheet.get_rows():
            cell: xlrd.sheet.Cell
            for cell in row:
                if cell.xf_index == XL_CELL_EMPTY or cell.xf_index == XL_CELL_BLANK:
                    this_sheet.write_string(self.pointer.row, self.pointer.col, '')

                elif cell.xf_index == XL_CELL_TEXT:
                    this_sheet.write_string(self.pointer.row, self.pointer.col, cell.value)

                elif cell.xf_index == XL_CELL_NUMBER:
                    this_sheet.write_number(self.pointer.row, self.pointer.col, cell.value)

                elif cell.xf_index == XL_CELL_DATE:
                    content = datetime.datetime(*xlrd.xldate_as_tuple(cell.value, workbook.datemode))
                    this_sheet.write_datetime(self.pointer.row, self.pointer.col, content)

                elif cell.xf_index == XL_CELL_BOOLEAN:
                    this_sheet.write_boolean(self.pointer.row, self.pointer.col, cell.value == 1)

                else:
                    this_sheet.write(self.pointer.row, self.pointer.col, cell.value)
                self.pointer.next_col()
                sheet_info._add_col()

            sheet_info._add_row()
            self.pointer.next_row()
