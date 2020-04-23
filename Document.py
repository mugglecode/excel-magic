import xlrd
import xlsxwriter
import datetime
from typing import List, Union

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


class SheetInfo:
    def __init__(self, name, sheet):
        self.name = name
        self.cols = 0
        self.rows = 0
        self.sheet = sheet

    def add_col(self):
        self.cols += 1

    def add_row(self):
        self.rows += 1

    def __str__(self):
        return self.name


class create_document:

    def __init__(self, path: str):
        """
        Create a new xlsx document
        :param path: path of the new document
        """
        self.file: str = path
        self.pointer: Pointer = Pointer(0, 0)
        self.xlsxDocument: xlsxwriter.Workbook = xlsxwriter.Workbook(path)
        self.sheet_info: List[SheetInfo] = []

    def close(self):
        self.xlsxDocument.close()

    def append(self, path: str) -> None:
        """
        add another document to the current this document
        :param path: path of the document
        :return: None
        """
        workbook = xlrd.open_workbook(path, formatting_info=True)
        sheet: xlrd.sheet.Sheet
        for sheet in workbook.sheets():
            sheet_info = self.get_sheet(sheet.name)
            if sheet_info is not None:
                # add to sheet
                self.pointer = Pointer(sheet_info.rows, 0)
                self._write_sheet(sheet, workbook, sheet_info, sheet_info.sheet)
            else:
                # create new sheet
                self.pointer = Pointer(0, 0)
                new_sheet = self.xlsxDocument.add_worksheet(sheet.name)
                sheet_info = SheetInfo(sheet.name, new_sheet)
                self.sheet_info.append(sheet_info)
                self._write_sheet(sheet, workbook, sheet_info, new_sheet)

    def get_sheet(self, sheet_name) -> Union[SheetInfo, None]:
        """
        Get SheetInfo instance if it exists in self.sheets
        :param sheet_name: name of a sheet
        :return: SheetInfo if exists, None if not
        """
        for s in self.sheet_info:
            if s.name == sheet_name:
                return s
        else:
            return None

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
                sheet_info.add_col()

            sheet_info.add_row()
            self.pointer.next_row()
