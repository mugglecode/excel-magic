import xlrd
import xlsxwriter
import os
from excel_magic.Excel import Pointer


def split_sheets(path: str, out: str = '', out_prefix: str = ''):
    original_workbook = xlrd.open_workbook(path, formatting_info=True)
    sheet: xlrd.sheet.Sheet
    for sheet in original_workbook.sheets():
        sheet_workbook = xlsxwriter.Workbook(os.path.join(out, out_prefix) + sheet.name + '.xlsx')
        new_sheet = sheet_workbook.add_worksheet(sheet.name)
        row_counter = 0
        col_counter = 0
        for row in sheet.get_rows():
            cell: xlrd.sheet.Cell
            for cell in row:
                new_sheet.write(row_counter, col_counter, cell.value)
                col_counter += 1

            col_counter = 0
            row_counter += 1
        sheet_workbook.close()


class MultipleSheetsError(Exception):
    pass


def split_rows(path: str, row_count: int, out: str = '', out_prefix: str = ''):
    file_counter = 1

    original_pointer = Pointer(0, 0)
    sheet_pointer = Pointer(0, 0)

    workbook = xlrd.open_workbook(path, formatting_info=True)
    if workbook.sheets().__len__() > 1:
        raise MultipleSheetsError('You have multiple sheets in this file')

    sheet = workbook.sheet_by_index(0)
    current_workbook = xlsxwriter.Workbook(os.path.join(out, out_prefix) + str(file_counter) + '.xlsx')
    current_sheet = current_workbook.add_worksheet(sheet.name)
    for row in sheet.get_rows():
        for cell in row:
            current_sheet.write(sheet_pointer.row, sheet_pointer.col, cell.value)
            sheet_pointer.next_col()
            original_pointer.next_col()

        if sheet_pointer.row == row_count - 1:
            sheet_pointer = Pointer(0, 0)
            file_counter += 1
            current_workbook.close()
            current_workbook = xlsxwriter.Workbook(os.path.join(out, out_prefix) + str(file_counter) + '.xlsx')
            current_sheet = current_workbook.add_worksheet(sheet.name)
        else:
            sheet_pointer.next_row()
        original_pointer.next_row()
    current_workbook.close()
