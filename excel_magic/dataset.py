import xlrd
from typing import Callable, Union, List
import os
import shutil
import xlsxwriter

__all__ = ['Sheet', 'Dataset']


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


class HorizontalAlignment:
    LEFT = 'left'
    CENTER = 'center'
    RIGHT = 'right'


class VerticalAlignment:
    TOP = 'top'
    CENTER = 'center'
    BOTTOM = 'bottom'


class Style:
    def __init__(self):
        self.horizontal_alignment = 'left'
        self.vertical_alignment = 'top'
        self.bold = False
        self.underline = False
        self.font_color = 'black'
        self.font_name = 'Calibri'
        self.font_size = '12'
        self.fill_color = 'white'

    def attr(self):
        return {'align': self.horizontal_alignment,
                'valign': self.vertical_alignment,
                'bold': self.bold,
                'underline': self.underline,
                'font_color': self.font_color,
                'font_name': self.font_name,
                'font_size': self.font_size,
                'bg_color': self.fill_color}


class Cell:
    def __init__(self, value='', style: Style = None):
        self.value = value
        if style is None:
            self.style = Style()
        else:
            self.style = style

    def set_style(self, style: Style):
        self.style = style

    def attr(self):
        return self.style.attr()


class Sheet:

    def __init__(self, sheet: Union[xlrd.sheet.Sheet, str] = ''):
        self.fields = []
        self.data_rows: List[dict] = []
        self.header_style: Style = Style()
        if isinstance(sheet, str):
            self.name = sheet
        else:
            self.name = sheet.name
            self._init_fields(sheet)
            self._init_data(sheet)

    def _init_fields(self, sheet: xlrd.sheet.Sheet):
        fields_row = sheet.row(0)
        for field in fields_row:
            self.fields.append(field.value)

    def _init_data(self, sheet: xlrd.sheet.Sheet):
        flg_first_row = True
        for row in sheet.get_rows():
            # skip the first row
            if flg_first_row:
                flg_first_row = False
                continue

            new_row = {}
            for i in range(len(self.fields)):
                # to prevent bug when there is an empty cell
                if i < len(row):
                    new_row[self.fields[i]] = Cell(row[i].value)
                else:
                    new_row[self.fields[i]] = ''
            self.data_rows.append(new_row)

    def set_header_style(self, style: Style):
        self.header_style = style

    def find(self, **kwargs):
        result = []
        # Check kwargs
        for kwarg in kwargs:
            if kwarg not in self.fields:
                raise NameError(f'field {kwarg} not found')

        for data_row in self.data_rows:
            for key in kwargs.keys():
                if data_row[key].value != kwargs[key]:
                    break
            else:
                result.append(data_row)
        return result

    def filter(self, callback: Callable[[dict], Union[None, bool]]):
        data_list = []

        for row in self.data_rows:
            result = callback(row)
            if bool(result):
                data_list.append(row)

        return data_list

    def append(self, content: Union[dict, List[str]]):
        new_row = {}
        if isinstance(content, dict):
            for field in self.fields:
                if field in content.keys():
                    new_row[field] = Cell(content[field])
                else:
                    new_row[field] = Cell('')
        elif isinstance(content, list):
            if content.__len__() != self.fields.__len__():
                raise ValueError(f'Expected {self.fields.__len__()} values, got {content.__len__()}')
            for i in range(len(self.fields)):
                new_row[self.fields[i]] = Cell(content[i])
        else:
            raise TypeError('Expected dict or list}')
        self.data_rows.append(new_row)

    def get_col(self, col: str):
        if col not in self.fields:
            raise NameError(f'field "{col}" does not exists')
        col = []
        for row in self.data_rows:
            col.append(row[col])
        return col

    def set_row_style(self, row: Union[dict, int], style: Style):
        pass

    def remove(self, row: dict):
        self.data_rows.remove(row)


class Dataset:

    # TODO: content.id

    def __init__(self, path: str):
        if not os.path.isfile(path):
            wb = xlsxwriter.Workbook(path)
            wb.close()
        self.workbook = xlrd.open_workbook(path, on_demand=True)
        self.sheets = []
        self.filename = os.path.basename(path)
        self.backup_name = self.filename + '.bak'
        self.path = os.path.dirname(path)
        sheet: xlrd.sheet.Sheet
        for sheet in self.workbook.sheets():
            try:
                sheet.row(0)
            except IndexError:
                continue
            self.sheets.append(Sheet(sheet))
            self.workbook.unload_sheet(sheet.name)

    def get_sheet(self, index: int) -> Sheet:
        return self.sheets[index]

    def get_sheet_by_name(self, name: str) -> Union[Sheet, None]:
        for t in self.sheets:
            if t.name == name:
                return t
        else:
            return None

    def does_exists(self, name: str) -> bool:
        for t in self.sheets:
            if t.name == name:
                return True
        else:
            return False

    def filter(self, table: Sheet, callback: Callable[[dict], Union[None, bool]]) -> List[dict]:
        return table.filter(callback)

    def find(self, sheet: Sheet, **kwargs):
        result = sheet.find(**kwargs)

        return result

    def append(self, sheet: Sheet, content: dict):
        sheet.append(content)

    def add_sheet(self, name: str, fields: List[str]) -> Sheet:
        table = Sheet(name)
        table.fields = fields
        return table

    def merge_file(self, path: str) -> None:
        workbook = xlrd.open_workbook(path)
        sheet: xlrd.sheet.Sheet
        for sheet in workbook.sheets():
            tbl = self.get_sheet_by_name(sheet.name)
            if tbl is not None:
                self._merge_table(sheet, tbl)
            else:
                tbl = Sheet(sheet.name)
                try:
                    headers = sheet.row(0)
                except IndexError:
                    raise ValueError('File has no headers')
                for h in headers:
                    tbl.fields.append(h.value)
                self._merge_table(sheet, tbl)

    def _merge_table(self, sheet, tbl):
        flg_first_row = True
        for row in sheet.get_rows():
            # Skip header
            if flg_first_row:
                flg_first_row = False
                continue

            new_row = []
            for cell in row:
                new_row.append(cell)
            tbl.append(new_row)

    def remove_sheet(self, sheet: Sheet) -> None:
        self.sheets.remove(sheet)

    def save(self):
        # make backup & delete
        shutil.copy(os.path.join(self.path, self.filename), os.path.join(self.path, self.backup_name))
        os.remove(os.path.join(self.path, self.filename))

        # open new file
        filename = os.path.join(self.path, self.filename)
        workbook = xlsxwriter.Workbook(filename)
        for table in self.sheets:
            sheet = workbook.add_worksheet(table.name)
            pointer = Pointer(0, 0)
            for field in table.fields:
                sheet.write(pointer.row, pointer.col, field, table.header_style)
                pointer.next_col()
            pointer.next_row()
            for data_row in table.data_rows:
                for data in data_row.values():
                    sheet.write(pointer.row, pointer.col, data.value, workbook.add_format(data.attr()))
                    pointer.next_col()
                pointer.next_row()
        workbook.close()
