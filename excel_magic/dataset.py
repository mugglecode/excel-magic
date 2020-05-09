import datetime
import sys
from copy import copy
import sqlite3
from io import BytesIO
import xlrd
from typing import Callable, Union, List, Any
import os
import shutil
import xlsxwriter
import csv
import json
from PIL import Image

__all__ = ['Sheet', 'Dataset', 'open_file']


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
        self.fill_color = ''
        self.num_format = ''

    def attr(self):
        attr = {'align': self.horizontal_alignment,
                'valign': self.vertical_alignment,
                'bold': self.bold,
                'underline': self.underline,
                'font_color': self.font_color,
                'font_name': self.font_name,
                'font_size': self.font_size}
        if self.fill_color != '':
            attr['bg_color'] = self.fill_color
        if self.num_format != '':
            attr['num_format'] = self.num_format


class Header:
    def __init__(self, value: str, style: Style, width: int = 20):
        self.value = value
        self.style = style
        self.width = width


class Cell:
    def __init__(self, value: Any = '', style: Style = None):
        self._value = value
        if style is None:
            self.style = Style()
        else:
            self.style = style

    @property
    def value(self):
        if isinstance(self._value, float) and self._value % 1 == 0:
            return int(self._value)
        else:
            return self._value

    @value.setter
    def value(self, value):
        self._value = value

    def set_style(self, style: Style):
        self.style = style

    def attr(self):
        return self.style.attr()

    def __str__(self):
        return str(self.value)

    def __eq__(self, other: Union['Cell', str]):
        if isinstance(other, str):
            return self.value == other
        elif isinstance(other, Cell):
            return self.value == other.value and\
                   self.style == other.style
        else:
            if isinstance(self.value, type(other)):
                return self.value == other
            else:
                if isinstance(other, int):
                    return self.value == float(other)
                return self.value is other


class ImageCell(Cell):
    def __init__(self, data: Union[BytesIO, str]):
        super().__init__()
        self.data = data
        self.value = ''

class Sheet:
    def __init__(self, path: str, sheet: Union[xlrd.sheet.Sheet, str] = ''):
        self.fields = []
        self.data_rows: List[dict] = []
        self.header_style: Style = Style()
        self.filename = path
        if isinstance(sheet, str):
            self.name: str = sheet
        else:
            self.name: str = sheet.name
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
                    if row[i].ctype == 3:
                        c = Cell(datetime.datetime(*xlrd.xldate_as_tuple(row[i].value, sheet.book.datemode)))
                        c.style.num_format = 'yyyy/mm/dd'
                        new_row[self.fields[i]] = c
                    else:
                        if isinstance(row[i].value, str):
                            if row[i].value.isnumeric():
                                print('Warning: Found a number stored in string format, converting...')
                                new_row[self.fields[i]] = Cell(float(row[i].value))
                        new_row[self.fields[i]] = Cell(row[i].value)
                else:
                    new_row[self.fields[i]] = ''
            self.data_rows.append(new_row)

    def set_header_style(self, style: Style):
        self.header_style = style

    def find(self, pairs: Union[dict, None] = None, **kwargs) -> List[dict]:
        result = []
        if pairs is not None:
            kwargs = pairs
        # Check kwargs
        for kwarg in kwargs:
            if kwarg not in self.fields:
                raise NameError(f'field {kwarg} not found')

        for data_row in self.data_rows:
            for key in kwargs.keys():
                if isinstance(kwargs[key], int):
                    if data_row[key].value != float(kwargs[key]):
                        break
                    else:
                        continue

                if data_row[key].value != kwargs[key]:
                    break
            else:
                result.append(data_row)
        return result

    def filter(self, callback: Callable[[dict], Union[None, bool]]) -> List[dict]:
        data_list = []

        for row in self.data_rows:
            result = callback(row)
            if bool(result):
                data_list.append(row)

        return data_list

    def append_row(self, content: Union[dict, List[Union[str, Cell]]]) -> None:
        new_row = {}
        if isinstance(content, dict):
            for field in self.fields:
                if field in content:
                    if isinstance(content[field], Cell):
                        new_row[field] = content[field]
                    else:
                        new_row[field] = Cell(content[field])
                else:
                    new_row[field] = Cell('')
        elif isinstance(content, list):
            for i in range(len(self.fields)):
                if isinstance(content[i], Cell):
                    new_row[self.fields[i]] = content[i]
                else:
                    new_row[self.fields[i]] = Cell(content[i])

            if len(content) < len(self.fields):
                for i in range(len(content) - 1, len(self.fields)):
                    new_row[self.fields[i]] = Cell('')
        else:
            raise TypeError('Expected dict or list}')
        self.data_rows.append(new_row)

    def get_rows(self) -> List[dict]:
        r = [*self.data_rows]
        return r

    def get_col(self, col: str):
        if col not in self.fields:
            raise NameError(f'field "{col}" does not exists')
        result = []
        for row in self.data_rows:
            result.append(row[col])
        return result

    def print_row(self, index: int):
        row = self.data_rows[index]
        result = ''
        for k in row:
            result += f'{k}: {row[k].value}, '
        return result

    def set_row_style(self, row: Union[dict, int], style: Style) -> None:
        if isinstance(row, int):
            row = self.data_rows[row]

        c: Cell
        for c in row:
            c.style = style

    def remove_row(self, row: dict) -> None:
        self.data_rows.remove(row)

    def import_json(self, path: str) -> None:
        with open(path, 'r') as f:
            data = json.load(f)
        if not isinstance(data, list):
            raise ValueError('invalid file format')
        for row in data:
            self.append_row(row)

    def to_csv(self, out: str = '') -> None:
        if out == '':
            out = self.name + '.csv'

        with open(out, 'w') as f:
            w = csv.DictWriter(f, self.fields)
            v = {}
            for r in self.data_rows:
                for key in r:
                    v[key] = r[key].value
                w.writerow(v)

    def to_json(self, out: str = '') -> None:
        if out == '':
            out = self.name + '.json'
        data = []
        for r in self.data_rows:
            v = {}
            for k in r:
                v[k] = r[k].value
            data.append(v)
        with open(out, 'w') as f:
            json.dump(data, f)

    def split_rows(path: str, row_count: int, name_by: str):
        filenames = {}

    def beautify(self, by: str) -> List[dict]:
        if isinstance(by, str):
            grouped = []
            ungrouped = copy(self.data_rows)
            while ungrouped.__len__() > 0:
                counter = 0
                current = ungrouped[0][by].value
                while counter < ungrouped.__len__():
                    if ungrouped[counter][by].value == current:
                        grouped.append(ungrouped[counter])
                        ungrouped.remove(ungrouped[counter])
                        counter -= 1

                    counter += 1
            return grouped

    def __eq__(self, other):
        if self.name == other.name:
            return True
        return False


class Dataset:

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
            self.sheets.append(Sheet(path, sheet))
            self.workbook.unload_sheet(sheet.name)

    def get_sheet_by_index(self, index: int) -> Sheet:
        return self.sheets[index]

    def get_sheet_by_name(self, name: str) -> Union[Sheet, None]:
        t: Sheet
        for t in self.sheets:
            if t.name.lower() == name.lower():
                return t
        else:
            return None

    def does_exist(self, name: str) -> bool:
        for t in self.sheets:
            if t.name == name:
                return True
        else:
            return False

    def filter(self, table: Sheet, callback: Callable[[dict], Union[None, bool]]) -> List[dict]:
        return table.filter(callback)

    def find(self, sheet: Sheet, **kwargs) -> List[dict]:
        result = sheet.find(**kwargs)

        return result

    def append_row(self, sheet: Union[Sheet, str], content: dict) -> None:
        if isinstance(sheet, str):
            sheet = self.get_sheet_by_name(sheet)
        if sheet is None:
            raise NameError(f'{sheet} does not exist')
        sheet.append_row(content)

    def add_sheet(self, name: str, fields: List[str]) -> Sheet:
        if self.does_exist(name):
            raise Exception('Sheet already exists')
        table = Sheet(self.filename, name)
        table.fields = fields
        self.sheets.append(table)
        return table

    def create_sheet_by_json(self, name: str, data: Union[str, list, dict]) -> Sheet:
        if isinstance(data, str):
            with open(data, 'r') as f:
                data: Union[list, dict] = json.load(f)

        if isinstance(data, list):
            header = data[0].keys()
        elif isinstance(data, dict):
            header = data.keys()
        else:
            raise ValueError('corrupted file')
        sheet = self.add_sheet(name, header)
        if isinstance(data, list):
            for d in data:
                sheet.append_row(d)
        return sheet

    def import_json(self, path: str) -> None:
        with open(path, 'r') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError('invalid format')

        for key in data:
            self.create_sheet_by_json(key, data[key])

    def export_json(self, out: str):
        json_sheets = {}
        for sheet in self.sheets:
            data = []
            for r in sheet.data_rows:
                v = {}
                for k in r:
                    v[k] = r[k].value
                data.append(v)
            json_sheets[sheet.name] = data
        with open(out, 'w') as f:
            json.dump(json_sheets, f)

    def to_sqlite(self, out: str):
        conn = sqlite3.connect(out)
        cur = conn.cursor()
        current_table = ''
        for sheet in self.sheets:
            current_table = sheet.name
            cmd = f"CREATE TABLE '{current_table}' ({','.join(sheet.fields)})"
            cur.execute(cmd)
            conn.commit()
            for row in sheet.data_rows:
                values = ''
                for cell in row.values():
                    if isinstance(cell.value, float):
                        values += str(cell.value)
                    else:
                        values += '"' + cell.value + '"'
                    values += ','
                values = values[: -1]
                cmd = f"INSERT INTO {current_table} VALUES ({values})"
                cur.execute(cmd)
        conn.commit()
        conn.close()

    def merge_file(self, path: str, force: bool = False) -> None:
        workbook = xlrd.open_workbook(path)
        sheet: xlrd.sheet.Sheet
        for sheet in workbook.sheets():
            tbl = self.get_sheet_by_name(sheet.name)
            if tbl is not None:
                if force:
                    headers_to_merge = sheet.row(0)
                    for i in range(len(headers_to_merge)):
                        headers_to_merge[i] = headers_to_merge[i].value
                    for h in headers_to_merge:
                        if h in tbl.fields:
                            headers_to_merge.remove(h)
                    tbl.fields.extend(headers_to_merge)
                self._merge_table(sheet, tbl)
            else:
                tbl = Sheet(self.filename, sheet.name)
                try:
                    headers = sheet.row(0)
                except IndexError:
                    raise ValueError('File has no headers')
                for h in headers:
                    tbl.fields.append(h.value)
                self._merge_table(sheet, tbl)
                self.sheets.append(tbl)

    def _merge_table(self, sheet, tbl, force: bool = False):
        flg_first_row = True
        for row in sheet.get_rows():
            # Skip header
            if flg_first_row:
                flg_first_row = False
                continue

            new_row = []
            for cell in row:
                new_row.append(cell.value)
            tbl.append_row(new_row)

    def split_sheets_to_file(self):
        for s in self.sheets:
            doc = open_file(s.name + '.xlsx')
            doc.add_sheet(s.name, s.fields)
            for row in s.data_rows:
                doc.append_row(s.name, row)
            doc.save(backup=False)

    def remove_sheet(self, sheet: Sheet) -> None:
        self.sheets.remove(sheet)

    def remove_sheet_by_index(self, index: int):
        pass

    def save(self, backup = True, row_height = 0, col_width = 0):
        # make backup & delete
        if os.path.isfile(os.path.join(self.path, self.filename)) and backup:
            shutil.copy(os.path.join(self.path, self.filename), os.path.join(self.path, self.backup_name))
            os.remove(os.path.join(self.path, self.filename))

        # open new file
        filename = os.path.join(self.path, self.filename)
        workbook = xlsxwriter.Workbook(filename, {'default_date_format':
                                                  'yyyy/mm/dd'})
        for table in self.sheets:
            sheet = workbook.add_worksheet(table.name)
            pointer = Pointer(0, 0)
            for field in table.fields:
                sheet.write(pointer.row, pointer.col, field, workbook.add_format(table.header_style.attr()))
                pointer.next_col()
            pointer.next_row()
            for data_row in table.data_rows:
                for data in data_row.values():
                    if isinstance(data.value, datetime.datetime):
                        if data.value.time().min == 0 and data.value.time().hour == 0 and data.value.time().second == 0:
                            sheet.write(pointer.row, pointer.col, str(data.value.date().isoformat()), workbook.add_format(data.attr()))
                        elif data.value.date().year == 0 and data.value.date().month == 0 and data.value.date().day == 0:
                            sheet.write(pointer.row, pointer.col, str(data.value.time().isoformat()), workbook.add_format(data.attr()))
                        else:
                            sheet.write(pointer.row, pointer.col, str(data.value), workbook.add_format(data.attr()))
                    else:
                        if isinstance(data, ImageCell):
                            if isinstance(data.data, str):
                                sheet.insert_image(pointer.row, pointer.col, data.data, {'y_offset': 10, 'x_offset': 10})
                            else:
                                data.data.seek(0)
                                sheet.insert_image(pointer.row, pointer.col, data.value, {'image_data': data.data, 'y_offset': 10, 'x_offset': 10})
                            img: Image.Image = Image.open(data.data)
                            width, height = img.size
                            if row_height == 0:
                                sheet.set_row(pointer.row, height)
                            else:
                                sheet.set_row(pointer.row, row_height)
                            if col_width != 0:
                                sheet.set_column(pointer.col, pointer.col, (col_width / 8))
                            else:
                                sheet.set_column(pointer.col, pointer.col, (width / 8))

                        else:
                            sheet.write(pointer.row, pointer.col, data.value, workbook.add_format(data.attr()))
                    pointer.next_col()
                pointer.next_row()
        workbook.close()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.save()

    def __enter__(self):
        return self


def open_file(path: str) -> Dataset:
    return Dataset(path)
