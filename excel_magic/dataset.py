import xlrd
from typing import Callable, Union, List
import os
import shutil
import xlsxwriter

__all__ = ['Table', 'Dataset']


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


class Table:

    def __init__(self, sheet: xlrd.sheet.Sheet):
        self.fields = []
        self.data_rows: List[dict] = []
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
                    new_row[self.fields[i]] = row[i].value
                else:
                    new_row[self.fields[i]] = ''
            self.data_rows.append(new_row)

    def find(self, **kwargs):
        result = []
        # Check kwargs
        for kwarg in kwargs:
            if kwarg not in self.fields:
                raise NameError(f'field {kwarg} not found')

        for data_row in self.data_rows:
            for key in kwargs.keys():
                if data_row[key] != kwargs[key]:
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

    def append(self, content: dict):
        new_row = {}
        for field in self.fields:
            if field in content.keys():
                new_row[field] = content[field]
            else:
                new_row[field] = ''
        self.data_rows.append(new_row)

    def remove(self, row: dict):
        self.data_rows.remove(row)


class Dataset:

    # TODO: content.id

    def __init__(self, path: str):
        if not os.path.isfile(path):
            wb = xlsxwriter.Workbook(path)
            wb.close()
        self.workbook = xlrd.open_workbook(path, on_demand=True)
        self.tables = []
        self.filename = os.path.basename(path)
        self.backup_name = self.filename + '.bak'
        self.path = os.path.dirname(path)
        sheet: xlrd.sheet.Sheet
        for sheet in self.workbook.sheets():
            try:
                sheet.row(0)
            except:
                continue
            self.tables.append(Table(sheet))
            self.workbook.unload_sheet(sheet.name)

    def get_table(self, index: int) -> Table:
        return self.tables[index]

    def filter(self, table: Table, callback: Callable[[dict], Union[None, bool]]) -> List[dict]:
        return table.filter(callback)

    def find(self, table: Table, **kwargs):
        result = table.find(**kwargs)

        return result

    def append(self, table: Table, content: dict):
        table.append(content)

    def save(self):
        # make backup & delete
        shutil.copy(os.path.join(self.path, self.filename), os.path.join(self.path, self.backup_name))
        os.remove(os.path.join(self.path, self.filename))

        # open new file
        filename = os.path.join(self.path, self.filename)
        workbook = xlsxwriter.Workbook(filename)
        for table in self.tables:
            sheet = workbook.add_worksheet(table.name)
            pointer = Pointer(0, 0)
            for field in table.fields:
                sheet.write(pointer.row, pointer.col, field)
                pointer.next_col()
            pointer.next_row()
            for data_row in table.data_rows:
                for data in data_row.values():
                    sheet.write(pointer.row, pointer.col, data)
                    pointer.next_col()
                pointer.next_row()
        workbook.close()
