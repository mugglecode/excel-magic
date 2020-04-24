import xlrd
from typing import Callable, Union, List

__all__ = ['Table', 'Dataset']

class Table:

    def __init__(self, sheet: xlrd.sheet.Sheet):
        self.fields = []
        self.data_rows = []

        self.init_fields(sheet)
        self.init_data(sheet)

    def init_fields(self, sheet: xlrd.sheet.Sheet):
        fields_row = sheet.row(0)
        for field in fields_row:
            self.fields.append(field.value)

    def init_data(self, sheet: xlrd.sheet.Sheet):
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


class Dataset:

    # TODO: content.id

    def __init__(self, path: str):
        self.workbook = xlrd.open_workbook(path, on_demand=True)
        self.tables = []
        sheet: xlrd.sheet.Sheet
        for sheet in self.workbook.sheets():
            try:
                sheet.row(0)
            except:
                continue
            self.tables.append(Table(sheet))
            self.workbook.unload_sheet(sheet.name)

    def filter(self, callback: Callable[[dict], Union[None, bool]], table: int=0) -> List[dict]:
        data_list = []

        table = self.tables[table]

        for row in table.data_rows:
            result = callback(row)
            if bool(result):
                data_list.append(row)

        return data_list

