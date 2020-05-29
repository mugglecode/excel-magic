from typing import List, Tuple, Union

from excel_magic.dataset import Sheet, Row


class DiffSet:
    def __init__(self):
        self.not_found_in_a: Sheet = None
        self.not_found_in_b: Sheet = None

    def get_notfound(self, in_sheet: str):
        if in_sheet == 'a':
            return self.not_found_in_a
        elif in_sheet == 'b':
            return self.not_found_in_b
        else:
            raise ValueError(f'There is no sheet {in_sheet}!')

    def __str__(self):
        return f'Summary: {self.not_found_in_b.__len__() + self.not_found_in_a.__len__()} differences in total.' \
               f'\nnot found in a: {self.not_found_in_a.__len__()}' \
               f'\nnot found in b: {self.not_found_in_b.__len__()}'


class StrictDiffRow:
    def __init__(self, row_a: Union[Row, None], row_b: Union[Row, None], diff_col: List[str], row_index: int):
        self.row_a: Union[Row, None] = row_a
        self.row_b: Union[Row, None] = row_b
        self.row_index = row_index
        self.diff_col: List[str] = diff_col


class StrictDiffSet:
    def __init__(self):
        self.diff: List[StrictDiffRow] = []

    def filter_diff_in(self, col: str) -> List[StrictDiffRow]:
        result = []
        for i in self.diff:
            if col in i.diff_col:
                result.append(i)
        return result

    def append(self, row: StrictDiffRow):
        self.diff.append(row)


def diff(sheet_a: Sheet, sheet_b: Sheet) -> DiffSet:
    """
    This will search every row from sheet a in sheet b
    :param sheet_a: sheet a, rows to search
    :param sheet_b: sheet b, where rows are searched in
    :return: DiffSet containing rows in sheet a that are not found in sheet b and reversed
    """
    result = DiffSet()

    not_found_in_b: Sheet = Sheet(sheet='diff_b')
    not_found_in_b.fields = [*sheet_a.fields]
    not_found_in_a: Sheet = Sheet(sheet='diff_a')
    not_found_in_a.fields = [*sheet_a.fields]
    for row in sheet_a.get_rows():
        r = sheet_b.find(**row, none_if_not_found=True)
        if r is None:
            not_found_in_b.append_row(row)

    for row in sheet_b.get_rows():
        r = sheet_a.find(**row, none_if_not_found=True)
        if r is None:
            not_found_in_a.append_row(row)

    result.not_found_in_a = not_found_in_a
    result.not_found_in_b = not_found_in_b

    return result


def strict_diff(sheet_a: Sheet, sheet_b: Sheet) -> StrictDiffSet:
    result = StrictDiffSet()
    if sheet_a.fields != sheet_b.fields:
        raise ValueError('Strict diff won\'t work with two sheets with different headers')
    i = 0
    for i in range(sheet_a.data_rows.__len__()):
        if i+1 > sheet_b.data_rows.__len__():
            row = StrictDiffRow(sheet_a.data_rows[i], None, [*sheet_a.fields], i)
            result.append(row)

        if sheet_a.data_rows[i] != sheet_b.data_rows[i]:
            diff_cols = []
            for k in sheet_a.data_rows[i]:
                if sheet_a.data_rows[i][k] != sheet_b.data_rows[i][k]:
                    diff_cols.append(k)
            row = StrictDiffRow(sheet_a.data_rows[i], sheet_b.data_rows[i], diff_cols, i)
            result.append(row)

    if i+1 < sheet_b.data_rows.__len__():
        for j in range(i, sheet_b.__len__()):
            row = StrictDiffRow(None, sheet_b.data_rows[i], [*sheet_a.fields], i)
            result.append(row)

    return result
