from excel_magic.dataset import Dataset

with Dataset(path="example.xlsx") as ex:
    sh = ex.get_sheet_by_name("sheet1")
    ex.remove_sheet(sh)


with Dataset(path="example-2.xlsx") as ex_2:
    sh = ex.get_sheet_by_name("sheet1")
    for r in sh.get_rows():
        sh.remove_row(r)