from excel_magic2.dataset import Dataset

ex = Dataset(path="example.xlsx")

sheet = ex.get_sheet(1)
for row in sheet.get_rows():
    if row["name"] == "luke":
        print(row)

result = sheet.filter(lambda x:x["age"] == 34)
print(result)
