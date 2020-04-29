from excel_magic.dataset import Dataset

ex = Dataset(path="haha.xlsx")
ex.add_sheet("sheeeet5",fields=["products","price"])
sh = ex.get_sheet(0)
# index starts at 0 like list,then 1,2,3...
for i in range(15):
    sh.append_row(["lightsaber",999])
ex.save()
# Don't forget to save you excel file.



