from excel_magic.dataset import Dataset,open_file

ex_1 = open_file(path="example.xlsx")
ex_1.save()
ex_2 = Dataset(path="example-2.xlsx")
ex_2.save()
with Dataset("example-3.xlsx") as ex_3:
    # Open file without worrying close file
    # Do stuff here
    pass