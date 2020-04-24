# Excel MAGIC

Do magic to your excel file!

## Quick Start

```python
# To create an empty workbook
from excel_magic import Excel
doc = Excel.create_document('/path/to/your/file.xlsx', {"Sheet1": ['id', 'name', 'age']})
# get your first sheet
sheet = doc.get_sheet(0)
# add a row
sheet.append_row([1. 'John Doe', 18])
# append another file
doc.merge('file_to_append.xlsx')
# close and save
doc.close()
```

## API Reference

---

### Excel module

- create_document(path, template) -> ExcelDocument

- open(path) -> Dataset

### ExcelDocument

- merge(path) -> None
  
  - append another file to your document

- add\_sheet(name: str, header: list[str]) -> MagicSheet
  
  - add a sheet to your file

- close() -> None
  
  - close and save

### MagicSheet

- append\_row(content: Union[Dict, List[str]])
  
  - append one new row

## dataset Module

### Dataset

- save()
  
  - save your stuff

- 

### Table

- find(\*\*kwargs)
  
  - find a list of rows

- filter(callback)
  
  - return a list of rows, filter by the callback function. return True if you want it

- append(content: dict)
  
  - add a row to your file, dict keys should be your headers

- remove(row: dict)
  
  - remove a row

## utils Module

### Document

- split\_sheets(out='', out\_prefix='')
  
  - split file by sheets

- split\_rows(row\_count: int, out='', out\_prefix='')
  
  - split file by rows
