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

- `create_document(path, template) -> ExcelDocument`

- `open(path) -> Dataset`

### ExcelDocument

- `merge(path) -> None`
  
  - append another file to your document

- `add_sheet(name: str, header: list[str]) -> MagicSheet`
  
  - add a sheet to your file

- `close() -> None`
  
  - close and save

### MagicSheet

- `append_row(content: Union[Dict, List[str]])`
  
  - append one new row

## dataset Module

### Dataset

- `save()`
  
  - save your stuff

### Table

- `find(**kwargs)`
  
  - find a list of rows

- `filter(callback)`
  
  - return a list of rows, filter by the callback function. return True if you want it

- `append(content: dict)`
  
  - add a row to your file, dict keys should be your headers

- `remove(row: dict)`
  
  - remove a row

## utils Module

### Document

- `split_sheets(out='', out_prefix='')`
  
  - split file by sheets

- `split_rows(row_count: int, out='', out_prefix='')`
  
  - split file by rows

# Use Cases

## CRUD using Dataset

let's say we have a file like this:

| id  | name     | age |
| --- | -------- | --- |
| 1   | John Doe | 12  |
| 2   | Kelly    | 18  |

```python
ds = Dataset('a_file.xlsx')
table = ds.get_table(0)
search_results = table.find(name='John Doe')
table.remove(search_results[0])
search_results[1]['name'] = 'Vicky'
# we can leave age empty if we do it like this!
table.append({'id': '3', 'name': 'Dick'})
# we can use filter if we have even more complex conditions
filter_results = table.filter(lambda row: row['age'] is '')
# don't forget to save!
table.save()
```
