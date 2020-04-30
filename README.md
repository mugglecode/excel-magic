# Excel MAGIC

Do magic to your excel file!

# Use Case

## CRUD using Dataset

let's say we have a file like this:

| id  | name     | age |
| --- | -------- | --- |
| 1   | John Doe | 12  |
| 2   | Kelly    | 18  |

```python
from excel_magic.dataset import Dataset
ds = Dataset('a_file.xlsx')
table = ds.get_sheet(0)
search_results = table.find(name='John Doe')
table.remove(search_results[0])
search_results[1]['name'] = 'Vicky'
# we can leave age empty if we do it like this!
table.append({'id': '3', 'name': 'Dick'})
# we can use filter if we have even more complex conditions
filter_results = table.filter(lambda row: row['age'] is '')
# don't forget to save!
ds.save()
```

## API Reference

---

## dataset Module

### Style

represents the style of a cell

| attribute            | description                         | default |
| -------------------- | ----------------------------------- | ------- |
| horizontal_alignment | how text align in a cell            | left    |
| vertical_alignment   | how text align vertically in a cell | top     |
| bold                 | is text bold                        | False   |
| underline            | is text underlined                  | False   |
| font_color           | color of the font                   | black   |
| font_name            | name of the font                    | Calibri |
| font_size            | font size                           | 12      |
| fill_color           | fill color                          | ''      |

### Cell

- `set_style(style: Style)`
  
  - se style of the cell

### Dataset

- `get_sheet(index: int)`
  
  - get sheet by index

- `get_sheet_by_name(name: str)`
  
  - get sheet by name

- `does_exist(name: str)`
  
  - check if a sheet exists

- `merge_file(path: str)`
  
  - merge another file to the current file
     
- `export_json(out: str)`
  - export all sheets to a json file

- `remove_sheet(sheet: Sheet)`
  
  - remove a sheet

- `save()`
  
  - save your stuff

### Sheet

- `find(**kwargs)`
  
  - find a list of rows

- `filter(callback)`
  
  - return a list of rows, filter by the callback function. return True if you want it

- `append_row(content: Union[dict, List[str]])`
  
  - add a row to your file, dict keys should be your headers

- `remove(row: dict)`
  
  - remove a row

- `set_header_style(style: Style)`
  
  - set style of the header
  
- `get_rows() -> List[dict]`
  - get all rows
  
- `print_row(index: int) -> str`
  - return a string of a row ready to be print

- `set_row_style(row: Union[dict, int], style: Style)`
  
  - set style of a row

- `to_csv(out: str = '')`
  
  - Convert sheet to csv

- `to_json(out: str = '')`
  
  - Convert sheet to json

- `import_json(path: str)`
  - Import data from a json file

- `beautify(by: str) -> List[dict]`
  - Group data by a column

## utils Module

### Document

- `split_sheets(out='', out_prefix='')`
  
  - split file by sheets

- `split_rows(row_count: int, out='', out_prefix='')`
  
  - split file by rows
