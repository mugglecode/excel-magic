# Excel MAGIC

> Simplify common Excel operations.

  
![Logo](https://raw.githubusercontent.com/guo40020/excel-magic/master/logo.png)  
[![PyPI version](https://badge.fury.io/py/excelmagic.svg )](https://pypi.org/project/excelmagic)


## Installation

```shell script
pip install excelmagic
```

## Usage

### Opening an Excel file

```python
from excel_magic2.dataset import open_file

file = open_file('test.xlsx')
```

Also supports **with** statement:

```python
from excel_magic2.dataset import open_file

with open_file('test.xlsx') as file:
        pass
```

### Query rows

Example data:

| Id  | Name  | Age | Score |
| --- | ----- | --- | ----- |
| 1   | John  | 22  | 89    |
| 2   | David | 23  | 93    |
| 3   | Emma  | 22  | 95    |

Query rows of a sheet in an excel file with **specific cell value**:

```python
from excel_magic2.dataset import open_file

with open_file('test.xlsx') as excel:
    # select a sheet by index or sheet name
    sheet = excel.get_sheet_by_index(0)
    # find rows containing the name 'David'
    rows = sheet.find(Name='David')
```

Or query rows with **callback function**:

```python
from excel_magic2.dataset import open_file

def score_over_90(rows):
    if rows['Score'].value > 90:
        return True

with open_file('test.xlsx') as excel:
    sheet = excel.get_sheet_by_index(0)
    # find rows with the score column greater than 90
    rows = sheet.filter(score_over_90)
```

### Getting cell values from a row

You can use `key: value` method to get the cell object in a rod, like operating a dict.

```python
cell = row['Score']
```

And get the value of the cell object through the **`value` attribute**.

```python
score_num = cell.value
```

### Split sheets

Split multiple sheets of excel file to independent excel files.

```python
from excel_magic2.dataset import open_file

file = open_file('test.xlsx')
file.split_sheets_to_files()
```

### Merge files

Combine sheets from files into a new excel file.

```python
from excel_magic2.dataset import open_file

excel_files = ['01.xlsx', '02.xlsx', '03.xlsx']

new_excel = open_file('test.xlsx')
for file in excel_files:
    new_excel.merge_file(file)
new_excel.save()
```

Or

```python
from excel_magic2.dataset import open_file

excel_files = ['01.xlsx', '02.xlsx', '03.xlsx']

with open_file('test.xlsx') as new_excel:
    for file in excel_files:
        new_excel.merge_file(file)
```

## API Reference

The hierarchical relationship in the excel file is:

> Excel (sheets) → Sheet → Row → Cell

And excelmagic provides similar hierarchical object API:

> Dataset Object → Sheet Object → Row Object → Cell Object

### Dataset Object

**Example:**

```python
from excel_magic2.dataset import open_file

dataset = open_file('test.xlsx')
```

**Methods:**

Search Sheet

- `get_sheet_by_index(index: int) -> Sheet`
  - get a sheet object by sheet index.
- `get_sheet_by_name(name: str) -> Sheet`
  - get a sheet object by sheet name.
- `does_exist(name: str) -> bool`
  - check if sheet name exists in your Dataset.

Create Sheet

- `add_sheet(name: str, fields: List[str]) -> Sheet`
  - append new sheet with sheet name and column headers.

Delete Sheet

- `remove_sheet(sheet: Sheet) -> None`
  - remove a sheet by passing a sheet object.

Others

- `save() -> None`
  
  - save changes.

- `split_sheets_to_files() -> None`
  
  - split multiple sheets to independent excel files.

- `merge_file(path: str) -> None`
  
  - merge another excel file to the current file.

- `export_json(out: str) -> None`
  
  - export all sheets to a json file.

### Sheet Object

**Example:**

```python
from excel_magic2.dataset import open_file

dataset = open_file('test.xlsx')
sheet = dataset.get_sheet_by_index(0)
```

**Methods:**

Search  rows

- `find(**kwargs: dict[str, Any]) -> List[dict]`
  - return list of row which is essentially a dict.
- `filter(callback: Callable[[dict], Union[None, bool]]) -> List[dict]`
  - return list of row, filter by the callback function with which return True. And the callback receives row object (a dict) as parameter.
- `get_rows() -> List[dict]`
  - return a list of all rows.

Create row

- `append_row(content: Union[dict, List[str]]) -> None`
  - append a row to your file. If you use dict-type parameter, the keys should be same as your column headers.

Delete row

- `remove_row(row: dict) -> None`
  - find and delete a row according to dict key and value.

Export and Import sheet

- `to_csv(out: str = '') -> None`
  - export the sheet to csv file.
- `to_json(out: str = '') -> None`
  - export the sheet to json file.
- `import_json(path: str)`
  - Import a json file and insert into the sheet.

Others:

- `print_row(index: int) -> str`
  
  - return a string of a row ready to be print.

- `beautify(by: str) -> List[dict]`
  
  - group data by column header.

- `set_header_style(style: Style) -> None`
  
  - set style of the header.

- `set_row_style(row: Union[dict, int], style: Style) -> None`
  
  - set style of a row.

### Row Object

**Example:**

```python
from excel_magic2.dataset import open_file

dataset = open_file('test.xlsx')
sheet = dataset.get_sheet_by_index(0)
rows = sheet.find(Name='David')   # return a list of found row object
```

**Methods:**

The row object is dict-type, with column headers as its key and cell object as the value.

So you can get the cell object of a row with `row[key]` or `row.get(key)`, like dict type dose.

Read cell

- `row[key].value`

Update cell

- `row[key].value = new_value`

Delete cell

- `row[key].value = ''`

### Cell Object

**Example:**

```python
from excel_magic2.dataset import open_file

dataset = open_file('test.xlsx')
sheet = dataset.get_sheet_by_index(0)
rows = sheet.find(Name='David')   # return a list of found row object
cell = rows[0].get('Score')   # then use value attribute to get the value of a cell
score = cell.value
```

**Attributes:**

- `value` 
  - get the value of cell object

**Methods:**

- `set_style(style: Style) -> None`
  - passing style object, set the style of the cell

### Style Object

Create the style object for cells.

**Example:**

```python
from excel_magic2.dataset import Style

my_style = Style()
my_style.fill_color = '#52de97'
my_style.font_size = 20
my_style.bold = True

cell.set_style(my_style)
```

The following attributes have been supported:

| Attribute            | Optional Value                                | Default Value |
|:-------------------- |:--------------------------------------------- |:------------- |
| font_color           | 'red' or '(255, 0, 0)' or '#FF0000' ...       | 'black'       |
| fill_color           | 'red' or '(255, 0, 0)' or '#FF0000' ...       | ''            |
| font_name            | 'Calibri' or 'Times New Roman' or 'Arial' ... | 'Calibri'     |
| font_size            | 12 or '12' ...                                | 12            |
| bold                 | True or False                                 | False         |
| underline            | True or False                                 | False         |
| horizontal_alignment | 'left' or 'center' or 'right'                 | 'left'        |
| vertical_alignment   | 'top' or 'center' or 'bottom'                 | 'top'         |

## Built With

- [xlrd](https://pypi.org/project/xlrd/)
- [xlsxwriter](https://pypi.org/project/XlsxWriter/) 


## Authors

Kelly 

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.
