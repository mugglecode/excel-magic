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
ds = Dataset(&#39;a_file.xlsx&#39;)
table = ds.get_sheet(0)
search_results = table.find(name=&#39;John Doe&#39;)
table.remove(search_results[0])
search_results[1][&#39;name&#39;] = &#39;Vicky&#39;
# we can leave age empty if we do it like this!
table.append({&#39;id&#39;: &#39;3&#39;, &#39;name&#39;: &#39;Dick&#39;})
# we can use filter if we have even more complex conditions
filter_results = table.filter(lambda row: row[&#39;age&#39;] is &#39;&#39;)
# don&#39;t forget to save!
table.save()
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

- `remove_sheet(sheet: Sheet)`
  
  - remove a sheet

- `save()`
  
  - save your stuff

### Sheet

- `find(**kwargs)`
  
  - find a list of rows

- `filter(callback)`
  
  - return a list of rows, filter by the callback function. return True if you want it

- `append(content: Union[dict, List[str]])`
  
  - add a row to your file, dict keys should be your headers

- `remove(row: dict)`
  
  - remove a row

- `set_header_style(style: Style)`
  
  - set style of the header

- get_col(col: str)
  
  - get rows of a column

- `set_row_style(row: Union[dict, int], style: Style)`
  
  - set style of a row

- `to_csv(out: str = '')`
  
  - Convert sheet to csv

- `to_json(out: str = '')`
  
  - Convert sheet to json

## utils Module

### Document

- `split_sheets(out='', out_prefix='')`
  
  - split file by sheets

- `split_rows(row_count: int, out='', out_prefix='')`
  
  - split file by rows
