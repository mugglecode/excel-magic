from excelMagic.Document import create_document
from excelMagic.dataset import Dataset
from excelMagic.utils import Document


def callback(content):
    return content['name'] == '小明'

if __name__ == '__main__':
    doc = create_document('new.xlsx')
    doc.append('add1.xls')
    doc.append('add1.xls')
    doc.close()

    doc = Document('add1.xls')
    doc.split_sheets(out_prefix='fuck_')

    doc = Document('split.xls')
    doc.split_rows(5, out_prefix='split_row_')
    html = doc.to_html()
    f = open('html.html', 'w')
    f.write(html)
    f.flush()
    f.close()

    ds = Dataset('filter_test.xls')
    result = ds.filter(callback)
    for r in result:
        print(r)
