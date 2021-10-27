# excellextrapicture
from xlsx extract figures

[中文](https://blog.csdn.net/mm644706215/article/details/120993901)

## Quick Start

```python
x = xlsx_pic('./test.xlsx')
x.sheetnum = 0 # set activate sheet
lst = x.read_row(5,read_cell_picture=True)
cell_content = x.read_cell('E2')
```

## save figure

```python
xlsx_pic.get_cell_pic(filename='./test.xlsx',sheetnum=0,position='E2',new_name='savepic.png',base64c=None)
```

## return base64

```python
base64_content = xlsx_pic.get_cell_pic(filename='./test.xlsx',sheetnum=0,position='E2',base64c=True)
```
