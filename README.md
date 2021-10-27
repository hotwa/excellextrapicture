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
