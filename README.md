# zts-tools

#### 打开excel文件
```wb = openpyxl.load_workbook()```

#### sheet名称list
```excel.sheetnames

#### sheet中图表的list
```sheet._charts```

#### 更改chart数据源指向
```
def get_datas_and_cats(sheet, row_table):
    datas = []
    cats = []
    datas.append(Reference(worksheet=sheet, min_row=row_table[1]+1, max_row=row_table[1]+7, min_col=2, max_col=4))
    cats.append(Reference(worksheet=sheet, min_row=row_table[1]+2, max_row=row_table[1]+7, min_col=1, max_col=1))

    datas.append(Reference(worksheet=sheet, min_row=row_table[3]+1, max_row=row_table[3]+7, min_col=2, max_col=2))
    cats.append(Reference(worksheet=sheet, min_row=row_table[3]+2, max_row=row_table[3]+7, min_col=1, max_col=1))

    datas.append(Reference(worksheet=sheet, min_row=row_table[2]+1, max_row=row_table[2]+7, min_col=2, max_col=3))
    cats.append(Reference(worksheet=sheet, min_row=row_table[2]+2, max_row=row_table[2]+7, min_col=1, max_col=1))

    datas.append(Reference(worksheet=sheet, min_row=row_table[0]+1, max_row=row_table[0]+7, min_col=2, max_col=3))
    cats.append(Reference(worksheet=sheet, min_row=row_table[0]+2, max_row=row_table[0]+7, min_col=1, max_col=1))

    return datas, cats


charts = []
datas, cats = get_datas_and_cats(new_sheet, row_table)

for chart in new_sheet._charts:
    chart.ser = []
    charts.append(chart)

for i in range(len(charts)):
    charts[i].add_data(datas[i], titles_from_data=True)
    charts[i].set_categories(cats[i])
    for j in range(len(charts[i].ser)):
        charts[i].ser[j].marker = openpyxl.chart.marker.Marker(symbol=mark[j], size=7)
        charts[i].ser[j].graphicalProperties.line.width = 30000
        charts[i].ser[j].graphicalProperties.line.solidFill = color[j]
        charts[i].ser[j].marker.graphicalProperties.solidFill = color[j]
    charts[i].dLbls.showVal = True
```