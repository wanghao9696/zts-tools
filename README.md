# zts-tools

#### 打开excel文件
```wb = openpyxl.load_workbook()```

#### sheet名称list
```excel.sheetnames```

#### sheet中图表的list
```sheet._charts```

#### 更改chart数据源指向
```
import openpyxl
from openpyxl.chart import Reference

mark = ['triangle', 'circle', 'diamond']
color = ["FF0000", "0097FF", "6E0000"]

data = Reference(worksheet=sheet, min_row=row_table[1]+1, max_row=row_table[1]+7, min_col=2, max_col=4) # 数据源
cat = Reference(worksheet=sheet, min_row=row_table[1]+2, max_row=row_table[1]+7, min_col=1, max_col=1) # 类别源

chart.ser = [] # 清空图表数据

chart.add_data(data, titles_from_data=True) # 重设数据
charts.set_categories(cat) # 重设类别

chart.ser.marker = openpyxl.chart.marker.Marker(symbol=mark[j], size=7) # 标记
chart.ser.graphicalProperties.line.width = 30000 # 连线
chart.ser.graphicalProperties.line.solidFill = color[j] # 连线颜色
chart.ser.marker.graphicalProperties.solidFill = color[j] # 标记颜色
chart.dLbls.showVal = True # 显示数据
```