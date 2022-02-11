# zts-tools

<img src="https://github.com/wanghao9696/zts-tools/blob/main/excel/test.png"   width="20%">

#### 更改chart数据源指向
```
import openpyxl
from openpyxl.chart import Reference

wb = openpyxl.load_workbook("") # 读取excel文件
sheet = wb[wb.sheetnames[-1]] # 读取最后一个sheet
chart = sheet._charts[0] # 读取sheet中的一个图表

chart.ser = [] # 清空图表数据

data = Reference(worksheet=sheet, min_row=2, max_row=8, min_col=2, max_col=4) # 数据源
cat = Reference(worksheet=sheet, min_row=3, max_row=row_table8, min_col=1, max_col=1) # 类别源

chart.add_data(data, titles_from_data=True) # 重设数据
chart.set_categories(cat) # 重设类别

chart.ser.marker = openpyxl.chart.marker.Marker(symbol='triangle', size=7) # 标记
chart.ser.graphicalProperties.line.width = 30000 # 连线
chart.ser.graphicalProperties.line.solidFill = "FF0000" # 连线颜色
chart.ser.marker.graphicalProperties.solidFill = "FF0000" # 标记颜色
chart.dLbls.showVal = True # 显示数据
```