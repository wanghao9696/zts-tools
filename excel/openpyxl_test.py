import openpyxl
import cx_Oracle
import yaml
import os

import warnings
warnings.filterwarnings('ignore')


# 读取config.yaml配置文件
curPath = os.path.dirname(os.path.realpath(__file__))
yamlPath = os.path.join(curPath, "config.yaml")
f = open(yamlPath, 'r', encoding='utf-8')
cfg = f.read()
cfg_data = yaml.safe_load(cfg)


# 存放sql语句
querysql = []
querysql.append([])
querysql[0].append(cfg_data["table1"]["querysql1"])
querysql[0].append(cfg_data["table1"]["querysql2"])
querysql.append([])
querysql[1].append(cfg_data["table2"]["querysql1"])
querysql[1].append(cfg_data["table2"]["querysql2"])
querysql[1].append(cfg_data["table2"]["querysql3"])
querysql.append([])
querysql[2].append(cfg_data["table3"]["querysql1"])
querysql[2].append(cfg_data["table3"]["querysql2"])
querysql.append([])
querysql[3].append(cfg_data["table4"]["querysql1"])
querysql[3].append(cfg_data["table4"]["querysql2"])
querysql[3].append(cfg_data["table4"]["querysql3"])


dbhandle = cx_Oracle.connect('username', 'pwd', 'tns')
dbcursor = dbhandle.cursor()

# 获取来自数据库的数据
data_from_ocrale = []
for i in range(len(querysql)):
    data_from_ocrale.append([])
    for j in range(len(querysql[i])):
        order = querysql[i][j]
        dbcursor.execute(order)
        data = dbcursor.fetchone()
        data_from_ocrale[i].append(data)


# 打印数据库获取的数据
for i in range(len(data_from_ocrale)):
    print("表格" + str(i) + "的数据如下：\n")
    for j in range(len(data_from_ocrale[i])):
        print(data_from_ocrale[i][j])
        print("\n")


# data_from_ocrale = [["数据库拉取表1数据1", "数据库拉取表1数据2"],
#                    ["数据库拉取表2数据1", "数据库拉取表2数据2", "数据库拉取表2数据3"],
#                    ["数据库拉取表3数据1", "数据库拉取表3数据2"],
#                    ["数据库拉取表5数据1", "数据库拉取表5数据2", "数据库拉取表5数据3"]]


indexs = 'ABCD'
indexs_2 = 'FGH'
filename_ori = cfg_data["filepathi"]["ori"]
filename_rel = cfg_data["filepathi"]["rel"]

excel = openpyxl.load_workbook(filename_ori)
sheet_list = excel.sheetnames
sheet1 = excel[sheet_list[0]]


# 获取日期
now_year = cfg_data["date"]["year"]
now_month = cfg_data["date"]["month"]
if now_month >= 6:
    start_month = now_month - 5
    last_month = now_month - 1
    start_year = now_year
    last_year = now_year
else:
    start_month = 12 - (5 - now_month)
    start_year = now_year - 1
    if now_month == 1:
        last_month = 12
        last_year = now_year - 1
    else:
        last_month = now_month - 1
        last_year = now_year


# 表头
title = "中泰证券" + str(start_year) + "年" + str(start_month) + "月-"  + str(now_year) + "年" + str(now_month) + "月"
table_title = ["客户开户统计表", "客户销户统计表",
               "累计客户号数量表", "期末持仓投资者占比"]


row_table = []
for i in range(len(sheet1['A'])):
    # 获取每个表格的行数位置
    if type(sheet1['A'][i].value) is type(""):
        if '开户统计表' in sheet1['A'][i].value:
            row_table.append(i + 1)
        if '销户统计表' in sheet1['A'][i].value:
            row_table.append(i + 1)
        if '客户号数量表' in sheet1['A'][i].value:
            row_table.append(i + 1)
        if '持仓投资者占比表' in sheet1['A'][i].value:
            row_table.append(i + 1)

for i in range(len(sheet1['E'])):
    if type(sheet1['E'][i].value) is type(""):
        if '休眠数量统计' in sheet1['E'][i].value:
            row_table.append(i + 1)
print(row_table)


# 校对
try:
    str(last_year) + "年" + str(last_month) + "月" in sheet1['A1']
except IOError:
    print("Error: 该月日期与上月表格日期不一致")


def copy_rows(sheet, n, m):
    # m行的值复制到n行
    for i in indexs:
        sheet[i + str(n)] = sheet[i + str(m)].value


for i in range(4):
    '''更新4个表格的数据'''
    # 更新表格的标题
    sheet1['A' + str(row_table[i])] = title + table_title[i]

    # 更新前5个月份数据
    for j in range(2, 7):
        copy_rows(sheet1, row_table[i]+j, row_table[i]+j+1)

    # 更新最新一个月份数据
    sheet1['A' + str(row_table[i]+7)] = str(now_year) + '年' + str(now_month) + '月'
    sheet1['A' + str(row_table[i]+7)].number_format = 'yyyy年mm月'

    if i != 3:
        for j in range(len(data_from_ocrale[i])):
            sheet1[indexs[j+1] + str(row_table[i]+7)] = data_from_ocrale[i][j]

# 更新休眠数据统计表格
for i in range(len(indexs_2) - 1):
    sheet1[indexs_2[i] + str(row_table[4]+1)] = sheet1[indexs_2[i+1] + str(row_table[4]+1)].value
    sheet1[indexs_2[i] + str(row_table[4]+2)] = sheet1[indexs_2[i+1] + str(row_table[4]+2)].value
    sheet1[indexs_2[i] + str(row_table[4]+3)] = sheet1[indexs_2[i+1] + str(row_table[4]+3)].value
    sheet1[indexs_2[i] + str(row_table[4]+5)] = sheet1[indexs_2[i+1] + str(row_table[4]+5)].value
sheet1[indexs_2[2] + str(row_table[4]+1)] = str(now_month) + "月份"
sheet1[indexs_2[2] + str(row_table[4]+2)] = data_from_ocrale[3][0]
sheet1[indexs_2[2] + str(row_table[4]+3)] = data_from_ocrale[3][1]
sheet1[indexs_2[2] + str(row_table[4]+5)] = data_from_ocrale[3][2]

# 更新表格4的数据
sheet1['B' + str(row_table[3]+7)] = sheet1['H110'].value / sheet1['H111'].value * 100
sheet1['B' + str(row_table[3]+7)].number_format = '##.00'

excel.save(filename_rel)