import openpyxl
import cx_Oracle
import yaml
import os
import threading

import warnings
warnings.filterwarnings('ignore')


# 读取config.yaml配置文件
def read_config():
    curPath = os.path.dirname(os.path.realpath(__file__))
    yamlPath = os.path.join(curPath, "config.yaml")
    f = open(yamlPath, 'r', encoding='utf-8')
    cfg = f.read()
    cfg_data = yaml.safe_load(cfg)

    return cfg_data


# 获取sql语句
def get_querysql(cfg_data):
    querysql = [[], [], [], []]
    querysql[0].append(cfg_data["table1"]["querysql1"])
    querysql[0].append(cfg_data["table1"]["querysql2"])

    querysql[1].append(cfg_data["table2"]["querysql1"])
    querysql[1].append(cfg_data["table2"]["querysql2"])
    querysql[1].append(cfg_data["table2"]["querysql3"])

    querysql[2].append(cfg_data["table3"]["querysql1"])
    querysql[2].append(cfg_data["table3"]["querysql2"])

    querysql[3].append(cfg_data["table4"]["querysql1"])
    querysql[3].append(cfg_data["table4"]["querysql2"])
    querysql[3].append(cfg_data["table4"]["querysql3"])

    return querysql


# sql语句中字符串替换
def sql_time(querysql, cur_year, cur_month):
    begin_time = str(cur_year) + str(cur_month).zfill(2) + "01"
    end_time = str(cur_year) + str(cur_month).zfill(2) + "31"
    print("查询起始时间：", begin_time)
    print("查询截止时间：", end_time)
    for i in range(len(querysql)):
        for j in range(len(querysql[i])):
            querysql[i][j] = querysql[i][j].replace("begin_time", begin_time)
            querysql[i][j] = querysql[i][j].replace("end_time", end_time)
            print(querysql[i][j])

    return querysql

# 数据库拉取数据
def ocrale_process(dbcursor, order_list, data_list):
    for i in range(len(order_list)):
        dbcursor.execute(order_list[i])
        data = dbcursor.fetchone()[0]
        data_list.append(data)


# 打印数据库获取的数据
def print_ocrale_data(data_from_ocrale):
    for i in range(len(data_from_ocrale)):
        print("表格" + str(i) + "的数据如下：")
        for j in range(len(data_from_ocrale[i])):
            print(data_from_ocrale[i][j])


# 获取每个表格的行数位置
def get_row_table(new_sheet):
    row_table = []
    for i in range(len(new_sheet['A'])):
        if type(new_sheet['A'][i].value) is type(""):
            if '开户统计表' in new_sheet['A'][i].value:
                row_table.append(i + 1)
            if '销户统计表' in new_sheet['A'][i].value:
                row_table.append(i + 1)
            if '客户号数量表' in new_sheet['A'][i].value:
                row_table.append(i + 1)
            if '持仓投资者占比表' in new_sheet['A'][i].value:
                row_table.append(i + 1)

    for i in range(len(new_sheet['E'])):
        if type(new_sheet['E'][i].value) is type(""):
            if '休眠数量统计' in new_sheet['E'][i].value:
                row_table.append(i + 1)

    return row_table


# m行的值复制到n行
def copy_rows(sheet, n, m):
    indexs = 'ABCD'
    for i in indexs:
        sheet[i + str(n)] = sheet[i + str(m)].value


# 更新excel表格数据
def update_table(new_sheet, row_table, cur_year, cur_month, data_from_ocrale):
    indexs = 'ABCD'
    indexs_2 = 'FGH'
    for i in range(4):
        # 更新表格的标题
        new_sheet['A' + str(row_table[i])] = title1 + table_title[i]

        # 更新前5个月份数据
        for j in range(2, 7):
            copy_rows(new_sheet, row_table[i] + j, row_table[i] + j + 1)

        # 更新月份
        new_sheet['A' + str(row_table[i] + 7)] = str(cur_year) + '年' + str(cur_month) + '月'
        new_sheet['A' + str(row_table[i] + 7)].number_format = 'yyyy年mm月'

        # 更新前3个表格（第4个表格需计算）
        if i != 3:
            for j in range(len(data_from_ocrale[i])):
                if i == 2:
                    new_sheet[indexs[j + 1] + str(row_table[i] + 7)] = data_from_ocrale[i][j] / 10000
                else:
                    new_sheet[indexs[j + 1] + str(row_table[i] + 7)] = data_from_ocrale[i][j]
    print("前三个表格数据更新完成！！！！")

    # 更新休眠数据统计表格
    for i in range(len(indexs_2) - 1):
        new_sheet[indexs_2[i] + str(row_table[4] + 1)] = new_sheet[indexs_2[i + 1] + str(row_table[4] + 1)].value
        new_sheet[indexs_2[i] + str(row_table[4] + 2)] = new_sheet[indexs_2[i + 1] + str(row_table[4] + 2)].value
        new_sheet[indexs_2[i] + str(row_table[4] + 3)] = new_sheet[indexs_2[i + 1] + str(row_table[4] + 3)].value
        new_sheet[indexs_2[i] + str(row_table[4] + 5)] = new_sheet[indexs_2[i + 1] + str(row_table[4] + 5)].value
    new_sheet[indexs_2[2] + str(row_table[4] + 1)] = str(cur_month) + "月份"
    new_sheet[indexs_2[2] + str(row_table[4] + 2)] = data_from_ocrale[3][0]
    new_sheet[indexs_2[2] + str(row_table[4] + 3)] = data_from_ocrale[3][1]
    new_sheet[indexs_2[2] + str(row_table[4] + 5)] = data_from_ocrale[3][2]
    print("休眠数据统计表格更新完成！！！！")

    # 更新表格4的数据
    new_sheet['B' + str(row_table[3] + 7)] = new_sheet['H' + str(row_table[4] + 2)].value / new_sheet['H' + str(row_table[4] + 3)].value * 100
    new_sheet['B' + str(row_table[3] + 7)].number_format = '##.00'
    print("表格4数据更新完成！！！！")


if __name__ == "__main__":
    cfg_data = read_config()

    filename_ori = cfg_data["filepath"]["ori"]
    filename_rel = cfg_data["filepath"]["rel"]

    excel = openpyxl.load_workbook(filename_ori)
    sheet_list = excel.sheetnames
    new_sheet = excel[sheet_list[-1]]
    new_sheet_name = excel[sheet_list[-1]].title

    # 获取日期
    cur_year = int(new_sheet_name[0:4])
    cur_month = int(new_sheet_name[4:6])
    sta_year = cur_year if cur_month >= 6 else cur_year - 1
    sta_month = cur_month - 5 if cur_month >= 6 else (12 - 5 + cur_month)

    # 表头
    title1 = "中泰证券" + str(sta_year) + "年" + str(sta_month) + "月-" + str(cur_year) + "年" + str(cur_month) + "月"
    table_title = ["客户开户统计表", "客户销户统计表",
                   "累计客户号数量表", "期末持仓投资者占比"]

    querysql = get_querysql(cfg_data)  # 获取sql语句
    querysql = sql_time(querysql, cur_year, cur_month)  # sql语句起止时间替换

    data_from_ocrale = [[], [], [], []]

    # 连接数据库并获取数据
    # dbhandle = cx_Oracle.connect('ql_read', 'ql_read', '10.29.180.151:2521/fzqsxt')
    # dbcursor = dbhandle.cursor()

    # threads = []
    # t1 = threading.Thread(target=ocrale_process,args=(dbcursor, querysql[0], data_from_ocrale[0]))
    # threads.append(t1)
    # t2 = threading.Thread(target=ocrale_process,args=(dbcursor, querysql[1], data_from_ocrale[1]))
    # threads.append(t2)
    # t3 = threading.Thread(target=ocrale_process,args=(dbcursor, querysql[2], data_from_ocrale[2]))
    # threads.append(t3)
    # t4 = threading.Thread(target=ocrale_process,args=(dbcursor, querysql[3], data_from_ocrale[3]))
    # threads.append(t4)

    # for t in threads:
    #     # t.setDaemon(True)
    #     t.start()
    # for thread in threads:
    #     thread.join()

    data_from_ocrale[0].append(57237)
    data_from_ocrale[0].append(52560)
    data_from_ocrale[1].append(2021)
    data_from_ocrale[1].append(3284)
    data_from_ocrale[1].append(2917)
    data_from_ocrale[2].append(7603607)
    data_from_ocrale[2].append(6215629)
    data_from_ocrale[3].append(1870248)
    data_from_ocrale[3].append(4565310)
    data_from_ocrale[3].append(6509292)

    row_table = get_row_table(new_sheet)
    print("5个表格的起始行数：", row_table)

    print_ocrale_data(data_from_ocrale)  # 打印数据库获取的数据

    update_table(new_sheet, row_table, cur_year, cur_month, data_from_ocrale)  # 更新

    print("done!!!!!")

    excel.save("testt.xlsx")