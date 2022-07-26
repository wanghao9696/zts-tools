import openpyxl


relate_path = "D:/database/营业部统计/分支机构隶属关系.xlsx"
# relate_path = "D:/database/营业部统计/待开北京市场账户数量统计_112-120.xlsx"
relate_excel = openpyxl.load_workbook(relate_path)
relate_sheet = relate_excel[relate_excel.sheetnames[0]]
relate_dict = {}

for i in range(1, len(relate_sheet['A'])):
    relate_dict[relate_sheet['A' + str(i)].value] = relate_sheet['B' + str(i)].value

# print(relate_dict.keys())
for key,value in relate_dict.items():
    print(key)