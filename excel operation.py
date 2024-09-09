'''
Description: 
Author: CrayonXiaoxin
Date: 2024-05-31 15:26:04
LastEditTime: 2024-09-09 14:29:23
LastEditors:  
'''
import os
import openpyxl
# open file
# path是文件所在位置
path = r"C:\\Users\\Tomato\\OneDrive\\桌面"
os.chdir(path)
# 括号里面填文件名
workbook = openpyxl.load_workbook("2.xlsx")
print(workbook.sheetnames)

# get Sheet1 information
sheet = workbook['1']
rows = sheet.max_row
column = sheet.max_column

# get top name in Sheet1
topName = []
for i in range(1,column+1):
    topName.append(sheet.cell(1,i).value)


# 如果想要将Sheet1中的数据按照第二列的值进行分类，并分别保存到不同的sheet中，可以按照以下步骤进行：
# names = []
# for i in range(1, rows+1):
#     names.append(sheet.cell(i, 2).value)
# names = list(set(names))
# 如果想要自己制定名字,可以修改names列表
names = ['苏秦海', '林新星', '雷霆']

for name in names:
    workbook.create_sheet(name+'')

    
# workbook.save('test.xlsx')

# get all values in Sheet1 ans save
for i in range(1,rows+1):
    name = sheet.cell(i,2).value
    if name in names:
        viceSheet = workbook[name+'']
        values = []
        for j in range(1,column+1):
            values.append(sheet.cell(i,j).value)
        vicerows = viceSheet.max_row
        for k in range(1,column+1):
            viceSheet.cell(vicerows+1,k,values[k-1])
        for l in range(1,column+1):
            viceSheet.cell(1,l,topName[l-1])
workbook.save('2.xlsx')
