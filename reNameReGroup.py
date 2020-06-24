# 文件批量改名与分组
import os
import xlrd
import shutil

opath = r'XXXXX'   # 待修改文件目录

fileList = os.listdir(opath)
fileList.sort(key=lambda x: int(x.split('-', 1)[0]))  # 提取现有文件名
# print(fileList)
data = xlrd.open_workbook(r'xxxx.xlsx')
table = data.sheet_by_name('All_Data_Readable')
# print(table.nrows)  # 56

name = table.col_values(9)[1:]  # 姓名
paper = table.col_values(13)[1:]  # 论文题目
direction = table.col_values(15)[1:]  # 分论坛名
form = table.col_values(17)[1:]  # 参会形式
# print(direction)
# for i, fs in enumerate(fileList):
#     # if i == 37:
#     #     continue
#     tp = fs.split('.')[-1]  # 扩展名
#     pName = opath + os.sep + fs
#     # 编号-姓名-分论坛-论文题目-参会形式
#     nName = opath + os.sep + str(i+1) + '-' + name[i] + '-' + direction[i] + '-' + paper[i] + '-' + form[i] + '.' + tp
#     # 注意：文件名中不能包含’/‘等字符，否则系统不能识别，报错
#     os.rename(pName, nName)

# 移动分组
for fs in fileList:
    dire = fs.split('-')[2]
    newDir = os.path.join(opath, dire)
    if not os.path.exists(newDir):
        os.makedirs(newDir)
    old = os.path.join(opath, fs)
    new = os.path.join(newDir, fs)
    shutil.move(old, new)
