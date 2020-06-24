# 将生成的word文件转换为pdf格式
from win32com.client import Dispatch, constants, gencache
import os
import xlrd
from wordGen import genWord

# w = Dispatch('Word.Application')   # 网上大多是这个,但会导致constants的属性无法使用
w = gencache.EnsureDispatch('Word.Application')
w.Visible = 0
w.DisplayAlerts = 0
source = r'xxx'   # 源，word目录
sink = r'xxx'     # 汇，pdf目录
# 输入收件人姓名，用于文件名与文件读取
def genPDF(name):
    wordName = os.path.join(source, name+'.docx')  # 根据word的生成规则来的
    pdfName = '录用通知' + '-' + name + '.pdf'
    outPath = os.path.join(sink, pdfName)

    doc = w.Documents.Open(wordName, ReadOnly=1)
    doc.ExportAsFixedFormat(outPath, constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    # doc.SaveAs(outPath, FileFormat=17)
    doc.Close()

genPDF('XX')


# 读入参会者信息
data = xlrd.open_workbook(r'xxxx.xlsx')
table = data.sheet_by_name('info')

name = table.col_values(0)[1:]   # 姓名
paper = table.col_values(4)[1:]  # 论文题目
direction = table.col_values(6)[1:]  # 分论坛名
form = table.col_values(8)[1:]   # 参会形式

# 生成word
# for pName, pap in zip(name, paper):
#     genWord(pName, pap)

# PDF转换
for pName in name:
    genPDF(pName)

w.Quit()
