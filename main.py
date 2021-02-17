from docx import Document
import xlwt
import docx
document = Document('123.docx')
cnt = 0

import xlrd
data = xlrd.open_workbook("22.xlsx")
table = data.sheets()[0]
nrows=table.nrows
ncols=table.ncols
ksh = []
xm = []
zkz = []
kddm = []
kch = []
zwh = []
for i in range(1 , nrows):
    ksh.append(str(table.row(i)[0].value)[:-2])
    xm.append(str(table.row(i)[1].value))
    zkz.append(str(table.row(i)[2].value)[:-2])
    kddm.append(str(table.row(i)[3].value)[:-2])
    kch.append(str(table.row(i)[5].value)[:-2])
    zwh.append(str(table.row(i)[6].value)[:-2])
print(ksh)
try :
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if "张欢鑫" in run.text:
                run.text=run.text.replace('张欢鑫',xm[cnt])
            if "库" in run.text:
                run.text=run.text.replace('库',ksh[cnt])
            if "炫" in run.text:
                run.text = run.text.replace('炫', zkz[cnt])
            if "七" in run.text:
                run.text=run.text.replace('七',kch[cnt])
            if "五" in run.text:
                run.text=run.text.replace('五',zwh[cnt])
                cnt = cnt+1
except :
    document.save("123456.docx")