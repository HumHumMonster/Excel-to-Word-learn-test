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
print(xm)
print(zkz)
print(kddm)
