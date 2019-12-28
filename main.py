import xlrd
import os
def writetotxt(works):
    f = open('结果.txt', 'w')
    for key in works.keys():
        f.write('%s共工作了%s次\n' % (key, works[key]))
# 通过数据表来获取‘内业人员’所在的行列号
def findNeiYeiRenYuan(worksheet):
    for x in range(worksheet.nrows):
        for y in range(worksheet.ncols):
            if worksheet.cell_value(x, y) == '内业人员':
                return x, y

works = {}  # {人名，工作次数}{key,value}
# os.walk 返回一个数组
for root, dirs, files in os.walk('E:\\EXCEL数据文件'):
    for file in files:
        filename = os.path.join(root, file)
        workbook = xlrd.open_workbook(filename)
        worksheet = workbook.sheet_by_index(0)
        #获得内业人员所在的行列
        x, y = findNeiYeiRenYuan(worksheet)

        for i in range(x + 1, worksheet.nrows):
            name = worksheet.cell_value(i, y)
            if name != '' and name != '内业人员':
                if name in works.keys():
                    works[name] = works[name] + 1
                else:
                    works[name] = 1
writetotxt(works)
