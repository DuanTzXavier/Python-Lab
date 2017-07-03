#coding=gb2312
import xlsxwriter
import time
import xlrd

maxValue = 15224

startTime = time.time()
print "正在初始化表格数据"
data = xlrd.open_workbook('~~.xlsx') 
table = data.sheets()[0]
print "初始化表格数据成功， 用时 : " + str(time.time() - startTime) + "s"
initTime = time.time()
print "正在初始化写入文件"
workbook = xlsxwriter.Workbook('filename2.xlsx')
worksheet1 = workbook.add_worksheet()
print "初始化新写入文件成功， 用时 : "+ str(time.time() - initTime) + "s"
print "开始写入数据"
i = 0
sum =  table.col_values(0)
for value in table.col_values(1):
    singleTime = time.time()
    isfound = False
    if i < maxValue:
        for lit in sum:
            if (value + " - ") in lit:
                worksheet1.write('A' + str(i), value)
                bValue = sum[sum.index(lit) - 1]
                bValue = bValue.replace("'", '')
                bValue = bValue.replace(':', '')
                worksheet1.write('B' + str(i), bValue) 
                print "完成" + str(i * 100 / maxValue) + "% 正在查找第" + str(i) + "个，已用时： " + str(time.time() - startTime) + 's\r',
                isfound = True
                break
    else:
        print "------------------------------我是分割线---------------------------------" 
        break
    if not isfound:
        print "第" + str(i) + "个用时: "+ str(time.time() - singleTime) + "s" + "没有找到"
    i = i + 1

workbook.close()

print "已完成，用时： " + str(time.time() - startTime) + 's',
