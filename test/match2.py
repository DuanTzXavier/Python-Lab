#coding=gb2312
import xlsxwriter
import time
import xlrd

maxValue = 15224

startTime = time.time()
print "���ڳ�ʼ���������"
data = xlrd.open_workbook('~~.xlsx') 
table = data.sheets()[0]
print "��ʼ��������ݳɹ��� ��ʱ : " + str(time.time() - startTime) + "s"
initTime = time.time()
print "���ڳ�ʼ��д���ļ�"
workbook = xlsxwriter.Workbook('filename2.xlsx')
worksheet1 = workbook.add_worksheet()
print "��ʼ����д���ļ��ɹ��� ��ʱ : "+ str(time.time() - initTime) + "s"
print "��ʼд������"
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
                print "���" + str(i * 100 / maxValue) + "% ���ڲ��ҵ�" + str(i) + "��������ʱ�� " + str(time.time() - startTime) + 's\r',
                isfound = True
                break
    else:
        print "------------------------------���Ƿָ���---------------------------------" 
        break
    if not isfound:
        print "��" + str(i) + "����ʱ: "+ str(time.time() - singleTime) + "s" + "û���ҵ�"
    i = i + 1

workbook.close()

print "����ɣ���ʱ�� " + str(time.time() - startTime) + 's',
