#coding=utf-8
import time
import xlrd
import xlsxwriter

def getProduct(item):
    if item in producItems :
        orderIndex = producItems.index(item)
        global productQuantity,productAmount
        productQuantity = productQuantity + int(productNum[orderIndex])
        productAmount =productAmount + float(productPrice[orderIndex])
        producItems.remove(item)
        getProduct(item)


def writeProduct(item, index):
    getProduct(item)
    global productQuantity, productAmount
    if productQuantity > 0:
        global indexLine,worksheet1
        indexLine = indexLine + 1
        worksheet1.write('A' + str(indexLine), item)
        worksheet1.write('B' + str(indexLine), getProductName(index))
        worksheet1.write('C' + str(indexLine), str(productAmount/productQuantity))
        worksheet1.write('D' + str(indexLine), productQuantity)
        worksheet1.write('E' + str(indexLine), productAmount)


def getProductName(index):
    global productName
    
    return productName[index]


startTime  = time.time()
data = xlrd.open_workbook('../excel/2017-07-01-2.xlsx')
productData = xlrd.open_workbook('../excel/all_products.xlsx')
workbook = xlsxwriter.Workbook('../excel/生成表.xlsx')
worksheet1 = workbook.add_worksheet(u'商品销售清单')

sheetNum = len(productData.sheet_names())

productQuantity = 0
productAmount = 0.0
indexLine = 1
table = data.sheets()[0]
producItems =  table.col_values(11)
productPrice =  table.col_values(16)
productNum =  table.col_values(14)


for i in range(0, sheetNum - 1):
    productTable = productData.sheets()[i]
    productName = productTable.col_values(4)
    k = 0
    for productItem in productTable.col_values(3):
        if (k > 1):
            productQuantity = 0
            productAmount = 0.0
            writeProduct(productItem, k)
        k = k + 1

worksheet1.write('A1', u'商品编号')
worksheet1.write('B1', u'商品名称')
worksheet1.write('C1', u'商品平均结算价')
worksheet1.write('D1', u'商品销售数量')
worksheet1.write('E1', u'商品结算总额')



workbook.close()
print 