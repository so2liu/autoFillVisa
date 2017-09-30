# -*- coding: utf-8 -*-
import  xdrlib ,sys
import xlrd
from pypinyin import pinyin, lazy_pinyin
import pypinyin

def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except IOError:
        print ("文件错误")
#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file,colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):

         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
def excel_table_byname(file,colnameindex=0,by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows #行数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):
         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list

def gettable_excel(filename,k):

    # tables = excel_table_byindex(filename)
    # for row in tables:
    #     print (row)

    tables = excel_table_byname(filename)
    return tables[k]

def datecn2datade(date_ch):
    if date_ch.find('-')>-1:
        date_de=date_ch[8:]+'.'+date_ch[5:7]+'.'+date_ch[0:4]
    else:date_de=date_ch
    return date_de

def is_chinese(uchar):
    if '\u4e00' <= uchar<='\u9fff':
        return True
    else:
        return False

def ch2de(inputs):
    if is_chinese(inputs):
        inputs=pypinyin.lazy_pinyin(inputs)
        inputs = ''.join(inputs)

    return inputs.capitalize()
