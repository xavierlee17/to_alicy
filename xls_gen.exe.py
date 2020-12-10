#!/usr/bin/env python3

__author__ = 'XavierStudio'

# -*- coding: UTF-8 -*-
import xlrd
import xlwt
import datetime
import sys
import os
import time

def xls_gen(filename,outputfile):
    data_excel = xlrd.open_workbook(filename)
    names = data_excel.sheet_names()
    table = data_excel.sheets()[0]
    n_rows = table.nrows
    n_cols = table.ncols
    data = []
    DATA = []
    for y in range(n_rows):
        for x in range(n_cols):
            cell_data = table.cell_value(rowx=y,colx=x)
            data.append("%s" % cell_data)
        DATA.append(data)
        data = []
    
    name = DATA[0][0]
    price= DATA[0][1]
    errrate= DATA[0][2]
    the_range=[]
    for one in range(5,len(DATA[0])):
        the_range.append(DATA[0][one])
    
    i = 0
    report = []
    data = []
    for one in DATA:
        if i != 0:
            postage = one[3]
            handling= one[4]
            data.append(one[0])
            data.append(one[1])
            data.append(one[2])
            data.append(float(one[3]))
            data.append(float(one[4]))
            for one_range in the_range:
                result = (((float(one_range) * float(postage)) + float(handling)) + float(price)) * (1 + float(errrate))
                data.append(result)
            report.append(data)
            data = []
        else:
            report.append(one)
        i=i+1
    
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet= workbook.add_sheet('Sheet1')
    x = 0
    y = 0
    for one_line in report:
        for one_item in one_line:
            worksheet.write(y,x,str(one_item))
            x = x+1
        y = y+1
        x = 0
    
    workbook.save(outputfile)

YYYY = str(datetime.date.today().year)
MM = str(datetime.date.today().month)
DD = str(datetime.date.today().day)
i = 0
for one in sys.argv:
    if i != 0:
        if os.path.isfile(one):
            filename = one
            tmp = os.path.splitext(one)
            if tmp[1] == '.xls' or tmp[1] == '.xlsx':
                outputfile=tmp[0]+'.out.'+YYYY+MM+DD+tmp[1]
                xls_gen(filename,outputfile)
    i = i+1
