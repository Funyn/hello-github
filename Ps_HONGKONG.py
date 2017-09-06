#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2017-09-05 22:13:52
# @Author  : Your Name (you@example.org)
# @Link    : http://example.org
# @Version : $Id$

import os, sys
import xlrd
import datetime

print(os.listdir())
print('+-' * 35)
print('''
    version = 1.0.0
    请输入你的文件名,请带上后缀,上面有对应的文件名列表,输入至>>> 后面即可,不明白看下面例子
    注意事项: 输入时请不要带上两边引号
    例如:['.idea', 'PSA3594.xlsx', 'Ps_HONGKONG.py'] 想要获取'PSA3594.xlsx'
        输入文件名PSA3594.xlsx
        >>>PSA3594.xlsx   输入完成后
        Enter开始执行校验程序
    ''')
print('+-' * 35)
check_ps = input('>>>')


if check_ps in os.listdir():
    excel = xlrd.open_workbook(check_ps)
else:
    raise NameError('你输入的%s不存在,请确认你输入文件名' %check_ps)

sheet1 = excel.sheet_by_index(0)
sheet1 = excel.sheet_by_name(sheet1.name)
collect = {}
for col in range(sheet1.ncols):
    for row in range(sheet1.nrows):
        point = sheet1.cell(row, col).value
        if  'Invoice Number' in str(point) :
            collect['Invoice Number'] = (row, col)
        if 'Shipping Method' in str(point) :
            collect['Shipping Method'] = (row, col)
        if  'Qty' in str(point) :
            collect['Qty'] = (row, col)
        if 'Unit Price' in str(point) :
            collect['Unit Price'] = (row, col)
        if str(point) == 'Total':
            collect['Total'] = (row, col)
        if str(point) == 'Weight (kg)' :
            collect['Weight (kg)'] = (row,col)
        if 'Total Weight (kg)' in str(point)  :
            collect['Total Weight (kg)'] = (row,col)
        if 'Date' in str(point):
            collect['Date'] = (row,col)

#-------文件名称-----#
name = sheet1.cell(collect['Invoice Number'][0], collect['Invoice Number'][1] + 2).value
def verify_filename(name,filename):
    filename = filename.split('.')
    try:
        if name == filename[0]:
            print('{}文件名校验合格'.format(name))
        else:
            raise ValueError('Invoice Number:%s 与文件名不符,请检查' %name)
    except ValueError as e:
        print(e)
#-------邮寄方式---#
shipping_times = sheet1.cell(collect['Shipping Method'][0] + 1, collect['Shipping Method'][1] + 2).value
def verify_shipping(shipping_times):
    shipping_times = shipping_times.split(' ')
    if int(shipping_times[0]) > 1:
        if shipping_times[1] == 'Cartons':
            print('邮寄数量校验合格')
        else:
            print('%s 邮寄数量校验不合格,请检查'%shipping_times[1])
    else:
        if shipping_times[1] == 'Carton':
            print('邮寄数量校验合格')
#-------时间-------#
timestamp = os.stat(os.getcwd() + '\\' + check_ps).st_mtime
modify_date = datetime.datetime.fromtimestamp(timestamp)
date = str(modify_date)[8:10] + '-' + str(modify_date)[5:7] + '-' + str(modify_date)[2:4]
in_date = sheet1.cell(collect['Date'][0], collect['Date'][1]).value
#------价格匹配-----#
total = float(sheet1.cell(collect['Total'][0], collect['Total'][1] + 1).value)
def verify_total(total):
    total_list = []
    end = int(collect['Total'][0])-int(collect['Qty'][0])
    for num in range(1,end):
        qty = sheet1.cell(collect['Qty'][0] + num, collect['Qty'][1]).value
        unitprice = sheet1.cell(collect['Qty'][0] + num, collect['Qty'][1] + 1).value
        subtotal = sheet1.cell(collect['Qty'][0] + num, collect['Qty'][1] + 2).value
        if qty == '' or unitprice == '':
            continue
        s_total = float(qty) * float(unitprice)
        if s_total == subtotal:
            total_list.append(subtotal)
            print('Qty = %s, unitprice = %s,subtotal = %s' % (qty, unitprice, subtotal))
            print('第 %s 行数据 金额校验合格 ' %(collect['Qty'][0] + num+1))
        else:
            total_list.append(subtotal)
            print('Qty = %s, unitprice = %s,subtotal = %s' % (qty, unitprice, subtotal))
            print('第 %s 行数据 金额校验不合格,请检查确认 ' %(collect['Qty'][0] + num+1))
    if sum(total_list) == total:
        print('总金额校验合格')
    else:
        print('总金额校验不合格,请检查确认')
#------重量匹配-----#
total_weight = sheet1.cell(collect['Total Weight (kg)'][0], collect['Total Weight (kg)'][1]+1).value
def verify_weight(total_weight):
    weight_list = []
    end = int(collect['Total Weight (kg)'][0]) - int(collect['Weight (kg)'][0])
    for wei in range(1, end):
        weight = sheet1.cell(collect['Weight (kg)'][0] + wei, collect['Weight (kg)'][1]).value
        if weight == '':
            continue
        weight_list.append(weight)
    if sum(weight_list) == total_weight:
        print('总重量校验合格')


def main():
    verify_filename(name,check_ps)
    print('+-' * 35)
    verify_total(total)
    print('+-' * 35)
    verify_shipping(shipping_times)
    print('+-' * 35)
    verify_weight(total_weight)
    print('+-' * 35)
    print('最后修改时间',date)
    print('文件内Date时间',in_date)

if __name__ == '__main__':
    main()











