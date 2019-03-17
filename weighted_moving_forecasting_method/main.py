# -*- coding: utf-8 -*-
# @Date     : 2019-03-08 00:10:00
# @Author   : Doratree (doratree@aliyun.com)
# @Language : Python3.7

import xlrd,xlwt

def printIntro():
    print("这个程序可用于加权滑动预测法")
    print("请将相关数据及权重按要求填写到填写到input.xlsx")
    print('''请注意：权值由小到大，从左到右依次填写在第一行;
重新运行文件请关闭excel的文件，防止文件被其他应用占用。''')
    input("请输入任意字符继续")
    print("-" * 30)

def datainput():    #导入数据
    path = 'input.xlsx'
    rb = xlrd.open_workbook(path)
    data_sheet = rb.sheets()[0]
    rowNum = data_sheet.nrows
    m = rowNum - 2    #m为数据的个数
    data = []
    for i in range(2, rowNum):
        data.append(data_sheet.cell_value(i, 0))
    colNum = data_sheet.ncols
    n = colNum - 1    #n为权值的个数
    weights = []
    for j in range(1, colNum):
        weights.append(data_sheet.cell_value(0, j))
    print("数据导入成功")
    print("导入的实际值为：",data)
    print("导入的权值为：", weights)
    print("正在进行n={n}的加权滑动预测计算...".format(n=n))
    print("-"*30)
    return data, weights, n, m

def action(data, weights, n, m):
    forecast = []
    dvalue = []
    for i in range(n,m+1):    #求预测值
        y = 0
        a = 0
        for j in range(n):
            y += weights[j]*data[i-n+a]
            a +=1
        y = round(int(y*1000)/1000,2)
        forecast.append(y)
    for i in range(m-n):    #求绝对误差值
        x = abs(data[i+n]-forecast[i])
        x = round(int(x * 1000) / 1000, 2)
        dvalue.append(x)
    s = 0
    for i in range(m-n):    #求平均绝对误差值
        s +=dvalue[i]
    average = s/(m-n)
    average = round(int(average * 1000) / 1000, 2)
    return forecast, dvalue, average


def print_Summary(forecast, dvalue, average, n):    #输出结果并保存在表格
    print("从第{n}个时期起，预测值为".format(n=n+1), forecast)
    print("从第{n}个时期起，绝对误差值为".format(n=n+1), dvalue)
    print("其平均绝对误差值为", average)

def set_style(name, height, bold=False):
    style = xlwt.XFStyle()   # 初始化样式
    font = xlwt.Font()       # 为样式创建字体
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

def write_excel(forecast, dvalue, average, m, n, data, weights):    #写入文件
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row1 = [u'时间', u'实际值', u'预测值', u'绝对误差', u'平均绝对误差']
    sheet1.write(0, 0, '权值', set_style('Times New Roman', 220, True))
    for i in range(n):
        sheet1.write(0, i+1, weights[i], set_style('Times New Roman', 220, True))
    for i in range(0, len(row1)):
        sheet1.write(1, i, row1[i], set_style('Times New Roman', 220, True))
    for i in range(m+1):
        sheet1.write(i+2, 0, i+1, set_style('Times New Roman', 220, True))
    for i in range(len(data)):
        sheet1.write(i+2, 1, data[i], set_style('Times New Roman', 220, True))
    for i in range(n,m+1):
        sheet1.write(i+2, 2, forecast[i-n], set_style('Times New Roman', 220, True))
    for i in range(n, m):
        sheet1.write(i+2, 3, dvalue[i-n], set_style('Times New Roman', 220, True))
    sheet1.write(2, 4, average, set_style('Times New Roman', 220, True))
    f.save('out.xls')
    print("已经将数据写入out.xls")

def mian():
    printIntro()
    data, weights, n, m= datainput()
    forecast, dvalue,average = action(data, weights, n, m)
    print_Summary(forecast, dvalue, average, n)
    write_excel(forecast, dvalue, average, m, n, data, weights)

mian()
