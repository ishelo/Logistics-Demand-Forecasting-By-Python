import xlrd,xlwt

def printIntro():
    print("这个程序可用于一次指数平滑法")
    print("请将相关数据填写到input.xlsx")
    print('''请注意：平滑系数填写在第一行;
重新运行文件请关闭excel的文件，防止文件被其他应用占用。''')
    input("请输入任意字符继续")
    print("-" * 30)

def datainput():    #导入数据
    path = 'input.xlsx'
    rb = xlrd.open_workbook(path)
    data_sheet = rb.sheets()[0]
    rowNum = data_sheet.nrows
    m = rowNum - 1    #m为数据的个数
    s = []
    for i in range(1, rowNum):
        s.append(data_sheet.cell_value(i, 0))
    alpha = [0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9]
    print("数据导入成功")
    print("导入的实际值为：",s)
    print("平滑系数为：", alpha)
    print("正在进行的指数滑动预测计算...")
    print("-"*30)
    return alpha, s, m

def exponential_smoothing(alpha, s, m):
    result = []
    for j in range(9):
        s_temp = [0 for i in range(m)]
        s_temp[0] = s[0]
        SI = alpha[j]
        for i in range(1, m):
            x = SI * s[i] + (1 - SI) * s_temp[i-1]
            s_temp[i] = round(x, 3)
        result.append(s_temp)
    #print(result)
    return result

def difference_value(result, m, s):
    d_result = []
    for j in range(9):
        d_value = [0 for i in range(m)]
        for i in range(m -1):    #求绝对误差值
            y = abs(s[i+1]-result[j][i])
            d_value[i] = round(y, 3)
        d_result.append(d_value)
    #print(d_result)
    sum = []
    for j in range(9):
        z = 0
        for i in range(m-1):    #求平均绝对误差值
            z +=d_result[j][i]
        sum.append(z)
    average = []
    for j in range(9):
        w = round((sum[j]/(m-1)), 4)
        average.append(w)
    #print(average)
    return  d_result, average

def set_style(name, height, bold=False):
    style = xlwt.XFStyle()   # 初始化样式
    font = xlwt.Font()       # 为样式创建字体
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

def write_excel(result, d_result, average, m, s):    #写入文件
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row1 = [u'时间', u'实际值', u'预测值', u'绝对误差']
    for i in range(0, len(row1)):
        sheet1.write(0, i, row1[i], set_style('Times New Roman', 220, True))
    row2 = ['a = 0.1', 'a = 0.2', 'a = 0.3', 'a = 0.4','a = 0.5', 'a = 0.6', 'a = 0.7', 'a = 0.8','a = 0.9']
    for i in range(0, 9):   #行，误差序列
        sheet1.write(1, 2*i + 2, row2[i], set_style('Times New Roman', 220, True))
        sheet1.write(m + 3, 2*i + 2, '平均绝对误差', set_style('Times New Roman', 220, True))
        sheet1.write(m + 3, 2*i + 3, average[i], set_style('Times New Roman', 220, True))
    for i in range(m + 1):  #时间序列
        sheet1.write(i + 2, 0, i + 1, set_style('Times New Roman', 220, True))
    for i in range(m):
        sheet1.write(i + 2, 1, s[i], set_style('Times New Roman', 220, True))
    for j in range(9):
        for i in range(m):
            sheet1.write(i + 3, 2 + 2*j, result[j][i], set_style('Times New Roman', 220, True))
            sheet1.write(i + 3, 3 + 2*j, d_result[j][i], set_style('Times New Roman', 220, True))
    f.save('out.xls')
    print("已经将结果写入out.xls")

def main():
    datainput()
    alpha, s, m = datainput()
    result = exponential_smoothing(alpha, s, m)
    d_result, average = difference_value(result, m, s)
    write_excel(result, d_result, average, m, s)

if __name__ == '__main__':
    main()