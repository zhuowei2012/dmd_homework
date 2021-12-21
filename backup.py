# -*- coding: utf-8 -*-

import os
import xlsxwriter
import xlrd
#from decimal import Decimal

fa = lambda x, y:x if x > y else y

# 新建excel
workbook = xlsxwriter.Workbook('第二题组合结果.xls')
# 新建工作薄
worksheet = workbook.add_worksheet()

count = 1
worksheet.write("A%s" % count, "马克汇率DM1")
worksheet.write("B%s" % count, "英镑汇率BP1")

worksheet.write("C%s" % count, "德国履约价格")
worksheet.write("D%s" % count, "德国卖权价格")
worksheet.write("E%s" % count, "德国数量")
worksheet.write("F%s" % count, "英国履约价格")
worksheet.write("G%s" % count, "英国卖权价格")
worksheet.write("H%s" % count, "英国数量")

worksheet.write("I%s" % count, "德国未规避收入")
worksheet.write("J%s" % count, "英国未规避收入")
worksheet.write("K%s" % count, "德国规避收入")
worksheet.write("L%s" % count, "英国规避收入")
worksheet.write("M%s" % count, "是否规避德国")
worksheet.write("N%s" % count, "是否规避英国")
worksheet.write("O%s" % count, "德国收入")
worksheet.write("P%s" % count, "英国收入")
worksheet.write("Q%s" % count, "德国收入折算美元")
worksheet.write("R%s" % count, "英国收入折算美元")
worksheet.write("S%s" % count, "合计收入折算美元")
worksheet.write("T%s" % count, "德国亏损折算美元")
worksheet.write("U%s" % count, "英国亏损折算美元")
worksheet.write("V%s" % count, "汇率波动损失总额")
#worksheet.write("E%s" % count, "注册时间")
count += 1
data = xlrd.open_workbook('汇率期权.xls')  # 打开xls文件
table2 = data.sheets()[2]  # 仿真模型表
table3 = data.sheets()[3]  # 组合表
#for i in range(1, 150):
#nrows = table.nrows  # 获取表的行数

#原始数据
mean_D = 0  #均值
mean_B = 0  #均值
std_dev_D = 9   #马克标准差
std_dev_B = 11  #英镑标准差
relate = 0.675  #相关系数
rate_D = 0.6513  #马克当前汇率
rate_B = 1.234   #英镑当前汇率
income_D_D = 420  #马克收入美元
income_B_D = 336  #英镑收入 美元
income_D = 644.9  #马克收入
income_B = 272.3  #英镑收入


#选择一组随机数
for i in range(1, 35):
    if i < 5:  # 跳过第一行
        continue
    # print (table.row_values(i)[:5]) # 取前十三列
    #print(count, table.row_values(i)[:5][0])

    worksheet.write("A%s" % count, '第 {name} 个随机数据的81种组合'.format(name=count-1))
    count += 1
    #再选择81中组合，结合这组随机数得到的汇率计算的结果
    for j in range(1, 82):
        if j < 2:
            continue
        # 写入数据
        # 设定第一列（A）宽度为20像素 A:E表示从A到E
        #worksheet.set_column('A:A', 30)
        #worksheet.set_column('B:E', 20)
        worksheet.write("A%s" % count, table2.row_values(i)[:][8])  #马克汇率DM1
        worksheet.write("B%s" % count, table2.row_values(i)[:][9])  #英镑汇率BP1

        worksheet.write("C%s" % count, table3.row_values(j)[:][0]) #德国履约价格
        worksheet.write("D%s" % count, table3.row_values(j)[:][1]) #德国卖权价格
        worksheet.write("E%s" % count, table3.row_values(j)[:][2]) #德国数量
        worksheet.write("F%s" % count, table3.row_values(j)[:][3]) #英国履约价格
        worksheet.write("G%s" % count, table3.row_values(j)[:][4]) #英国卖权价格
        worksheet.write("H%s" % count, table3.row_values(j)[:][5]) #英国数量

        income_D_no_avoid = income_D * float(table2.row_values(i)[:][8])
        worksheet.write("I%s" % count, income_D_no_avoid) #德国未规避收入
        income_B_no_avoid = income_B * float(table2.row_values(i)[:][9])
        worksheet.write("J%s" % count, income_B_no_avoid) #英国未规避收入

        #德国规避收入
        temp1 = float(table3.row_values(j)[:][0]) - float(table2.row_values(i)[:][8])
        income_D_avoid = income_D_no_avoid + float(table3.row_values(j)[:][2]) * (fa(temp1, 0) - float(table3.row_values(j)[:][1]))
        worksheet.write("K%s" % count, income_D_avoid) #德国规避收入

        #英国规避收入
        temp2 = float(table3.row_values(j)[:][3]) - float(table2.row_values(i)[:][9])
        income_B_avoid = income_B_no_avoid + float(table3.row_values(j)[:][5]) * (
                    fa(temp2,0) - float(table3.row_values(j)[:][4]))
        worksheet.write("L%s" % count, income_B_avoid) #英国规避收入

        if income_D_no_avoid - income_D_avoid > 0:
            temp3 = '否'
        else:
            temp3 = '是'
        worksheet.write("M%s" % count,  temp3) #是否规避德国

        if income_B_no_avoid - income_B_avoid > 0:
            temp3 = '否'
        else:
            temp3 = '是'
        worksheet.write("N%s" % count, temp3) #是否规避英国

        income_D_real = fa(income_D_no_avoid, income_D_avoid)
        income_D_real_D = income_D_real * float(table2.row_values(i)[:][8])
        income_B_real = fa(income_B_no_avoid, income_B_avoid)
        income_B_real_D = income_B_real * float(table2.row_values(i)[:][9])

        worksheet.write("O%s" % count, income_D_real) #德国收入
        worksheet.write("P%s" % count, income_B_real) #英国收入
        worksheet.write("Q%s" % count, income_D_real_D) #德国收入折算美元
        worksheet.write("R%s" % count, income_B_real_D) #英国收入折算美元
        worksheet.write("S%s" % count, income_D_real_D + income_B_real_D) #合计收入折算美元
        worksheet.write("T%s" % count, income_D_D - income_D_real_D) #德国亏损折算美元
        worksheet.write("U%s" % count, income_B_D - income_B_real_D) #英国亏损折算美元
        worksheet.write("V%s" % count, income_D_D - income_D_real_D + income_B_D - income_B_real_D) #汇率波动损失总额

        count += 1

# 关闭保存
workbook.close()
