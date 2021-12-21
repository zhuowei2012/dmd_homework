# -*- coding: utf-8 -*-

import os
import xlsxwriter
import xlrd
import numpy as np
import math
#from decimal import Decimal

fa = lambda x, y:x if x > y else y

# 新建excel
workbook = xlsxwriter.Workbook('第二题组合结果1217.xls')
# 新建工作薄
worksheet = workbook.add_worksheet()
worksheet_sum = workbook.add_worksheet()

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
worksheet.write("M%s" % count, "00")
worksheet.write("N%s" % count, "01")
worksheet.write("O%s" % count, "10")
worksheet.write("P%s" % count, "11")
#worksheet.write("Q%s" % count, "德国收入折算美元")
#worksheet.write("R%s" % count, "英国收入折算美元")
#worksheet.write("S%s" % count, "合计收入折算美元")
#worksheet.write("T%s" % count, "德国亏损折算美元")
#worksheet.write("U%s" % count, "英国亏损折算美元")
#worksheet.write("V%s" % count, "汇率波动损失总额")

count_sum = 1
worksheet_sum.write("A%s" % count_sum, "投资组合")
worksheet_sum.write("B%s" % count_sum, "平均数")
worksheet_sum.write("C%s" % count_sum, "标准差")
worksheet_sum.write("D%s" % count_sum, "置信度95%下限")
worksheet_sum.write("E%s" % count_sum, "置信度95%上限")
worksheet_sum.write("F%s" % count_sum, "大于706的最优概率")
worksheet_sum.write("G%s" % count, "马不英不") #00
worksheet_sum.write("H%s" % count, "马不英规") #01
worksheet_sum.write("I%s" % count, "马规英不") #10
worksheet_sum.write("J%s" % count, "马规英规") #11

#worksheet.write("E%s" % count, "注册时间")
count += 1
count_sum+=1
data = xlrd.open_workbook('汇率期权_new.xls')  # 打开xls文件
table2 = data.sheets()[3]  # 仿真模型表
table3 = data.sheets()[4]  # 组合表
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
profit_red_line = 706

round_count = 300

#先获取81种组合
for j in range(1, 83):
    if j < 2:
        continue

    imcome_sz = []
    income_00,income_01,income_10,income_11 = [],[],[],[]
    Positive_profit_count = 0
    Positive_profit_count_00,Positive_profit_count_01,Positive_profit_count_10,Positive_profit_count_11 = 0,0,0,0

    worksheet.write("A%s" % count, '第 {cn} 组({x},{y})投资组合的200种数据'.format(cn=j-1, x=(j-2)%9+1, y=math.ceil((j-1)/9)))
    count += 1
    worksheet.write("A%s" % count, "平均数")
    worksheet.write("B%s" % count, "标准差")
    worksheet.write("C%s" % count, "置信度95%下限")
    worksheet.write("D%s" % count, "置信度95%上限")
    worksheet.write("E%s" % count, "最优比例")
    worksheet.write("F%s" % count, "马不英不") #00
    worksheet.write("G%s" % count, "马不英规") #01
    worksheet.write("H%s" % count, "马规英不") #10
    worksheet.write("I%s" % count, "马规英规") #11
    profile_line = count+1
    count += 2
    # 81种组合与每组随机数进行组合
    for i in range(1, round_count+5):
        if i < 5:  # 跳过第一行
            continue

        print('第 {i},{j}随机数据的81种组合'.format(i=i,j=j))
        worksheet.write("A%s" % count, table2.row_values(i)[:][8])  #马克汇率DM1
        worksheet.write("B%s" % count, table2.row_values(i)[:][9])  #英镑汇率BP1

        worksheet.write("C%s" % count, table3.row_values(j)[:][0]) #德国履约价格
        worksheet.write("D%s" % count, table3.row_values(j)[:][1]) #德国卖权价格
        worksheet.write("E%s" % count, table3.row_values(j)[:][2]) #德国数量
        worksheet.write("F%s" % count, table3.row_values(j)[:][3]) #英国履约价格
        worksheet.write("G%s" % count, table3.row_values(j)[:][4]) #英国卖权价格
        worksheet.write("H%s" % count, table3.row_values(j)[:][5]) #英国数量

        #计算未规避收入
        income_D_no_avoid = income_D * float(table2.row_values(i)[:][8])
        worksheet.write("I%s" % count, income_D_no_avoid) #德国未规避收入
        income_B_no_avoid = income_B * float(table2.row_values(i)[:][9])
        worksheet.write("J%s" % count, income_B_no_avoid) #英国未规避收入

        #德国规避收入
        temp1 = float(table3.row_values(j)[:][0]) - float(table2.row_values(i)[:][8])
        income_D_avoid = income_D_no_avoid + float(table3.row_values(j)[:][2]) * \
                         (fa(temp1, 0) - float(table3.row_values(j)[:][1]))
        worksheet.write("K%s" % count, income_D_avoid) #德国规避收入

        #英国规避收入
        temp2 = float(table3.row_values(j)[:][3]) - float(table2.row_values(i)[:][9])
        income_B_avoid = income_B_no_avoid + float(table3.row_values(j)[:][5]) * (
                    fa(temp2,0) - float(table3.row_values(j)[:][4]))
        worksheet.write("L%s" % count, income_B_avoid) #英国规避收入

        #00
        income_sum = income_D_no_avoid + income_B_no_avoid
        income_00.append(income_sum)
        if income_sum - profit_red_line > 0:
            Positive_profit_count_00 += 1
        worksheet.write("M%s" % count, income_sum)  # 00

        # 01
        income_sum = income_D_no_avoid + income_B_avoid
        income_01.append(income_sum)
        if income_sum - profit_red_line > 0:
            Positive_profit_count_01 += 1
        worksheet.write("N%s" % count, income_sum)  # 00

        # 10
        income_sum = income_D_avoid + income_B_no_avoid
        income_10.append(income_sum)
        if income_sum - profit_red_line > 0:
            Positive_profit_count_10 += 1
        worksheet.write("O%s" % count, income_sum)  # 00

        # 11
        income_sum = income_D_avoid + income_B_avoid
        income_11.append(income_sum)
        if income_sum - profit_red_line > 0:
            Positive_profit_count_11 += 1
        worksheet.write("P%s" % count, income_sum)  # 00

        count += 1

    # 计算平均数、标准差、置信度，大于706的概率
    Positive_profit_count = Positive_profit_count_00
    imcome_sz = income_00[:]
    if Positive_profit_count_01 > Positive_profit_count:
        Positive_profit_count = Positive_profit_count_01
        imcome_sz = income_01[:]
    if Positive_profit_count_10 > Positive_profit_count:
        Positive_profit_count = Positive_profit_count_10
        imcome_sz = income_10[:]
    if Positive_profit_count_11 > Positive_profit_count:
        Positive_profit_count = Positive_profit_count_11
        imcome_sz = income_11[:]

    average = np.average(imcome_sz)
    std = np.std(imcome_sz, ddof=1)
    #下线= 均值- C x S/根号N
    low_confidence = average - 1.96*std/(round_count**0.5)
    high_confidence = average + 1.96*std/(round_count**0.5)
    rate = Positive_profit_count / round_count
    worksheet.write("A%s" % profile_line, average)
    worksheet.write("B%s" % profile_line, std)
    worksheet.write("C%s" % profile_line, low_confidence)
    worksheet.write("D%s" % profile_line, high_confidence)
    worksheet.write("E%s" % profile_line, rate)
    worksheet.write("F%s" % profile_line, Positive_profit_count_00 / round_count)
    worksheet.write("G%s" % profile_line, Positive_profit_count_01 / round_count)
    worksheet.write("H%s" % profile_line, Positive_profit_count_10 / round_count)
    worksheet.write("J%s" % profile_line, Positive_profit_count_11 / round_count)

    worksheet_sum.write("A%s" % count_sum, '({x},{y})'.format(x=(j-2)%9+1, y=math.ceil((j-1)/9)))
    worksheet_sum.write("B%s" % count_sum, average)
    worksheet_sum.write("C%s" % count_sum, std)
    worksheet_sum.write("D%s" % count_sum, low_confidence)
    worksheet_sum.write("E%s" % count_sum, high_confidence)
    worksheet_sum.write("F%s" % count_sum, rate)
    worksheet_sum.write("G%s" % count_sum, Positive_profit_count_00/ round_count)
    worksheet_sum.write("H%s" % count_sum, Positive_profit_count_01/ round_count)
    worksheet_sum.write("I%s" % count_sum, Positive_profit_count_10/ round_count)
    worksheet_sum.write("J%s" % count_sum, Positive_profit_count_11/ round_count)
    count_sum+=1
# 关闭保存
workbook.close()
