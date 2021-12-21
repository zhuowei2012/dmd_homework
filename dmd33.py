# -*- coding: utf-8 -*-

import os
import xlsxwriter
import xlrd
import numpy as np
import math
fa = lambda x, y:x if x > y else y

NDM = [100,300,500]
NBP = [100,300,500]

# 新建excel
workbook = xlsxwriter.Workbook('第三题组合结果1217.xls')
# 新建工作薄
worksheet_sum = workbook.add_worksheet()
count_sum = 1
worksheet_sum.write("A%s" % count_sum, "Ndm")
worksheet_sum.write("B%s" % count_sum, "Nbp")
worksheet_sum.write("C%s" % count_sum, "投资组合")
worksheet_sum.write("D%s" % count_sum, "平均数")
worksheet_sum.write("E%s" % count_sum, "标准差")
worksheet_sum.write("F%s" % count_sum, "置信度95%下限")
worksheet_sum.write("G%s" % count_sum, "置信度95%上限")
worksheet_sum.write("H%s" % count_sum, "大于706的最优概率")
worksheet_sum.write("I%s" % count_sum, "马不英不") #00
worksheet_sum.write("J%s" % count_sum, "马不英规") #01
worksheet_sum.write("K%s" % count_sum, "马规英不") #10
worksheet_sum.write("L%s" % count_sum, "马规英规") #11
#worksheet_sum.write("I%s" % count_sum, "总收入")
count_sum+=1

data = xlrd.open_workbook('汇率期权_new.xls')  # 打开xls文件
table2 = data.sheets()[3]  # 仿真模型表
table3 = data.sheets()[4]  # 组合表
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

def arrange_data(n_dm, n_pm):
    for j in range(1, 83):
        if j < 2:
            continue

        imcome_sz = []
        income_00, income_01, income_10, income_11 = [], [], [], []
        Positive_profit_count = 0
        Positive_profit_count_00, Positive_profit_count_01, Positive_profit_count_10, Positive_profit_count_11 = 0, 0, 0, 0

        # 81种组合与每组随机数进行组合
        for i in range(1, round_count+5):
            if i < 5:  # 跳过第一行
                continue

            print('第 {i},{j}随机数据的81种组合'.format(i=i,j=j))
            # 马克汇率DM1
            dm1 = float(table2.row_values(i)[:][8])
            # 英镑汇率BP1
            bp1 = float(table2.row_values(i)[:][9])

            # 德国履约价格
            d_invest_price = float(table3.row_values(j)[:][0])
            # 德国卖权价格
            d_invest_cost = float(table3.row_values(j)[:][1])
            #德国数量
            d_invest_num = n_dm

            # 英国履约价格
            b_invest_price = float(table3.row_values(j)[:][3])
            # 英国卖权价格
            b_invest_cost = float(table3.row_values(j)[:][4])
            # 英国数量
            b_invest_num = n_pm

            #德国未规避
            income_D_no_avoid = income_D * dm1
            #英国未规避收入
            income_B_no_avoid = income_B * bp1

            #德国规避收入
            temp1 = d_invest_price - dm1
            income_D_avoid = income_D_no_avoid +  d_invest_num * (fa(temp1, 0) - d_invest_cost)

            #英国规避收入
            temp2 = b_invest_price - bp1
            income_B_avoid = income_B_no_avoid + b_invest_num * (fa(temp2,0) - b_invest_cost)


            #德国收入
           # income_D_real = fa(income_D_no_avoid, income_D_avoid)
            #德国收入折算美元
           # income_D_real_D = income_D_real * dm1
            #英国收入
            #income_B_real = fa(income_B_no_avoid, income_B_avoid)
            #英国收入折算美元
            # income_B_real_D = income_B_real * bp1
            #汇总收入
            # 00
            income_sum = income_D_no_avoid + income_B_no_avoid
            income_00.append(income_sum)
            if income_sum - profit_red_line > 0:
                Positive_profit_count_00 += 1

            # 01
            income_sum = income_D_no_avoid + income_B_avoid
            income_01.append(income_sum)
            if income_sum - profit_red_line > 0:
                Positive_profit_count_01 += 1

            # 10
            income_sum = income_D_avoid + income_B_no_avoid
            income_10.append(income_sum)
            if income_sum - profit_red_line > 0:
                Positive_profit_count_10 += 1

            # 11
            income_sum = income_D_avoid + income_B_avoid
            income_11.append(income_sum)
            if income_sum - profit_red_line > 0:
                Positive_profit_count_11 += 1

            #income_sum = income_D_real + income_B_real
            #imcome_sz.append(income_sum)
            #if income_sum - profit_red_line > 0:
            #    Positive_profit_count += 1

        # 计算平均数、标准差、置信度，大于706的概率
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
        #下限= 均值- C x S/根号N
        low_confidence = average - 1.96*std/(round_count**0.5)
        high_confidence = average + 1.96*std/(round_count**0.5)
        rate = Positive_profit_count / round_count
        global count_sum
        worksheet_sum.write("A%s" % count_sum, n_dm)
        worksheet_sum.write("B%s" % count_sum, n_bp)
        worksheet_sum.write("C%s" % count_sum, '({x},{y})'.format(x=(j - 2) % 9 + 1, y=math.ceil((j - 1) / 9)))
        worksheet_sum.write("D%s" % count_sum, average)
        worksheet_sum.write("E%s" % count_sum, std)
        worksheet_sum.write("F%s" % count_sum, low_confidence)
        worksheet_sum.write("G%s" % count_sum, high_confidence)
        worksheet_sum.write("H%s" % count_sum, rate)
        worksheet_sum.write("I%s" % count_sum, Positive_profit_count_00 / round_count)
        worksheet_sum.write("J%s" % count_sum, Positive_profit_count_01 / round_count)
        worksheet_sum.write("K%s" % count_sum, Positive_profit_count_10 / round_count)
        worksheet_sum.write("L%s" % count_sum, Positive_profit_count_11 / round_count)

        count_sum+=1

for n_dm in NDM:
    for n_bp in NBP:
        print("x:{x},y={y}".format(x=n_dm,y=n_bp))
        arrange_data(n_dm, n_bp)
# 关闭保存
workbook.close()
