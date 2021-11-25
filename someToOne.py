#reportGenerator
#autor:Jade
#创建日期:2021/5/10
#-*- coding:utf-8 -*-

#将多个sheet合并为一个

import pandas as pd



df1 = pd.read_excel(r"C:\Users\Jade\Desktop\1.xlsx")
df2 = pd.read_excel(r"C:\Users\Jade\Desktop\2.xls")
print(df2.head())
result=df1.append(df2)
print(result.head())

result.to_excel("data.xlsx")
