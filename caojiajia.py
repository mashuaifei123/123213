import numpy as np
import pandas as pd
from os import listdir
import os
import re
import zipfile
import copy

inpath = r'K:\mashuaifei\新建文件夹 (2)\Fw_A2020030-T083-01-spss'
print('start')
xlsx_list = [fn for fn in listdir(inpath) if fn.endswith('.xlsx')]
xlsx_list.sort()
xlsx_list_path = list(map(lambda y: inpath + '\\' + y, xlsx_list))
print(xlsx_list_path)
for i in xlsx_list_path:
    df = pd.read_excel(i)
    df.dropna(how='all',axis=1,inplace=True)
    number = 5
    length = int(int(df.shape[1]+1)/number)
    print(length)
    index1 = list(df.columns)[:length]  # 时间段
    index2 = list(df.iloc[0,:])  # 坐标名字
    index2_1 = [index2[i:i+len(index1)] for i in range(0,len(index2),len(index1))]
    index = [index2_1[i][j] + '_' + index1[j] for i in range(len(index2_1))  for j in range(len(index1))]
    print(index)
    df.reset_index(drop=False,inplace=True)
    df.drop(0,axis=0,inplace=True)
    df.columns = ['动物编号'] + index
    df.insert(0, 'Group', df[:]['动物编号'].map(lambda x: str(x).strip()[:1]))  # 插入一列group
    df['Group'] = df['Group'].apply(pd.to_numeric, errors='errors')

    for k in range(2, 2+len(index1)*2):
        df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.1f' % (x + 0.000000000001)))
        df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='errors')
    for k in range(2+len(index1)*2,df.shape[1] ):
        df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.0f' % (x)))
        df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='errors')

    df_solid = df.iloc[:, 0:2]
    name = list(set(index2))
    name .sort(key = index2.index) # 原顺序指标名字

    for i in range(int((df.shape[1] - 2) / len(index1))):
        df1 = df.iloc[:, i * len(index1) + 2:i * len(index1) + len(index1) + 2]
        df2 = pd.concat([df_solid, df1], axis=1)
        # df2.to_excel(writer1, sheet_name='sheet'+str(i+2), index=None)
        df2.to_excel(inpath + '\\' + name[i] + '.xlsx', index=None)
        df2['sex'] = df2['动物编号'].apply(lambda x :x[1])
        df2.sort_values(['sex','Group'],inplace=True)
        df2.drop('sex',axis = 1,inplace = True)
        df2_F = df2.iloc[:int(df.shape[0]/2),:]
        df2_F.to_excel(inpath + '\\' + name[i] + 'F.xlsx', index=None)
        df2_M = df2.iloc[int(df.shape[0]/2):,:]
        df2_M.to_excel(inpath + '\\' + name[i] + 'M.xlsx', index=None)
    #df1 = df.iloc[:length, 1:]



print(df)