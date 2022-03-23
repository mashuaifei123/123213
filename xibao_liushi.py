## 提取细胞数据的 流式
import pandas as pd
import os
from os import listdir
import numpy as np
import math
import copy
import time



def read_excel_xibao_liushi(path,outpath):
    list_path = [fn for fn in listdir(path) if fn.endswith('.xls')]
    xlsx_list_path = list(map(lambda y: path + '\\'+ y, list_path))
    for f in xlsx_list_path:
        writer = pd.ExcelWriter(outpath + '\\' + f[:-4].split('\\')[-1] + 'result.xlsx')
        df1 = pd.read_excel(f,  skiprows=1)
        df1 = pd.DataFrame(df1, columns=['Animal\nStudy #', 'Data Entry Method', 'Group'])
        df1.rename(columns={'Animal\nStudy #': 'animal'}, inplace = True)
        df1 = df1.dropna(how='all', axis=0)
        df1 = df1[df1['animal'].apply(lambda x: len(str(x)) < 10)]
        df1.reset_index(drop=True, inplace=True)

        df1_columns1_list = [str(i) for i in list(df1['animal'])]
        for i in range(len(df1_columns1_list)-1):
            if df1_columns1_list[i+1][1:3] != df1_columns1_list[i][1:3]:
                df1_columns1_list[i+1] = df1_columns1_list[i]
        df1['animal'] = df1_columns1_list
        df1.dropna(how='any',axis=0, inplace=True)
        df1['Group'] = df1['Group'].apply(lambda x: x.split(' ')[0])
        number = len(list(set(df1['Data Entry Method'])))
        list2 = list(set(df1['Data Entry Method']))
        print(list2)
        df1_1_top = df1.iloc[:number, :]
        df1_2_top = df1.iloc[number:number+number, :]
        for i in range(12, df1.shape[0], number):
            if int(i/number) % 2 == 0:
                df_concat = df1.iloc[i:i+number, :]
                df1_1_top = pd.concat([df1_1_top, df_concat], axis=0, ignore_index=True)
            else:
                df_concat = df1.iloc[i:i+number, :]
                df1_2_top = pd.concat([df1_2_top, df_concat], axis=0, ignore_index=True)

        list1 = ['CD3-CD2+P', 'CD3+CD4+P',  'CD3+CD8+P', 'CD3-CD14+P', 'CD3-CD20+P', 'CD20+CD21+P']

        df1_1_top.to_excel(writer, sheet_name='个体1', index=None)
        df1_2_top.to_excel(writer, sheet_name='个体2', index=None)

        PDR_set_l = []
        for x, i in enumerate(list1, 1):
            PDR_set_l.append((i, '{:02d}_{}'.format(x, i)))

        def id_to_idn(id, id_list):  # 根据及生成的列表加01- 用于排序
            for id_old in id_list:
                if id == id_old[0]:
                    return id_old[1]

        df1_1_top['Data Entry Method'] = df1_1_top[:]['Data Entry Method'].map(lambda x: id_to_idn(str(x), PDR_set_l))
        df1_2_top['Data Entry Method'] = df1_2_top[:]['Data Entry Method'].map(lambda x: id_to_idn(str(x), PDR_set_l))
        df1_1_top = df1_1_top.pivot('animal', 'Data Entry Method', 'Group')
        df1_2_top = df1_2_top.pivot('animal', 'Data Entry Method', 'Group')
        df1_1_top.reset_index(inplace=True)
        df1_2_top.reset_index(inplace=True)
        # df1_1_top.drop('Data Entry Method', axis=1, inplace=True)
        # df1_2_top.drop('Data Entry Method', axis=1, inplace=True)

        # for k in range(1, df1_1_top.shape[1]):
        #     df1_1_top.iloc[:, k] = df1_1_top.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)
        #     df1_2_top.iloc[:, k] = df1_2_top.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)
        list3 = ['CD3-CD2+NK细胞比例(%)', 'CD3+CD4+T细胞比例(%)', 'CD3+CD8+T细胞比例(%)', 'CD3-CD14+单核细胞比例(%)',
                 'CD3-CD20+B细胞比例(%)', 'CD21+CD20+B细胞比例(%)']
        df1_1_top.columns = ['动物编号'] + list3
        df1_2_top.columns = ['动物编号'] + list3
        df1_1_top.insert(0, 'Group', df1_1_top[:]['动物编号'].map(lambda x: str(x).strip()[:1]))  # 插入一列group
        df1_1_top['Group'] = df1_1_top['Group'].apply(pd.to_numeric, errors='errors')
        df1_2_top.insert(0, 'Group', df1_2_top[:]['动物编号'].map(lambda x: str(x).strip()[:1]))  # 插入一列group
        df1_2_top['Group'] = df1_2_top['Group'].apply(pd.to_numeric, errors='errors')

        # for k in range(1, df1_1_top.shape[1]):
        #     df1_1_top.iloc[:, k] = df1_1_top.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)
        #     df1_2_top.iloc[:, k] = df1_2_top.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)
        df1_1_top.to_excel(writer, sheet_name='Sheet1', index=None)
        df1_2_top.to_excel(writer, sheet_name='Sheet2', index=None)

        writer.save()
    print(df1_1_top)
    print(df1_2_top)
    # print(df1)




if __name__ == "__main__":
    start = time.time()
    path = r'K:\mashuaifei\11月1日\A2020006-T10-01\\'
    read_excel_xibao_liushi(path, path)
    end = time.time()
    print("运行时间为:%s" % (end - start))