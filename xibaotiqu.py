## 提取细胞数据的

"""
output  首行要空出两行，
只需要两列对照
末尾的— 可以删除
"""

import pandas as pd
import os
from os import listdir
import numpy as np
import math
import copy
import time


def Merge(dict1, dict2):
    res = {**dict1, **dict2}
    return res


def id_to_idn(id, id_list):  # 根据及生成的列表加01- 用于排序
    for id_old in id_list:
        if id == id_old[0]:
            return id_old[1]


def geshi2(df2,time):
    length = df2.shape[0]
    df_solid = df2.iloc[:, 0:2]
    df_leader = df2.iloc[:, :2 + len(time)]
    #     print(df_leader)
    df_index = df2.columns.tolist()[2:]
    df_index1 = [i.split('_')[0] for i in df_index]
    df_index2 = list(set(df_index1))
    df_index2.sort(key=df_index1.index)
    #     print(df_index2)

    df_leader.columns = ['Group', '动物编号'] + time

    print(int((df2.shape[1] - 2) / len(time)))
    for i in range(1, int((df2.shape[1] - 2) / len(time))):
        df_concat = df2.iloc[0:, i * len(time) + 2:i * len(time) + len(time) + 2]
        #print(df_concat)
        df_concat1 = pd.concat([df_solid, df_concat], axis=1)  # 添上前缀group和编号
        df_concat1.columns = ['Group', '动物编号'] + time
        df_leader = pd.concat([df_leader, df_concat1], axis=0, ignore_index=True)
    #     print(length)
    df_index = [df_index2[int(i / length)] if i % int(length) == 0 else '' for i in range(df_leader.shape[0])]
    #     print(df_index)
    df_leader['Group'] = df_index
    df_leader = df_leader.rename(columns={'Group': '指标'})
    return  df_leader


def excel_spss(df_mean):
    # 读取路径中的所有excel，放入列表等待处理
    df_mean_list = df_mean.iloc[:, 1].tolist()
    PDR_set = list(set(df_mean_list))
    PDR_set.sort(key=df_mean_list.index)  # 保留顺序转换列表->去重复集合->列表原始顺序排序
    PDR_set_l = []
    for x, i in enumerate(PDR_set, 1):
        PDR_set_l.append((i, '{:02d}_{}'.format(x, i)))

    df_mean['检测时间'] = df_mean[:]['检测时间'].map(lambda x: id_to_idn(str(x), PDR_set_l))
    df_mean.index = [df_mean['动物编号'].tolist(), df_mean['检测时间'].tolist()]
    df_mean.drop(['动物编号'], axis=1, inplace=True)
    df_mean.drop(['检测时间'], axis=1, inplace=True)
    df_mean = df_mean.unstack()

    CO1_name = list(df_mean.columns.levels[0])
    CO2_name = list(df_mean.columns.levels[1])
    RO2_name = list(df_mean.index)
    '''组合成一个列名字'''
    title_list = list(
        map(lambda x: '{}'.format(str(x), str(x)), [str(i) + '_' + str(j)[3:] for i in CO1_name for j in CO2_name]))
    df_spss_ok = pd.DataFrame(df_mean.values, columns=title_list, index=RO2_name)
    df_spss_ok = df_spss_ok.reset_index()
    df_spss_ok.rename(columns={'index': '动物编号'}, inplace=True)  # 重名命列名字
    df_spss_ok.insert(0, 'Group', df_spss_ok[:]['动物编号'].map(lambda x: str(x).strip()[:1]))  # 插入一列group
    # print(df_spss_ok.dtypes)
    df_spss_ok['Group'] = df_spss_ok['Group'].apply(pd.to_numeric, errors='coerce')
    second_name_list = CO1_name
    time = PDR_set

    for k in range(2 + len(time), df_spss_ok.shape[1]):
        df_spss_ok.iloc[:, k] = df_spss_ok.iloc[:, k].apply(lambda x: str('%.2f' % (x)))
        df_spss_ok.iloc[:, k] = df_spss_ok.iloc[:, k].apply(pd.to_numeric, errors='coerce')

    return df_spss_ok


def read_excel_xibao(path,outpath):
    list_path = [fn for fn in listdir(path) if fn.endswith('.xls')]
    xlsx_list_path = list(map(lambda y: path + '\\'+ y, list_path))
    if 'Output.xls' in list_path:
        df_sub = pd.read_excel(path + '\\'+ 'Output.xls', header=None, skiprows=3)
        try:
            df_sub = df_sub[:df_sub[df_sub[0] == 'Pretest Animals not Assigned to Dose Groups: '].index.tolist()[0]]
            df_sub = df_sub.iloc[:, [0, 1]]
        except:
            df_sub = df_sub.iloc[:, [0, 1]]
            print(df_sub)
        iddict1 = dict((i, j) for i, j in zip(list(df_sub[1]), list(df_sub[0])))
        print(iddict1)
        # iddict2 = dict( (i,j) for i,j in zip(list(df_sub[0]),list(df_sub[0])))
        # iddict = Merge(iddict1,iddict2)
    for f in xlsx_list_path:
        if 'Output' not in f:
            df_box = []
            writer = pd.ExcelWriter(outpath + '\\' + f[:-4].split('\\')[-1] + 'result.xlsx')
            df1 = pd.read_excel(f, header=None, skiprows=1)

            # 获得excel的文件拆分的行list1_index
            df_2 = df1[0].tolist()
            list1 = list(set((df1[0])))
            list1.sort(key=df_2.index)
            # list2 = list(map(lambda x :0.0 if math.isnan(x) else x ,list1))
            list1.remove(np.nan)
            print(list1)
            list1_index = [df1[df1[0] == i].index.tolist()[0] for i in list1]
            list1_index.append(df1.shape[0])
            print(list1_index)
            list2 = [x.split(' ')[2] for x in list1]
            # 之后根据index拆分文件操作
            #print(df1)
            for m in range(len(list1_index) - 1):
                df = df1[list1_index[m] + 10 + 2:list1_index[m + 1]]
                #df = df.iloc[:, [4, 11]]
                s = 4   #   初始是4
                v =11    #  初始是11
                df = df.iloc[:, [s, v]]
                df[s] = df[s].apply(lambda x: str(x).split('.')[0])

                # 如果有output文件需要修正动物id
                if 'Output.xls' in list_path:
                    df.reset_index(drop=True)
                    df.insert(0, 'a', df[4])
                    df['a'] = df['a'].apply(lambda x: x.split('(')[0])
                    df[4] = df[4].apply(lambda x: (x.split('(')[1]).split(')')[0])
                    df = df.reset_index(drop=True)
                    df['a'] = df['a'].apply(lambda x: x.replace(' ', ''))
                    list_id = list(df['a'])
                    list_id = [i for i in list_id if len(i) == 5]
                    iddict2 = dict((i, j) for i, j in zip(list_id, list_id))
                    iddict = Merge(iddict1, iddict2)
                    # print(iddict)

                    df['a'] = df['a'].replace(iddict1)
                    df_mean_list = df[4].tolist()
                    time = list(set(df_mean_list))
                    time.sort(key=df_mean_list.index)
                    PDR_set_l = []
                    for x, i in enumerate(time, 1):
                        PDR_set_l.append((i, '{:02d}_{}'.format(x, i)))

                    df[4] = df[:][4].map(lambda x: id_to_idn(str(x), PDR_set_l))
                    for k in range(2, df.shape[1]):
                        df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)
                    print(df)
                    df.to_excel(r'K:\mashuaifei\Fw_A2020038-T10-01 恒河猴静脉注射给予8MW2311免疫毒性试验 细胞因子 统计分析\1.xlsx', index=None)
                    df_box.append(df)
                    print(df.shape[0])
                    print(len(list(set(list(df['a'])))))



                    df = df.pivot('a', 4, 11)
                    df = df.reset_index()
                    for k in range(1, df.shape[1]):
                        df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)
                        # df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='errors')

                    df.columns = ['动物编号'] + time
                    # df = df[df['动物编号'].map(len) < 6]
                    df.to_excel(writer, sheet_name=list2[m], index=None)
                else:
                    df.reset_index(drop=True)
                    df.insert(0, 'a', df[s])
                    df['a'] = df['a'].apply(lambda x: x.split('(')[0])
                    df[s] = df[s].apply(lambda x: (x.split('(')[1]).split(')')[0])

                    df_mean_list = df[s].tolist()
                    time = list(set(df_mean_list))
                    time.sort(key=df_mean_list.index)
                    PDR_set_l = []
                    for x, i in enumerate(time, 1):
                        PDR_set_l.append((i, '{:02d}_{}'.format(x, i)))

                    df[s] = df[:][s].map(lambda x: id_to_idn(str(x), PDR_set_l))
                    df = df.reset_index(drop=True)
                    for k in range(2, df.shape[1]):
                        df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)

                    df_box.append(df)

                    df = df.pivot('a', s, v)
                    df = df.reset_index()
                    df.columns = ['动物编号'] + time
                    for k in range(1, df.shape[1]):
                        df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.2f' % (x))).astype(float)
                    print(list1[m])
                    df.to_excel(writer, sheet_name=list2[m], index=None)

            dff_col = [x.split(' ')[2] for x in list1]
            print(dff_col)
            dff = df_box[0]
            print(len(df_box))
            for i in range(len(df_box) - 1):
                dff = dff.merge(df_box[i + 1], how='left', on=['a', s])
            print(dff)
            dff.columns = ['动物编号', '检测时间'] + dff_col

            dff['a1'] = dff['动物编号'].apply(lambda x: x[1])
            dff['a2'] = dff['动物编号'].apply(lambda x: x[0])
            dff = dff.sort_values(by=['检测时间', '动物编号'])
            dff = dff.sort_values(by=['a1', '检测时间', '动物编号'])
            dff['检测时间'] = dff['检测时间'].apply(lambda x: x.split('_')[1])
            dff = dff.drop(['a1', 'a2'], axis=1)

            dff.to_excel(writer, sheet_name='汇总', index=None)
            dff['panduan'] = dff['动物编号'].apply(lambda x: x[1])
            dff1 = dff[dff['panduan'].isin(['F'])]
            dff2 = dff[dff['panduan'].isin(['M'])]
            dff1 = dff1.drop('panduan', axis=1)
            dff2 = dff2.drop('panduan', axis=1)

            dff1_spss = excel_spss(dff1)
            dff2_spss = excel_spss(dff2)

            dff11 = copy.deepcopy(dff1_spss)
            dff22 = copy.deepcopy(dff2_spss)

            dff1.reset_index(inplace=True)
            dff2.reset_index(inplace=True)
            dff1.rename(columns = {'level_0':'动物编号','level_1':'检测时间'},inplace = True)
            dff2.rename(columns = {'level_0':'动物编号','level_1':'检测时间'},inplace = True)

            dff1['检测时间'] = dff1['检测时间'].apply(lambda x: x.split('_')[1])
            dff2['检测时间'] = dff2['检测时间'].apply(lambda x: x.split('_')[1])

            df_list1  = list(dff1['检测时间'])
            time_f = list(set(list(dff1['检测时间'])))

            time_f.sort(key =df_list1.index)

            time2 = list(dff2['检测时间'])
            time_m = list(set(time2))
            time_m.sort(key = time2.index)

            df_geshiF = geshi2(dff11,time_f)
            df_geshiM = geshi2(dff22,time_m)
            df_geshiF.to_excel(writer, sheet_name='汇总F', index=None)
            df_geshiM.to_excel(writer, sheet_name='汇总M', index=None)


            dff1.to_excel(writer, sheet_name='F',index = None)
            dff2.to_excel(writer, sheet_name='M',index = None)

            dff1_spss.to_excel(writer, sheet_name='F_spss格式', index=None)
            dff2_spss.to_excel(writer, sheet_name='M_spss格式', index=None)
            writer.save()


if __name__ == "__main__":
    start = time.time()
    path = r'K:\mashuaifei\translation-谷歌\A2021007-T014-01 细胞因子统计分析\\'
    read_excel_xibao(path, path)
    end = time.time()
    print("运行时间为:%s" % (end - start))
