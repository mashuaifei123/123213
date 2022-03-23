# #### 健康检查有一位，两位，整数####
# N=3 在下面
import xlrd
import numpy as np
import pandas as pd
from os import listdir
import os
import re
import zipfile
from excel_xianzhux_linjian import get_df1, get_df3
#from excel_xianzhuxing3 import get_df4
'''
Variable requirement:  varibale_time

varibale:指标变量


time：实验时间 

Notes ：we need '_'  to  to split variables
such as:  心率(次/分)_p4  ,T波(mV)_p2233

'''
def get_decimal(df):
    if "a. Limited to first 100 cases." in df[0]:
        index = df[df[0] == "a. Limited to first 100 cases."].index.tolist()
    else:
        data_list1 = df.loc[df[0].str.contains('试验',na =False)].index
        data_list2 = df.loc[df[0].str.contains('QC：',na =False)].index
        print(data_list1)
        print(data_list2)
        index = [max(list(data_list1)), max(list(data_list2)) - 1]
    # print(index)
    lines_decimal = []
    df1 = df[index[0] + 2:index[1] ]
    # print(df1)
    for i in range(4, df1.shape[1]):
        lines_col = list(df1.iloc[:, i])
        lines__col_decimal = [len(str(j).split('.')[1]) if '.' in str(j) else 0 for j in lines_col]
        # print(list(set(lines__col_decimal)))
        # print(lines__col_decimal)
        if len(list(set(lines__col_decimal))) == 1 and max(list(set(lines__col_decimal))) <= 5:
            decimal = lines__col_decimal[0]
        elif len(list(set(lines__col_decimal))) == 1 and max(list(set(lines__col_decimal))) > 5:
            decimal = lines__col_decimal[0]
        elif len(list(set(lines__col_decimal))) == 2 and max(list(set(lines__col_decimal))) <= 5:
            decimal = max(list(set(lines__col_decimal)))
        elif len(list(set(lines__col_decimal))) == 2 and max(list(set(lines__col_decimal))) > 5:
            ccc = list(set(lines__col_decimal))
            ccc.pop(ccc.index(max(ccc)))
            decimal = max(ccc)
        elif len(list(set(lines__col_decimal))) >= 3 and max(list(set(lines__col_decimal))) <= 5:
            decimal = max(list(set(lines__col_decimal)))
        elif len(list(set(lines__col_decimal))) >= 3 and max(list(set(lines__col_decimal))) > 5:
            ccc = list(set(lines__col_decimal))
            ccc = [i for i in ccc if i <8]
            # ccc.pop(ccc.index(max(ccc)))
            decimal = max(ccc)
        else:
            print('数据有问题')
        # print(decimal)
        lines_decimal.append(decimal)
        #lines_decimal = [2, 2, 2, 2, 2, 2, 2, 2]
    #lines_decimal = [2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2]
    #lines_decimal = [1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]
    #lines_decimal = [2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2,2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]
    lines_decimal=[2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2,
     2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]
    #lines_decimal = [1,1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
    return lines_decimal


def clean(df, sigList, time, lines_decaimal):
    print(lines_decaimal)
    # Lie hang
    print(df.shape[1])
    print(df.shape[0])
    for i in range(1, df.shape[1], 3):
        for j in range(df.shape[0]):
            mean_decimal = lines_decaimal[j] + 1
            sd_decimal = lines_decaimal[j] + 2
            df.iloc[j, i] = str(('%.' + str(mean_decimal) + 'f') % df.iloc[j, i]) if df.iloc[j, i] != '' else ''
            df.iloc[j, i + 1] = str(('%.' + str(sd_decimal) + 'f') % df.iloc[j, i + 1]) if df.iloc[j, i + 1] != '' else ''
            if df.iloc[j, i + 1] != '':
                df.iloc[j, i] = df.iloc[j, i] + '±' + df.iloc[j, i + 1]
            else:
                df.iloc[j, i] = df.iloc[j, i]
    #print(df)
    #print('之后')
    df = sig_work(df, sigList)
    return df


def Complex_clean(df: pd.DataFrame,name: str, sigList: list, time: int) -> pd.DataFrame:
    for i in range(1, df.shape[1], 3):
        if name == '呼吸':
            df.iloc[0:time, i:i + 1] = df.iloc[0:time, i:i + 1].applymap(
                lambda x: ("{:.1f} ").format(x+0.0000001) if x != '' else '').astype(str)
            df.iloc[0:time, i + 1:i + 2] = df.iloc[0:time, i + 1:i + 2].applymap(
                lambda x: ("{:.2f}").format(x+0.0000001) if x != '' else '').astype(str)
            df.iloc[time:df.shape[0], i:i + 1] = df.iloc[time:df.shape[0], i:i + 1].applymap(
                lambda x: ("{:.2f}").format(x+0.0000001) if x != '' else '')
            df.iloc[time:df.shape[0], i + 1:i + 2] = df.iloc[time:df.shape[0], i + 1:i + 2].applymap(
                lambda x: ("{:.3f}").format(x+0.0000001) if x != '' else '')
        #print(df)

        elif name == '心电':
            df.iloc[0:time, i:i + 1] = df.iloc[0:time, i:i + 1].applymap(
                lambda x: ("{:.1f} ").format(x) if x != '' else '').astype(str)
            df.iloc[0:time, i + 1:i + 2] = df.iloc[0:time, i + 1:i + 2].applymap(
                lambda x: ("{:.2f}").format(x) if x != '' else '').astype(str)
            df.iloc[time:df.shape[0], i:i + 1] = df.iloc[time:df.shape[0], i:i + 1].applymap(
                lambda x: ("{:.3f}").format(x) if x != '' else '')
            df.iloc[time:df.shape[0], i + 1:i + 2] = df.iloc[time:df.shape[0], i + 1:i + 2].applymap(
                lambda x: ("{:.4f}").format(x) if x != '' else '')
        elif name == 'PD':
            df[i] = df[i].apply(lambda x: str('%.4f' % x) if x != '' else '')
            df[i + 1].apply(
                lambda x: str('%.5f') % x if x != '' else '')

        elif name == '精子':
            list_t= [i for i in range(df.shape[0])]
            list1 = [2, 12, 13, 14,  17, 27, 28, 29]
            list2 = (list(set(list_t).difference(set(list1))))

            for j in list1:
                df.iloc[j, i:i+1] = df.iloc[j, i:i + 1].apply(
                    lambda x: ("{:.1f} ").format(x) if x != '' else '').astype(str)
                df.iloc[j, i + 1:i + 2] = df.iloc[j, i + 1:i + 2].apply(
                    lambda x: ("{:.2f}").format(x) if x != '' else '').astype(str)
            for k in list2:
                df.iloc[k, i:i+1] = df.iloc[k, i:i + 1].apply(
                    lambda x: ("{:.2f} ").format(x) if x != '' else '').astype(str)
                df.iloc[k, i + 1:i + 2] = df.iloc[k, i + 1:i + 2].apply(
                    lambda x: ("{:.3f} ").format(x) if x != '' else '').astype(str)

        elif  name == '血液学':
            decimal_1 = [1, 1, 0, 1, 0, 1, 3, 2, 2, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 0]
            decimal_1_after = [val for val in decimal_1 for i in range(2)]
            for j in range(df.shape[0]):
                mean_decimal = decimal_1_after[j] + 1
                sd_decimal = decimal_1_after[j] + 2
                df.iloc[j, i] = str(('%.' + str(mean_decimal) + 'f') % df.iloc[j, i]) if df.iloc[j, i] != '' else ''
                df.iloc[j, i + 1] = str(('%.' + str(sd_decimal) + 'f') % df.iloc[j, i + 1]) if df.iloc[j, i + 1] != '' else ''
                if df.iloc[j, i + 1] != '':
                    df.iloc[j, i] = df.iloc[j, i] + '±' + df.iloc[j, i + 1]
                else :
                    df.iloc[j, i] = df.iloc[j, i]
        elif name == '血生化':
            decimal_1 = [1, 1, 0, 1, 0, 0, 0, 0, 1, 1, 2, 2, 2, 1, 2, 1, 2, 0, 1, 2, 2]
            decimal_1_after = [val for val in decimal_1 for i in range(2)]
            for j in range(df.shape[0]):
                mean_decimal = decimal_1_after[j] + 1
                sd_decimal = decimal_1_after[j] + 2
                df.iloc[j, i] = str(('%.' + str(mean_decimal) + 'f') % df.iloc[j, i]) if df.iloc[j, i] != '' else ''
                df.iloc[j, i + 1] = str(('%.' + str(sd_decimal) + 'f') % df.iloc[j, i + 1]) if df.iloc[j, i + 1] != '' else ''
                df.iloc[j, i] = df.iloc[j, i] + '±' + df.iloc[j, i + 1]

        elif name == '凝血' or name == '体重':
            df[i] = df[i].apply(lambda x: str('%.2f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.3f') % x if x != '' else '')

        elif name == '摄食量':
            df[i] = df[i].apply(lambda x: str('%.3f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.4f') % x if x != '' else '')
        elif name == '脏脑比' or name == '脏器重量':
            df[i] = df[i].apply(lambda x: str('%.5f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.6f') % x if x != '' else '')
        elif name == '脏体比':
            df.iloc[0:time, i:i + 1] = df.iloc[0:time, i:i + 1].applymap(
                lambda x: '%.2f' % x if x != '' else '').astype(str)
            df.iloc[0:time, i + 1:i + 2] = df.iloc[0:time, i + 1:i + 2].applymap(
                lambda x: '%.3f' % x if x != '' else '').astype(str)
            df.iloc[time:df.shape[0], i:i + 1] = df.iloc[time:df.shape[0], i:i + 1].applymap(
                lambda x: '%.5f' % x if x != '' else '')
            df.iloc[time:df.shape[0], i + 1:i + 2] = df.iloc[time:df.shape[0], i + 1:i + 2].applymap(
                lambda x: '%.6f' % x if x != '' else '')
            df[i] = df[i].apply(lambda x: x if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + x if x != '' else '')

        else:
            df[i] = df[i].apply(lambda x: str('%.3f' % x) if x != '' else '')
            df[i + 1].apply(
                lambda x: str('%.4f') % x if x != '' else '')
    #print(df)
    #print('之后')
    # for i in range(1, df.shape[1], 3):
    #     df[i] = df[i].apply(lambda x: str(x) if x != '' else '') + df[i + 1].apply(
    #         lambda x: '±' + str(x) if x != '' else '')


    df = sig_work(df, sigList)
    return df


def get_df2(df: pd.DataFrame, time: int) :
    print(df)
    for i in range(1, df.shape[1], 3):
        df[i + 2] = df[i + 2].apply(lambda x:  'n=' + str(x) if x != '' else '')
        df.drop(i + 1, axis=1, inplace=True)
    #print(df)
    df = pd.DataFrame(np.repeat(df.values, 2, axis=0))
    for j in range(1, df.shape[0], 2):
        for k in range(1, df.shape[1], 2):
            df.iloc[j, k] = df.iloc[j, k+1]
    #df = df.drop(1, axis=1)
    #print(df)
    df = AddIndex(df, time)
    return df

def easy_clean(df: pd.DataFrame, name: str, sigList: list, time):
    #print('dasdasdasdasdddddddddddddddddddddddddddd')
    for i in range(1, df.shape[1], 3):
        if name == '血压':
            df[i] = df[i].apply(lambda x: str('%.1f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.2f') % x if x != '' else '')
        elif name == 'kg':
            df[i] = df[i].apply(lambda x: str('%.3f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.4f') % x if x != '' else '')
        elif name == '细胞':
            df[i] = df[i].apply(lambda x: str('%.3f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.4f') % x if x != '' else '')
        elif name == 'OTC':
            df[i] = df[i].apply(lambda x: str('%.3f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.4f') % x if x != '' else '')
        else:
            df[i] = df[i].apply(lambda x: str('%.2f' % x) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.3f') % x if x != '' else '')
    #print(sigList)
    df = sig_work(df, sigList)
    return df


def count_cells(df, out_path, title, time, group_time):

    df11 = df.T
    df11.columns = df11.iloc[0, :]
    df11.drop(df11.index[0], axis=0, inplace=True)

    df_column = list(set(df11.columns))
    df_column.sort(key=list(df11.columns).index)
    df_column.pop(1)
    df_index = list(df11.index)
    df_new = []
    s_len = 0

    for i in range(len(group_time)):
        df12 = df11.iloc[:, s_len:s_len + group_time[i]]
        s_len += group_time[i]
        y = list(df12.isnull().sum(axis=1))  # 空值的个数
        df12.fillna(1)
        x = list((df12 != 0).astype(int).sum(axis=1))  # 不等于0的值,会把空值也计算进去
        xy_Divide = [str(x[ii] - y[ii]) + '/' + str(group_time[i] - y[ii]) for ii in range(len(x))]
        df_new.append(xy_Divide)



    dff = pd.DataFrame(df_new)
    df2 = dff

    df2.index = df_column
    df2.columns = df_index
    df1 = df2.T
    df1 = df1.reset_index()
    df1.insert(0, '指标', df1['index'])
    sep = '_'
    #print(df1)
    df1['index'] = df1['index'].apply(lambda x: x.split(sep, 1)[1])
    df1['指标'] = df1['指标'].apply(lambda x: x.split(sep, 1)[0])
    b = [i for i in range(0, df1.shape[0], time)]  # time 问题 即时间都有两行
    bb = [i for i in range(0, df1.shape[0], 2)]
    a = [df1['指标'][i] if (i) in b else '' for i in range(0, len(df1['指标']))]
    aa = [df1['index'][i] if (i) in bb else '' for i in range(0, len(df1['index']))]
    df1['指标'] = a

    df_11 = list(set([x.split('_')[0] for x in df_index]))
    df_11.sort(key=[x.split('_')[0] for x in df_index].index)
    df_12 = list(set([x.split('_')[1] for x in df_index]))
    df_12.sort(key=[x.split('_')[1] for x in df_index].index)
    df_12 = [str(i) + df_12[i] for i in range(len(df_12))]
    # print(df_11)
    # print(df_12)
    # print(df_column)
    # print(df_new)
    df_2 = pd.DataFrame(df_new, index=df_column, columns=pd.MultiIndex.from_product([df_11, df_12]))
    df_2 = df_2.stack()
    df_2 = df_2.reset_index()
    df_2.columns = ['组别', '时间'] +df_11
    df_2['时间'] = df_2['时间'].apply(lambda x :x[1:])
    return df1, df_2


# 开始加*操作的前奏，df排序
def sig_work(df, sigList:list ):
    sort_key = list(df[0])
    by = list(df[3])
    # 排序df，先把大于3进行oneway的指标放在前面，不做的放在后面，并以此进行排序
    sort_key1 = [sort_key[m] for m in range(len(by)) if by[m] !='' and by[m] >= 3]

    for n in range(len(by)):
        if by[n] == ''or by[n] < 3:
            sort_key1.append(sort_key[n])
    #print(sort_key1)
    df[0] = df[0].astype('category')
    df[0].cat.reorder_categories(sort_key1, inplace=True)
    df.sort_values(by=0, inplace=True)
    df = df.reset_index(drop=True)
    ## 排序结束,开始标*
    print(df)

    if sigList:
        print(sigList)
        for j in sigList:
            if j[2] == 1:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '*'
            else:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '**'
     # 标*结束，之后返回顺序

    df[0] = df[0].astype('category')
    df[0].cat.reorder_categories(sort_key, inplace=True)
    df.sort_values(by=0, inplace=True)
    df = df.reset_index(drop=True)
    print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
    #print(df)
    return df


def AddIndex(df: pd.DataFrame, time: int):
    for i in range(2,df.shape[1],2):
        df = df.drop(i,axis=1)
    df.insert(0, 'd', df[0])
    if '_' in df[0][1]:
        sep = '_'
        df[0] = df[0].apply(lambda x: x.split(sep, 1)[1])
        df['d'] = df['d'].apply(lambda x: x.split(sep, 1)[0])


    b = [i for i in range(0, df.shape[0], time*2)] # time 问题 即时间都有两行
    bb = [i for i in range(0, df.shape[0], 2)]
    a = [df['d'][i] if (i) in b else '' for i in range(0, len(df['d']))]
    aa = [df[0][i] if (i) in bb else '' for i in range(0, len(df[0]))]
    df['d'] = a
    #print(a)
    #print(aa)

    df[0] = aa
    return df


def select(df1,Group_name):
    '''
    Intercepting  tables  from excel exported by SPSS.
    and become the format you need.

    0        start
    1    [6, 2, 2]

    '''
    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'Oneway'].index.tolist()[0] - 13]
    #sig = df1[df1[df1[0] == 'Dunnett t (2-sided)a'].index.tolist()[0]:df1[df1[0] == 'a. Dunnett t-tests treat one group as a control, and compare all other groups against it.'].index.tolist()[ 0]]
    #sig = list(sig.iloc[:, 5])[3:-1]

    df = df.iloc[:3*len(Group_name)+1, :]  # 截取正确行数
    df = df.drop([0, 1], axis=1)  # 删除前三列无用列

    df = df.dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    df = df.T
    df = df.reset_index(drop=True)
    df.replace(np.nan, '', inplace=True)
    # print(df)
    # print('初始')
    return df


def select1(df1):
    index = df1[df1[0] == "a. Limited to first 100 cases."].index.tolist()
    #print(index)
    df = df1[index[0] + 3: index[1]-1]

    df = df.dropna(axis=1, how='all')
    #print(df)
    df = df.drop([0,2,3], axis= 1)
    df = df.reset_index(drop= True)
    df.iat[0, 0] = '组别'
    df.columns = list(df.iloc[0,:])
    df = df.drop(0,axis= 0)
    return  df


def getTime(df):
    list1 = list(df[0])
    print(list1)
    s = int(len(list1) / len(set([x.split('_')[0] for x in list1])))
    return s


def GetGRoup(file_list):
    df1 = pd.read_excel(file_list, header=None)
    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'Oneway'].index.tolist()[0] - 6]

    Group_name = list(df[0].dropna(axis=0))[1:]

    if 'Total' in Group_name:
        Group_name.remove('Total')
    group = len(Group_name)-1
    print(Group_name)
    return Group_name, group


def Get_homogeneity(df1):
    df = df1[df1[df1[0] == 'Test of Homogeneity of Variances'].index.tolist()[0]+2:df1[df1[0] == 'ANOVA'].index.tolist()[0]-1 ]
    #df = df.dropna(axis=1, how='all')
    var_list = list(df[0])
    df.fillna(0.05, inplace= True)
    sig = list(df[4])
    #print(sig)
    return sig, var_list


def Get_anove(df1):
    df = df1[df1[df1[0] == 'ANOVA'].index.tolist()[0]+2:df1[df1[0] == 'Post Hoc Tests'].index.tolist()[0]-2 ]
    print(len(df[6]))
    #df = df.dropna(axis=1, how='all')
    df.fillna(0.05, inplace = True)

    sig = [list(df[6])[i] for i in range(len(df[6])) if i%3 == 0]
    print(sig)

    return sig


def Get_comparisions(df1):
    df = df1[df1[df1[0] == 'Dunnett t (2-sided)a'].index.tolist()[0]+3:df1[df1[
                                                                              0] == 'a. Dunnett t-tests treat one group as a control, and compare all other groups against it.'].index.tolist()[
        0]-1]
    df = df.dropna(axis=1, how='all')
    sig = list(df[5])
    list3 = list(df[0])
    return sig, list3


def Get_kwtest(df1):
    df = df1[df1[df1[0] == 'Test Statisticsa,b'].index.tolist()[0]+1:df1[df1[0] == 'a. Kruskal Wallis Test'].index.tolist()[0] ]
    df = df.dropna(axis=1, how='all')
    sig = list(df.iloc[3,:])[1:]
    return sig
    #print(sig)


def Get_manhui(df1,liebiao):
    df = df1[df1[df1[0] == 'Mann-Whitney Test'].index.tolist()[0] + 1:
             df1[df1[0] == 'a. Grouping Variable: Group'].index.tolist()[-1]]

    df = df.dropna(axis=1, how='all')
    list5 = []
    #print(df.shape[1])
    #print(df)
    for j in range(1, len(liebiao)+1):
        for i in range(len(list(df[0]))):
            if list(df[0])[i] == 'Asymp. Sig. (2-tailed)':
                list5.append(list(df[j])[i])
    #print(list5)
    #list6.append(list5)
    return list5


def judge(df1,group):
    print(group)
    sig_list = []
    list1, var_list = Get_homogeneity(df1)
    print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxsssssssssssss')

    print(len(var_list))
    # remove_list =[]
    # for i in range(1,len(var_list)):
    #     if (i+1)% 8 == 0:
    #         remove_list.append(var_list[i])
    # print(remove_list)
    # print(var_list)
    CanAnoveList = [var_list[i] for i in range(len(list1)) if list1[i] > 0.05]
    CanAnoveList_locate = [i for i in range(len(list1)) if list1[i] > 0.05]
    CantAnoveList = [var_list[i] for i in range(len(list1)) if list1[i] < 0.05]
    CantAnoveList_locate = [i for i in range(len(list1)) if list1[i] < 0.05]
    print("CanAnoveList%s" %(CanAnoveList))
    print("CantAnoveList%s" %(CantAnoveList))
    print("CanAnoveList_locate%s" %(CanAnoveList_locate))
    list2 = Get_anove(df1)
    print("list2%s" %(list2))
    anove_list = [list2[i] for i in range(len(list2)) if i in CanAnoveList_locate]       # 去掉了方差不齐的变量
    anove_list_locate = [i for i in range(len(list2)) if i in CanAnoveList_locate]        # 方差不齐变量的位置
    anove_list1 = [i%(group) for i in anove_list]
    print("anove_list1%s" %(anove_list1))
    different_var = [CanAnoveList[i] for i in range(len(anove_list1)) if anove_list1[i] < 0.05]    # 去掉方差不齐的变量后存在不同的变量
    print(different_var)
    if different_var:
        comparisons_list, list3 = Get_comparisions(df1)

        '''
        comparisons_list 是Multiple Comparisons 表格的第一列指标名 是str   a list of str
        list3 是该表的sig  ,是 a list of  float number
        '''
        group_comparisons_list = [comparisons_list[j:j+group] for i in different_var for j in range(len(comparisons_list)) if list3[j] == i]
        group_comparisons_list_locate = [j for i in different_var for j in range(len(comparisons_list)) if list3[j] == i]
        print('group_comparisons_list%s'% group_comparisons_list)
        for i in range(len(group_comparisons_list)):
            for j in range(len(group_comparisons_list[i])):
                if 0.01 < group_comparisons_list[i][j] < 0.05:
                    sig_list.append([var_list.index(different_var[i]), (j+1)*3+1, 1])
                if group_comparisons_list[i][j] < 0.01:
                    sig_list.append([var_list.index(different_var[i]), (j+1)*3+1, 2])
                else:
                    pass
        #print(different_var[0])
        #print(var_list.index(different_var[0]))
        print(sig_list)
    if CantAnoveList:
        kw_test_sig_list = Get_kwtest(df1)
        print(kw_test_sig_list)
        ManW_list = [CantAnoveList[i] for i in range(len(kw_test_sig_list)) if kw_test_sig_list[i] < 0.05]
        print('ManW_list%s' % ManW_list)
        if  ManW_list:
            '''
            ManW_list: 是需要进行满惠特尼分析的变量列表  ['P波(mV)_R1']
            mhtn_list: 是读取excel的获得判断数据列表   
            new_mhtn_list : 是把list嵌套起来反映不同变量  [[0.07580017458236125, 0.00820773610734887, 0.09369261949324824]]
            '''
            #print('进行满惠特尼')
            mhtn_list = Get_manhui(df1,ManW_list)
            #print(mhtn_list)
            new_mhtn_list = [mhtn_list[x:x + group] for x in range(0, len(mhtn_list), group)]
            #print(new_mhtn_list)
            #print(new_mhtn_list[0][1])
            for i in range(len(ManW_list)):
                for j in range(len(new_mhtn_list[i])):
                    if 0.01 < new_mhtn_list[i][j] < 0.05:
                        sig_list.append([var_list.index(ManW_list[i]), (j+1)*3+1, 1])
                    if  new_mhtn_list[i][j] < 0.01:
                        sig_list.append([var_list.index(ManW_list[i]), (j+1)*3+1, 2])
    print(sig_list)
    print('judge finished')
    return(sig_list)


def convert_zip(zip_path, save_path):
    '''
    # zip_path 要压缩文件的路径
    # save_path 文件压缩后保存的路径
    '''
    xlsx_list = [fn for fn in listdir(zip_path) if fn.endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_list_path = list(map(lambda y: zip_path + '\\'+y, xlsx_list))
    zip = zipfile.ZipFile(save_path + '\\total.zip', "w")  # zipfile.ZIP_DEFLATED    逗号前是压缩包的名字
    for file in xlsx_list_path:
        zip.write(file)
    zip.close()


def xianzhuxing1(path_list, out_path):

    '''
    s: 获得实验的时间数
    group：获得实验的小组数
    group_name :组别的名字
    '''
    print('start')
    xlsx_list = [fn for fn in listdir(path_list) if fn.endswith('.xlsx')]
    bb = list(map(lambda y: path_list + '\\' + y, xlsx_list))
    print(bb)
    Group_name, group = GetGRoup(bb[0])
    df_col = ['指标', '时间段']
    df_col= df_col + Group_name
    '''
    df1: 读取整个excel
    从df1中读取到时间段
    '''
    for f in bb:
        title = f[:-5].split('\\')[-1]
        print(title)
        df1 = pd.read_excel(f, header=None)
        df1_decimal = get_decimal(df1)
        df = select(df1, Group_name)
        s = getTime(df)

        sig_list = judge(df1, group)
        #sig_list = []
        df = clean(df, sig_list, s, df1_decimal)
        df11 = get_df2(df, s)
        df11.columns = df_col
        df22 = get_df1(df, s, Group_name)
        df33 = get_df3(df, Group_name, s)


        # if '血压' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list = judge(df1, group)
        #     df = easy_clean(df, '血压', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #     #df11.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        # elif '心电' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list=judge(df1, group)
        #     df = Complex_clean(df, '心电', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #     #df33.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        # elif '呼吸' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1,Group_name)
        #     sig_list = judge(df1, group)
        #     if '呼吸频率' in df[0][0]:
        #         s = getTime(df)
        #         df = Complex_clean(df,'呼吸',  sig_list, s)
        #     else:
        #         # 当没有呼吸频率时，呼吸的操作与体温一样
        #         s = getTime(df)
        #         df = easy_clean(df, '体温', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #     #df11.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        # elif '体温' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1,Group_name)
        #     s = getTime(df)
        #     sig_list = judge(df1, group)
        #     df = easy_clean(df, '体温', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #     #df11.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        # elif 'PD' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list = judge(df1, group)
        #     df = easy_clean(df, 'PD', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #     # df11.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        # elif '精子' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = 2
        #     sig_list = judge(df1, group)
        #     df = Complex_clean(df, '精子',  sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #     #df11.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        # elif '体重' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     #print(df)
        #     sig_list = judge(df1, group)
        #     if 'kg' in f:
        #         df = easy_clean(df, 'kg', sig_list, s)
        #     else:
        #         df = easy_clean(df, '体温', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #     #df11.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        # elif '血液学' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list=judge(df1, group)
        #     df = Complex_clean(df, '血液学', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #
        # elif '血生化' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list=judge(df1, group)
        #     df = Complex_clean(df, '血生化', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        # elif '凝血' in f or '体重' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list=judge(df1, group)
        #     df = Complex_clean(df, '凝血', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        # elif '脏器' in f or '脏脑' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list=judge(df1, group)
        #     df = Complex_clean(df, '脏脑比', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #
        # elif '脏体' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list=judge(df1, group)
        #     df = Complex_clean(df, '脏体比', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #
        # elif '摄食量' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list=judge(df1, group)
        #     df = Complex_clean(df, '摄食量', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        #
        # elif  '细胞' in f:
        #     df1 = pd.read_excel(f, header=None)
        #     if 'Oneway' in list(df1[0]):
        #         df = select(df1, Group_name)
        #         df_initial = select1(df1)
        #         print(df_initial)
        #         df_initial.to_excel(out_path + '\\' +  'dasok.xlsx', index=None)
        #         s = getTime(df)
        #         sig_list = judge(df1, group)
        #         df_1, df_2 = count_cells(df_initial, path_list, title, s)
        #         df = easy_clean(df, '细胞', sig_list,s)
        #         df11 = get_df2(df, s)
        #         df11.columns = df_col
        #         df22 = get_df1(df, s, Group_name)
        #         df33 = get_df3(df, Group_name, s)
        #         #df11.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        #
        #         writer = pd.ExcelWriter(out_path + '\\' + title + '检出率.xlsx')
        #         df_1.to_excel(writer, sheet_name='sheet1', index=None)
        #         df_2.to_excel(writer, sheet_name='sheet2', index=None)
        #         writer.save()
        # else :
        #     df1 = pd.read_excel(f, header=None)
        #     df = select(df1, Group_name)
        #     s = getTime(df)
        #     sig_list = judge(df1, group)
        #     df = easy_clean(df, 'OTC', sig_list, s)
        #     df11 = get_df2(df, s)
        #     df11.columns = df_col
        #     df22 = get_df1(df, s, Group_name)
        #     df33 = get_df3(df, Group_name, s)
        if '细胞' in f:
            df1 = pd.read_excel(f, header=None)
            if 'Oneway' in list(df1[0]):
                df = select(df1, Group_name)
                df_initial = select1(df1)
                s = getTime(df)
                sig_list = judge(df1, group)


                group_time = [j for i, j in enumerate(list(df.iloc[0, :])) if i % 3 == 0][1:]
                print(f'每个组的数量是{group_time}')
                df_1, df_2 = count_cells(df_initial, path_list, title, s, group_time)


                df = easy_clean(df, '细胞', sig_list,s)
                df11 = get_df2(df, s)
                df11.columns = df_col
                df22 = get_df1(df, s, Group_name)
                df33 = get_df3(df, Group_name, s)
                writer = pd.ExcelWriter(out_path + '\\' + title + '检出率.xlsx')
                df_1.to_excel(writer, sheet_name='sheet1', index=None)
                df_2.to_excel(writer, sheet_name='sheet2', index=None)
                writer.save()

        writer1 = pd.ExcelWriter(out_path + '\\' + title + 'ok.xlsx')
        df22.to_excel(writer1, sheet_name='sheet1', index=None)
        df11.to_excel(writer1, sheet_name='sheet2', index=None)
        df33.to_excel(writer1, sheet_name='sheet3', index=None)

        writer1.save()

    convert_zip(out_path, out_path)

if __name__ == "__main__":
    ## 记得修改小数位数的列表
    ## 记得修改小数位数的列表
    ## 记得修改小数位数的列表
    path_list = r'K:\mashuaifei\死亡进行的显著性\2022-02-22\\'
    xianzhuxing1(path_list, path_list)


