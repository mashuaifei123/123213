# #### 健康检查有一位，两位，整数####
# su wan ying


import numpy as np
import pandas as pd
from os import listdir
import os
import re
import zipfile


'''
Variable requirement:  varibale_time

varibale:指标变量

time：实验时间 

Notes ：we need '_'  to  to split variables
such as:  心率(次/分)_p4  ,T波(mV)_p2233
读取数据结构：           mean    sd      N
0       心率(次/分)_P3  108.800  24.5     5  
1       心率(次/分)_P9  114.200  33.1     5  
2       心率(次/分)_R1  122.200  26.3     5 
3      心率(次/分)_R28   87.500   4.9     2  

'''


def Complex_clean(df: pd.DataFrame, name: str, time: int, sigList: list, group_list:list, group:int) -> pd.DataFrame:
    '''

    依次：df， 文件名字，实验时间，判断列表，组别名字，组别数
    Take decimals for 心电 and breath.

    notes: There are two condition in breathing, Whether there is breathing-range is the criterion.

    心率 is no longer the first indicator, sometimes it would show up in a line
    before : 心率 pr期间  qrs期间 ......
    after : pr 期间 心率 ......

    '''

    if name == '心电':
        for i in range(1, df.shape[1], 3):
            df.iloc[0:time, i:i + 1] = df.iloc[0:time, i:i + 1].applymap(
                lambda x: ("{:.1f} ").format(x) if x != '' else '').astype(str)
            df.iloc[0:time, i + 1:i + 2] = df.iloc[0:time, i + 1:i + 2].applymap(
                lambda x: ("{:.2f}").format(x) if x != '' else '').astype(str)
            df.iloc[time:df.shape[0], i:i + 1] = df.iloc[time:df.shape[0], i:i + 1].applymap(
                lambda x: ("{:.3f}").format(x) if x != '' else '')
            df.iloc[time:df.shape[0], i + 1:i + 2] = df.iloc[time:df.shape[0], i + 1:i + 2].applymap(
                lambda x: ("{:.4f}").format(x) if x != '' else '')
    if name == '呼吸':
        for i in range(1, df.shape[1], 3):
            df.iloc[0:time, i:i + 1] = df.iloc[0:time, i:i + 1].applymap(
                lambda x: ("{:.1f} ").format(x) if x != '' else '').astype(str)
            df.iloc[0:time, i + 1:i + 2] = df.iloc[0:time, i + 1:i + 2].applymap(
                lambda x: ("{:.2f}").format(x) if x != '' else '').astype(str)
            df.iloc[time:df.shape[0], i:i + 1] = df.iloc[time:df.shape[0], i:i + 1].applymap(
                lambda x: ("{:.2f}").format(x) if x != '' else '')
            df.iloc[time:df.shape[0], i + 1:i + 2] = df.iloc[time:df.shape[0], i + 1:i + 2].applymap(
                lambda x: ("{:.3f}").format(x) if x != '' else '')

    for i in range(1, df.shape[1], 3):
        df[i] = df[i].apply(lambda x: str(x) if x != '' else '') + df[i + 1].apply(
            lambda x: '±' + str(x) if x != '' else '')

    df = sig_work(df, sigList)
    df = data_del(df, time, group_list, group)
    return df


def easy_clean(df: pd.DataFrame, name: str, sigList: list, time: int, group_list:list, group:int):
    '''
    Easy_clean  for 血压 and 体温
    first step:  reserved decimal
    second step : mark * or **
    third step :  process data format

    '''
    # 保留小数并相加一列

    if name == '血压':
        for i in range(1, df.shape[1], 3):
            df[i] = df[i].apply(lambda x: str('%.1f' % (x + 0.0000000001)) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.2f') % x if x != '' else '')

    if name == '体温':
        for i in range(1, df.shape[1], 3):
            df[i] = df[i].apply(lambda x: str('%.2f' % (x + 0.0000000001)) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.3f') % x if x != '' else '')
    if name == '呼吸':
        for i in range(1, df.shape[1], 3):
            df[i] = df[i].apply(lambda x: str('%.1f' % (x + 0.0000000001)) if x != '' else '') + df[i + 1].apply(
                lambda x: '±' + str('%.2f') % x if x != '' else '')

    # 进行加*操作，返回数据格式不变
    df = sig_work(df, sigList)
    df = data_del(df, time, group_list, group)

    return df


def data_del(df: pd.DataFrame,  time: int, Group_name, group):
    '''
    main def ,start del with talble
    :param df:  no why
    :param time:  no what
    :return:   a df
    '''
    t = df[3]

    for i in range(1, df.shape[1], 3):
        df.drop(i + 1, axis=1, inplace=True)
        df.drop(i + 2, axis=1, inplace=True)
    df.insert(1, 't', t)
    df['t'] = df['t'].apply(lambda x: str(x) if x != '' else '')
    # 复制一列变量列 ，拆分变量名 MBP-P7 -->  MBP 和P7

    df.insert(0, 'd', df[0])
    if '_' in df[0][1]:
        sep = '_'
        df[0] = df[0].apply(lambda x: x.split(sep, 1)[1])
        df['d'] = df['d'].apply(lambda x: x.split(sep, 1)[0])
    # 生成一列时间+每组个数  P7(n=5)
    df[0] = df[0].apply(lambda x: x + '(n=').astype(str) + df['t'].apply(lambda x: str(x) + ')').astype(str)

    # 把指标列存起来，从df中删除，获得df的实验设置的组数，为下面的转置铺垫
    df.drop('t', axis=1, inplace=True)
    var = df['d']  # 指标名
    df.drop('d', axis=1, inplace=True)
    group_n = df.shape[1]  # 从未转置的excel得到实验设置的组数
    # 转置 ，获得目标格式的前一步,它是目标格式的延伸版，如目标为3*5，此时即为1*15
    df = df.T
    # 开始裁剪，将第一部分裁剪命名为df0，并且在原df中去掉，防止影响循环。
    # 拼接需要索引相同，从下往上接，给一个new-name，确保几个部分都是相同索引
    # 成功将1*15格式，拼接成3*5
    df0 = df.iloc[:, 0:time]
    new_name = []
    for jj in range(time):
        df.drop(jj, axis=1, inplace=True)  # 这里df删掉了第一部分
        new_name.append(jj)  # 新索引   [0,1,2,3,4,5]
    for ii in range(1, int(df.shape[1] / time) + 1):
        df1 = df.iloc[:, (ii - 1) * time: (ii) * time]  # [6:12]  [12:18]
        df1.columns = new_name
        df0 = pd.concat([df0, df1], ignore_index=True)
    df = df0
    # 拼接结束，此时没有变量名，需要把之前保存的变量名插入
    # 去重之前的变量列表，set法中sort，bykey报错不知原因
    # 获得新的变量列表插入
    var1 = []
    for i in var:
        if i not in var1:
            var1.append(i)
    new_index = []
    for i in range(len(var1)):
        for j in range(group_n):
            new_index.append(var1[i])
    new_index[0] = '指标'

    #print(len(new_index))
    #print(df.shape[0])
    df.insert(0, 'a', new_index)

    # 插入新变量之后，如果以后期要每个指标的N不同，就需要这块代码
    for j in range(0, df.shape[0], group_n):
        if j > 1:
            df.drop(j, axis=0, inplace=True)
    # 删除每组的p7(n=5)，只留下第一行之后，对第一列进行精简
    b = []
    c = []
    df = df.reset_index(drop=True)
    for i in range(df.shape[0]):
        if df['a'][i] not in b:
            b.append(df['a'][i])
            c.append(i)
    #print(df)
    a = [df['a'][i] if i in c else '' for i in range(0, len(df['a']))]
    df['a'] = a
    #list_group1 = [v for v in Group_name for i in range(int(df.shape[0]/(group+1)))]
    #print(list_group1)
    #['安慰剂对照组', '安慰剂对照组', 'DS002低剂量组', 'DS002低剂量组', 'DS002中剂量组', 'DS002中剂量组', 'DS002高剂量组', 'DS002高剂量组']
    list_group = Group_name * (int((df.shape[0]/(group+1))))
    #['安慰剂对照组', 'DS002低剂量组', 'DS002中剂量组', 'DS002高剂量组', '安慰剂对照组', 'DS002低剂量组', 'DS002中剂量组', 'DS002高剂量组', '安慰剂对照组', 'DS002低剂量组', 'DS002中剂量组', 'DS002高剂量组']

    list_group.insert(0,'组别')
    df.insert(1, 'b', list_group)


    return df


def sig_work(df, sigList: list):
    '''
    this def  is used to mark

    :param df:  no what
    :param sigList:   a list about how to locate the location of the table
    :return:  a df with *or **
    '''
    # 开始加*操作的前奏，df排序
    # print(df)
    sort_key = list(df[0])
    by = list(df[3])

    # 排序df，先把大于3进行oneway的指标放在前面，不做的放在后面，并以此进行排序
    sort_key1 = [sort_key[m] for m in range(len(by)) if by[m] != '' and by[m] > 3]

    for n in range(len(by)):
        if by[n] == '' or by[n] < 3:
            sort_key1.append(sort_key[n])
            # print(sort_key1)
    df[0] = df[0].astype('category')

    df[0].cat.reorder_categories(sort_key1, inplace=True)
    df.sort_values(by=0, inplace=True)
    df = df.reset_index(drop=True)
    ## 排序结束,开始标*
    #print(df)
    if sigList:
        #print(sigList)
        for j in sigList:
            if j[2] == 1:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '*'
                #print(df.iloc[j[0], j[1]])
            else:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '**'
                #print(df.iloc[j[0], j[1]])
    # 标*结束，之后返回顺序
    df[0] = df[0].astype('category')
    df[0].cat.reorder_categories(sort_key, inplace=True)
    df.sort_values(by=0, inplace=True)
    df = df.reset_index(drop=True)
    return df


def select(df1):
    '''
    Intercepting  tables  from excel exported by SPSS.
    and become the format you need.

    0        start
    1    [6, 2, 2]

    '''
    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'Oneway'].index.tolist()[0] - 5]
    #sig = df1[df1[df1[0] == 'Dunnett t (2-sided)a'].index.tolist()[0]:df1[df1[0] == 'a. Dunnett t-tests treat one group as a control, and compare all other groups against it.'].index.tolist()[ 0]]
    #sig = list(sig.iloc[:, 5])[3:-1]
    df = df.drop([0, 1], axis=1)  # 删除前三列无用列

    df = df.dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    df = df.T
    df = df.reset_index(drop=True)
    df.replace(np.nan, '', inplace=True)
    return df


def getTime(df, var: str) -> int:
    '''
    Due to different time, it is necessary to determine the time of special indicators.
    like 呼吸频率 0 decimal，心率 0 decimal ，血压 0 decimal .

    '''
    list1 = df[0]
    # print(list1)
    s = 0
    if var == '呼吸':
        for i in list1:
            if i[:4] == '呼吸频率':
                s += 1
    if var == '心电':
        for i in list1:
            if i[:2] == '心率':
                s += 1
    if var == '血压':
        for i in list1:
            if i[:3] == 'SBP':
                s += 1
    if var == '体温':
        s = len(list1)
    return s


def GetGRoup(file_list):
    '''
    :param file_list: 随机挑选一份文件
    :return:  本次试验的组别列表 , 和组数
    for forample : ['溶媒对照组', 'TL139低剂量组', 'TL139中剂量组', 'TL139高剂量组'] ,则group = 3
    '''
    df1 = pd.read_excel(file_list, header=None)
    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'Oneway'].index.tolist()[0] - 5]
    Group_name = list(df[0].dropna(axis=0))[1:]
    group = len(Group_name)-1
    print(Group_name)
    return Group_name, group


def Get_homogeneity(df1):
    df = df1[df1[df1[0] == 'Test of Homogeneity of Variances'].index.tolist()[0]+2:df1[df1[0] == 'ANOVA'].index.tolist()[0]-1 ]
    df = df.dropna(axis=1, how='all')
    sig = list(df[4])
    var_list = list(df[0])
    #print(sig)
    return sig,var_list


def Get_anove(df1):
    df = df1[df1[df1[0] == 'ANOVA'].index.tolist()[0]+2:df1[df1[0] == 'Post Hoc Tests'].index.tolist()[0]-2 ]
    df = df.dropna(axis=1, how='all')
    sig=[i for i in list(df[6]) if str(i) != 'nan']
    #print(sig)

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


def Get_manhui(df1):
    df = df1[df1[df1[0] == 'Mann-Whitney Test'].index.tolist()[0] + 1:
             df1[df1[0] == 'a. Grouping Variable: Group'].index.tolist()[-1]]
    df = df.dropna(axis=1, how='all')
    list5 = []
    for i in range(len(list(df[0]))):
        if list(df[0])[i] == 'Asymp. Sig. (2-tailed)':
            list5.append(list(df[1])[i])
    return list5


def judge(df1,group):
    sig_list=[]
    list1, var_list = Get_homogeneity(df1)
    CanAnoveList = [var_list[i] for i in range(len(list1)) if list1[i] > 0.05]
    #print(CanAnoveList)
    CantAnoveList = [var_list[i] for i in range(len(list1)) if list1[i] < 0.05]
    CantAnoveList_locate = [i for i in range(len(list1)) if list1[i] < 0.05]
    #print(CantAnoveList)

    list2= Get_anove(df1)
    anove_list = [list2[i] for i in range(len(list2)) if i not in CantAnoveList_locate]       # 去掉了方差不齐的变量
    anove_list_locate =  [i for i in range(len(list2)) if i  in CantAnoveList_locate]        # 方差不齐变量的位置

    different_var = [CanAnoveList[i] for i in range(len(anove_list)) if anove_list[i] < 0.05]    # 去掉方差不齐的变量后存在不同的变量
    #print(different_var)
    if different_var:
        comparisons_list, list3 = Get_comparisions(df1)

        '''
        comparisons_list 是Multiple Comparisons 表格的第一列指标名 是str   a list of str
        list3 是该表的sig  ,是 a list of  float number
        '''
        group_comparisons_list = [comparisons_list[j:j+3] for i in different_var for j in range(len(comparisons_list))  if list3[j] == i]
        group_comparisons_list_locate = [j for i in different_var for j in range(len(comparisons_list))  if list3[j] == i]
        #print(group_comparisons_list)
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
        #print(sig_list)
    if CantAnoveList:
        kw_test_sig_list = Get_kwtest(df1)
        #print(kw_test_sig_list)
        ManW_list = [CantAnoveList[i] for i in range(len(kw_test_sig_list)) if kw_test_sig_list[i] < 0.05]
        #print(ManW_list)
        if  ManW_list:
            '''
            ManW_list: 是需要进行满惠特尼分析的变量列表  ['P波(mV)_R1']
            mhtn_list: 是读取excel的获得判断数据列表   
            new_mhtn_list : 是把list嵌套起来反映不同变量  [[0.07580017458236125, 0.00820773610734887, 0.09369261949324824]]
            '''
            #print('进行满惠特尼')
            mhtn_list = Get_manhui(df1)
            new_mhtn_list = [mhtn_list[x:x + group] for x in range(0, len(mhtn_list), group)]
            print(new_mhtn_list[0][1])
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

def xianzhuxin(path_list, out_path):

    '''
    s: 获得实验的时间数
    group：获得实验的小组数
    group_name :组别的名字
    '''
    print('start')
    xlsx_list = [fn for fn in listdir(path_list) if fn.endswith('.xlsx')]
    bb = list(map(lambda y: path_list + '\\' + y, xlsx_list))
    Group_name, group = GetGRoup(bb[0])
    '''
    df1: 读取整个excel
    从df1中读取到时间段
    '''
    for f in bb:
        if '血压' in f:
            title = f[-8:-5]
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            s = getTime(df, '血压')  # 5
            sig_list = judge(df1, group)
            df = easy_clean(df, '血压', sig_list, s,Group_name,group)
            df.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None, header=None)

        if '心电' in f:
            title = f[-8:-5]
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            s = getTime(df, '心电')
            sig_list = judge(df1, group)
            df = Complex_clean(df, '心电', s, sig_list,Group_name,group)
            df.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None, header=None)

        if '呼吸' in f:
            title = f[-8:-5]
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            sig_list = judge(df1, group)
            if '呼吸频率' in df[0][0]:
                s = getTime(df, '呼吸')
                df = Complex_clean(df, '呼吸', s, sig_list,Group_name,group)
            else:
                # 当没有呼吸频率时，呼吸的操作与体温一样,easy_clean ,mean一位,sd两位
                s = getTime(df, '体温')
                df = easy_clean(df, '呼吸', sig_list, s,Group_name,group)
            df.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None, header=None)

        if '体温' in f:
            title = f[-8:-5]
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            sig_list = judge(df1, group)
            s = getTime(df, '体温')
            df = easy_clean(df, '体温', sig_list, s,Group_name,group)
            df.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None, header=None)
    convert_zip(out_path, out_path)


if __name__ == "__main__":
    path_list = r'K:\mashuaifei\死亡进行的显著性\5.6\\'

    xianzhuxin(path_list,path_list)
