import numpy as np
import pandas as pd
from os import listdir
from excel_xianzhux import count_cells,select1
from excel_xianzhux_linjian import get_df1, get_df3
from excel_xianzhux import get_df2
import copy

def get_decimal(df):
    if "a. Limited to first 100 cases." in df[0]:
        index = df[df[0] == "a. Limited to first 100 cases."].index.tolist()
    else:
        data_list1 = df.loc[df[0].str.contains('试验',na =False)].index
        data_list2 = df.loc[df[0].str.contains('QC：',na =False)].index
        # print(data_list1)
        # print(data_list2)
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
        if len(list(set(lines__col_decimal))) == 1:
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
            ccc.pop(ccc.index(max(ccc)))
            decimal = max(ccc)
        else:
            print('数据有问题')
        # print(decimal)
        lines_decimal.append(decimal)
        #lines_decimal = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]
    return lines_decimal



def clean(df, sig_list_ZuJian, sig_list_ZuNei, f, lines_decaimal):
    print(lines_decaimal)
    if '操佳佳' in f:
        for i in range(1, df.shape[1], 3):
            df[i] = df[i].apply(lambda x: str('%.1f' % (x + 0.00000000001)) if x != '' else '')
            df[i + 1]=df[i + 1].apply(
                lambda x: str('%.1f' % (x + 0.00000000001)) if x != '' else '')
    elif '细胞' in f:
        for i in range(1, df.shape[1], 3):
            df[i] = df[i].apply(lambda x: str('%.3f' % (x + 0.00000000001)) if x != '' else '')
            df[i + 1]=df[i + 1].apply(
                lambda x: str('%.4f' % (x + 0.00000000001)) if x != '' else '')
    else:
        for i in range(1, df.shape[1], 3):
            for j in range(df.shape[0]):
                mean_decimal = lines_decaimal[j] + 1
                sd_decimal = lines_decaimal[j] + 2
                df.iloc[j, i] = str(('%.' + str(mean_decimal) + 'f') % (df.iloc[j, i]+0.000001)) if df.iloc[j, i] != '' else ''
                df.iloc[j, i + 1] = str(('%.' + str(sd_decimal) + 'f') % (df.iloc[j, i + 1]+0.000001)) if df.iloc[j, i + 1] != '' else ''
                if df.iloc[j, i + 1] != '':
                    df.iloc[j, i] = df.iloc[j, i] + '±' + df.iloc[j, i + 1]
                else:
                    df.iloc[j, i] = df.iloc[j, i]

    for i in range(1, df.shape[1], 3):
        df[i] = df[i].apply(lambda x: str(x) if x != '' else '') + df[i + 1].apply(
            lambda x: '±' + str(x) if x != '' else '')
    print(df)
    #print('之后')
    if sig_list_ZuJian:
        print('开始组间标注')
        df = sig_work1(df,  sig_list_ZuJian)

    if sig_list_ZuNei:
        print('开始组内标注')
        df = sig_work(df,  sig_list_ZuNei)
    return df




def easy_clean(df: pd.DataFrame,  sig_list_ZuJian, sig_list_ZuNei, time, Group_name, f):
    '''
    Easy_clean
    '''
    #print(sig_list_ZuNei)
    #print(sig_list_ZuJian)
    if '操佳佳' in f:
        for i in range(1, df.shape[1], 3):
            df[i] = df[i].apply(lambda x: str('%.1f' % (x + 0.00000000001)) if x != '' else '')
            df[i + 1]=df[i + 1].apply(
                lambda x: str('%.1f' % (x + 0.00000000001)) if x != '' else '')
    else:
        for i in range(1, df.shape[1], 3):
            df[i] = df[i].apply(lambda x: str('%.2f' % (x + 0.00000000001)) if x != '' else '')
            df[i + 1]=df[i + 1].apply(
                lambda x: str('%.3f' % (x + 0.00000000001)) if x != '' else '')

    for i in range(1, df.shape[1], 3):
        df[i] = df[i].apply(lambda x: str(x) if x != '' else '') + df[i + 1].apply(
            lambda x: '±' + str(x) if x != '' else '')
    print(df)
    if sig_list_ZuJian:
        print('开始组间标注')
        df = sig_work1(df,  sig_list_ZuJian)

    if sig_list_ZuNei:
        print('开始组内标注')
        df = sig_work(df,  sig_list_ZuNei)
    #df.to_excel(r'K:\mashuaifei\dsaok.xlsx', index=None)
    #print(df)
    ''' 
    df = pd.DataFrame(np.repeat(df.values, 2, axis=0))

    for j in range(1, df.shape[0], 2):
        for i in range(1,df.shape[1],2):
            df.iloc[j,i] = df.iloc[j,i+1]
    '''
    #print(df)

    return df

def get_df4(df: pd.DataFrame, time: int) :
    df_n = df[3]

    for i in range(1, df.shape[1], 3):
        df.drop(i + 2, axis=1, inplace=True)
        df.drop( i+ 1,axis=1, inplace=True)
    df.insert(1, 'a', df_n)
    df['a'] = df['a'].apply(lambda x : 'n=' + str(x))
    df.insert(0, 'd', df[0])
    if '_' in df[0][1]:
        sep = '_'
        df[0] = df[0].apply(lambda x: x.split(sep, 1)[1])
        df['d'] = df['d'].apply(lambda x: x.split(sep, 1)[0])
    b = [i for i in range(0, df.shape[0], time)]  # time 问题 即时间都有两行
    bb = [i for i in range(0, df.shape[0])]
    a = [df['d'][i] if (i) in b else '' for i in range(0, len(df['d']))]
    aa = [df[0][i] if (i) in bb else '' for i in range(0, len(df[0]))]
    df['d'] = a
    df[0] = aa
    #print(df)
    return df


def AddIndex(df: pd.DataFrame, time: int, Group_name):
    #print(df)
    var = [x.split('_')[0] for x in list(df[0])]
    var1 = list(set(var))
    var1.sort(key=var.index)  # 变量列表
    var = [var[v] if v % time == 0 else '' for v in range(len(var))]
    times_list1 = [x.split('_')[-1] for x in list(df[0])]
    times_list = list(set(times_list1))
    times_list.sort(key=times_list1.index)
    # print(var)
    # print(times_list)
    df.drop(0, axis=1, inplace=True)
    # group_n = df.shape[1]

    df = df.T
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
    # 拼接完成
    var = [v for v in var1 for i in range(int(len((Group_name)*2)))]
    var = [var[v] if v % (len(Group_name)*2) == 0 else '' for v in range(len(var))]
    # var[0] = '检测指标'
   # print(df)
    #print(var)
    df.insert(0, 'a', var)
    Group_name = [v if i % 2 == 0 else '' for v in Group_name for i in range(2)]
    list_group = Group_name * (int(df.shape[0] / (len(Group_name))))
    # print(list_group)
    # print(len(list_group))
    # print(len(var))
    #print(list_group)
    #print(df)
    df.insert(1, 'b', list_group)

    df_col = ['检测指标', '组别'] + times_list
    df.columns = df_col
    # print(df)
    return df


def AddIndex1(df, time, Group_name):
    # 删除不需要的列
    for i in range(2, df.shape[1], 3):
        df = df.drop(i, axis=1)
        df = df.drop(i+1, axis=1)
    df.insert(0, 'd', df[0])

    if '_' in df[0][1]:
        sep = '_'
        df[0] = df[0].apply(lambda x: x.split(sep, 1)[1])
        df['d'] = df['d'].apply(lambda x: x.split(sep, 1)[0])

    b = [i for i in range(0, df.shape[0], time )]  # time 问题 即时间都有两行
    #bb = [i for i in range(0, df.shape[0], 2)]
    a = [df['d'][i] if (i) in b else '' for i in range(0, len(df['d']))]
    #aa = [df[0][i] if (i) in bb else '' for i in range(0, len(df[0]))]
    df['d'] = a
    df_col = ['检测指标','时间'] +Group_name
    # print(a)
    # print(aa)
    df.columns = df_col
    #print(df)
    #df[0] = aa
    # print(df)
    return df


def sig_work(df, sigList):
    # 开始加*操作的前奏，df排序
    print('开始标注')
    '''
    #print(df)
    sort_key = list(df[0])
    by = list(df[3])
    # 排序df，先把大于3进行oneway的指标放在前面，不做的放在后面，并以此进行排序
    sort_key1 = [sort_key[m] for m in range(len(by)) if by[m] != '' and by[m] >= 3]

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
    #print(sigList)
    '''
    if sigList:
        for j in sigList:
            if j[2] == 1:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '△'
            else:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '△△'
    # 标*结束，之后返回顺序
    #df[0] = df[0].astype('category')
    #df[0].cat.reorder_categories(sort_key, inplace=True)
    #df.sort_values(by=0, inplace=True)
    #df = df.reset_index(drop=True)
    # print(df)
    #df.to_excel('K:\mashuaifei\死亡进行的显著性\dasdasok.xlsx', index=None)
    return df


def sig_work1(df, sigList):
    print(df)
    print(df.iloc[1,:])
    if sigList:
        for j in sigList:
            if j[2] == 1:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '*'
            else:
                df.iloc[j[0], j[1]] = df.iloc[j[0], j[1]] + '**'
    return df


def GetGRoup(file_list):
    '''
    :param file_list: 随机挑选一份文件
    :return:  本次试验的组别列表 , 和组数
    for forample : ['溶媒对照组', 'TL139低剂量组', 'TL139中剂量组', 'TL139高剂量组'] ,则group = 3
    '''
    df1 = pd.read_excel(file_list, header=None)

    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'General Linear Model'].index.tolist()[0] - 15]
    Group_name = list(df[0].dropna(axis=0))[1:]
    print(df)
    if 'Total' in Group_name:
        Group_name.remove('Total')
    group = len(Group_name)-1
    print(Group_name)
    return Group_name, group


def select(df1, Group_name):
    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'General Linear Model'].index.tolist()[0] - 8]
    df = df.iloc[:3*len(Group_name)+1, :]  # 截取正确行数
    df = df.drop([0, 1], axis=1)  # 删除前三列无用列
    df = df.dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    df = df.T
    df = df.reset_index(drop=True)
    df.replace(np.nan, '', inplace=True)

    return df


def getTime(df):
    list1 = list(df[0])
    list1 = [x.split('_')[0] for x in list1]

    zhibiao = len(set(list1))
    s = int(len(list1) / len(set([x.split('_')[0] for x in list1])))
    print(s)
    return s,zhibiao


def Get_MTS1(df1, f, get, zhibiao):
    #####  重点标记这里第一次球形检验的结果,提取的位置直接决定了之后的所有操作
    if zhibiao == 1 :
        list1 = list(df1[df1[df1[0] == "Measure: "].index.tolist()[0]:df1[df1[0] == "Measure: "].index.tolist()[0]+1][1])
        df = df1[df1[df1[0] == "Mauchly's Test of Sphericitya"].index.tolist()[0] + 4 \
                :df1[df1[0] == "b. May be used to adjust the degrees of freedom for the averaged tests of significance. Corrected tests are displayed in the Tests of Within-Subjects Effects table."].index.tolist()[0] - 2]
        df = df.dropna(axis=1, how='all')
        mts1_sig = list(df[4])
        print(list1)
    else:
        df = df1[df1[df1[0] == "Mauchly's Test of Sphericitya"].index.tolist()[0] + 3 \
                :df1[df1[0] == "b. May be used to adjust the degrees of freedom for the averaged tests of significance. Corrected tests are displayed in the Tests of Within-Subjects Effects table."].index.tolist()[0] - 2]
        #df = df.dropna(axis=1, how='all')
        mts1_sig = list(df[5])
        list1 = list(df[1])
    return mts1_sig, list1

def get_get(df1):
    get = df1.iloc[1,:][0]
    return get

def Get_UT1(df1, list1):
    if len(list1)>1:
        df = df1[df1[df1[0] == "Univariate Tests"].index.tolist()[0] + 2 \
                 :df1[df1[0] == "Tests of Within-Subjects Contrasts"].index.tolist()[0] - 1]
        df = df.dropna(axis=1, how='all')
        tu1_sig = list(df[7])[:4 * len(list1)]
        tu1_sig = [tu1_sig[x:x + 4] for x in range(0, len(tu1_sig), 4)]
    else :
        df = df1[df1[df1[0] == "Tests of Within-Subjects Effects"].index.tolist()[0] + 3 \
                 :df1[df1[0] == "Tests of Within-Subjects Contrasts"].index.tolist()[0] - 1]
        df = df.dropna(axis=1, how='all')
        #print(df)
        tu1_sig = list(df[6])[:4 * len(list1)]
        tu1_sig = [tu1_sig[x:x + 4] for x in range(0, len(tu1_sig), 4)]

    return tu1_sig
    # print(df)


def Get_ZuNei_sig(df1, Group_name, Zunei_list):
    index_a = df1[df1[0] == "Mauchly's Test of Sphericitya,b"].index.tolist()
    index_b = df1[df1[
                      0] == "Tests the null hypothesis that the error covariance matrix of the orthonormalized transformed dependent variables is proportional to an identity matrix."].index.tolist()
    index_b = index_b[1:len(Group_name) + 1]

    group_mts_sig_total = []
    group_mts_GG_total = []
    print(index_a)
    print(index_b)
    print(Group_name)
    for i in range(len(Group_name)):
        if len(Zunei_list) == 1:
            df = df1[index_a[i] + 4: index_b[i]]
            df = df.dropna(axis=1, how='all')
            print(df)
            if '4' in df.columns:
                group_mts_sig = list(df[4])
            else:
                group_mts_sig = []
            group_mts_GG = list(df[5])
            group_mts_sig_total.append(group_mts_sig)
            group_mts_GG_total.append(group_mts_GG)
        else:

            df = df1[index_a[i] + 3: index_b[i]]
            df = df.dropna(axis=1, how='all')
            if '5' in df.columns:
                group_mts_sig = list(df[5])
            else:
                group_mts_sig = []
            group_mts_GG = list(df[6])
            group_mts_sig_total.append(group_mts_sig)
            group_mts_GG_total.append(group_mts_GG)

    # print(df)
    # print(group_mts_sig_total)
    # print(group_mts_GG_total)

    return group_mts_sig_total, group_mts_GG_total


def Get_ZuNei_sig2(df1, Group_name,Zunei_list):
    if len(Zunei_list)==1:
        index_a = df1[df1[0] == "Tests of Within-Subjects Effectsa"].index.tolist()
        index_b = df1[df1[0] == "Tests of Within-Subjects Contrastsa"].index.tolist()
        group_ut_total = []
        print(index_a)
        print(index_b)
        print(Group_name)
        for i in range(len(Group_name)):
            df = df1[index_a[i] + 3: index_b[i]-2]
            df = df.dropna(axis=1, how='all')
            print (df)
            df = list(df[6].dropna(axis=0, how='all'))

            ut_sig = [df[x:x + 4] for x in range(0, len(df), 4)]
            #print(ut_sig)
            group_ut_total.append(ut_sig)
    else :
        index_a = df1[df1[0] == "Univariate Testsa"].index.tolist()
        index_b = df1[df1[0] == "Tests of Within-Subjects Contrastsa"].index.tolist()
        group_ut_total = []
        #print(index_a)
        #print(index_b)
        for i in range(len(Group_name)):
            df = df1[index_a[i] + 2: index_b[i]]
            df = df.dropna(axis=1, how='all')
            df = list(df[7].dropna(axis=0, how='all'))
            ut_sig = [df[x:x + 4] for x in range(0, len(df), 4)]
            group_ut_total.append(ut_sig)

    return group_ut_total
    # 三层嵌套组别〉 指标〉 四个


def Get_Zunei_list(df1, Group_name, time,LIST):
    index_a = df1[df1[0] == "Pairwise Comparisonsa"].index.tolist()
    index_b = df1[df1[0] == "Based on estimated marginal means"].index.tolist()
    #print(index_a)
    #print(index_b)
    if index_a:
        df = df1[index_a[0] + 3: index_b[0]]
        #print(df)
        df_index =list(df[0].dropna())
    else:
        return []
    #print(df_index)
    return df_index

def Get_ZuNei_pc(df1, Group_name, time,number):
    index_a = df1[df1[0] == "Pairwise Comparisonsa"].index.tolist()
    index_b = df1[df1[0] == "Based on estimated marginal means"].index.tolist()
    pw_total1 = []
    pw_total = []
    print(index_a)
    print(index_b)
    for i in range(0,len(Group_name)):
        if number ==1:
            df = df1[index_a[i] + 4: index_b[i]]
            df = df.dropna(axis=1, how='all')
            df_index =list(df[0].dropna())
            print(df)
            df = list(df[4])
            for j in range(number):
                pw_list = df[j * time * (time - 1):j * time * (time - 1) + time - 1]
                pw_total1.append(pw_list)

        else:
            df = df1[index_a[i] + 3: index_b[i]]
            df = df.dropna(axis=1, how='all')
            df_index =list(df[0].dropna())
            df = list(df[5])
            for j in range(number):
                pw_list = df[j * time * (time - 1):j * time * (time - 1) + time - 1]
                pw_total1.append(pw_list)

    pw_total = [pw_total1[x:x + number] for x in range(0, len(pw_total1), number)]
    #print(pw_total)
    #print(df_index)
    return pw_total

'''
组间的判断

'''
def Get_ZuJian_sig(df1, list1):

    index_a = df1[df1[0] == "Tests of Between-Subjects Effects"].index.tolist()
    index_b = df1[df1[0] == "General Linear Model"].index.tolist()
    #print(index_a)
    #print(index_b)
    if len(list1) > 1:
        if len(index_b) > 1:
            df = df1[index_a[0] + 3: index_b[1] - 2]
        else:
            df = df1[index_a[0] + 3: ]
    else:
        if len(index_b) > 1:
            df = df1[index_a[0] + 4: index_b[1] - 2]
        else:
            df = df1[index_a[0] + 4: ]

    df = df.dropna(axis=1, how='all')
    if len(list1)>1:
        df_sig = df[6]
    else:
        df_sig = df[5]
    zujian_sig = list(df_sig[len(list1): len(list1) * 2])
    return zujian_sig


def Get_ZuJian_time_sig(df1, ZuJian_More_list, time):
    index_a = df1[df1[0] == "Tests of Between-Subjects Effects"].index.tolist()
    index_b = df1[df1[0] == "Post Hoc Tests"].index.tolist()
    df = df1[index_a[1] + 2: index_b[0]]
    df = df.dropna(axis=1, how='all')
    df = df[6]
    # time实际进行分析的数量 ############################################################
    df_sig = list(df[2 * len(ZuJian_More_list) * time: 2 * len(ZuJian_More_list) * time + len(ZuJian_More_list) * time])
    #print('dasd')
    #print(df_sig)
    df_sig = [df_sig[x*time:x*time + time] for x in range(0, len(ZuJian_More_list))]
    #print(df_sig)
    return df_sig


def Get_ZuJian_pc(df1, Group_name, time, ZuNei_list):
    index_a = df1[df1[0] == "Post Hoc Tests"].index.tolist()
    index_b = df1[df1[0] == "*. The mean difference is significant at the .05 level."].index.tolist()
    if index_b == []:
        index_b = df1[df1[0] == "Output Export"].index.tolist()
    print(index_a)
    print(index_b)
    df = df1[index_a[0] + 9: index_b[0]]
    df = list(df[5])
    #print(df)
    df_sig_total = []
    # print(list(df))
    for i in range(len(ZuNei_list) * time):
        # print(i*len(Group_name)*3)
        df_sig = df[i * len(Group_name) * 3: i * len(Group_name) * 3 + 3]

        # print(df_sig)
        df_sig_total.append(df_sig)
    df_sig_total = [df_sig_total[x*time:x*time + time] for x in range(0, len(ZuNei_list))]
    return df_sig_total


def judge(df1, Group_name, time, f, get, zhibiao):
    sig_list_ZuNei = []
    sig_list_ZuJian = []

    mts1_sig, list1 = Get_MTS1(df1, f, get, zhibiao)
    # list1 所有指标的列表
    print(mts1_sig)   # 球形检验的sig值列表
    print(list1)    # ['RR', 'HR', 'PR', 'QRS', 'QT', 'QTcB']

    #组间分析
    zujian_sig = Get_ZuJian_sig(df1, list1)
    print(zujian_sig)  # 组间判断sig值列表
    ZuJian_Less_list_locate = [i for i, j in enumerate(zujian_sig) if j < 0.05]
    ZuJian_Less_list = [list1[i] for i, j in enumerate(zujian_sig) if j < 0.05]
    # 列表示需要组间的变量
    # 列表是该元素在总元素的位置[4,5]
    print(ZuJian_Less_list)
    if ZuJian_Less_list_locate :
        print('开始组间分析')
        # 获得组建的Group数据表，两层 总-指标
        zujian_time_sig = Get_ZuJian_time_sig(df1, ZuJian_Less_list_locate, time)
        print(zujian_time_sig)
        zujian_time = []
        for i in range(len(zujian_time_sig)):
            for j in range(len(zujian_time_sig[i])):
                if zujian_time_sig[i][j] < 0.05:
                    zujian_time.append([i, j])

        print(zujian_time)
        # 列表示在该2*time的列表中位置 [[0, 1], [0, 3], [1, 0], [1, 2]]   0代表在这几个组间变量的第一个,后面的1代表时间的第二个
        if zujian_time:
            zujian_pc_sig_list = Get_ZuJian_pc(df1, Group_name, time, ZuJian_Less_list_locate)
            print(zujian_pc_sig_list)  #没问题
            for ss,i in enumerate(zujian_time):
                #print(i)
                #print(zujian_pc_sig_list[i[0]][i[1]])
                for j, k in enumerate(zujian_pc_sig_list[i[0]][i[1]]):
                    if 0.01<k < 0.05:
                        sig_list_ZuJian.append([list1.index(ZuJian_Less_list[i[0]])*time + i[1], (j+1)*3+1,1])
                    if k < 0.01:
                        sig_list_ZuJian.append([list1.index(ZuJian_Less_list[i[0]]) * time + i[1], (j + 1) * 3 + 1, 2])

                        #print((j+1)*3+1)
                        #print((i[0]+1)*time -1)
                        #print(list1.index(ZuJian_Less_list[i[0]])*time + i[1] +1)

                        #print('组间存在差异')

        print(sig_list_ZuJian)
    else:
        print("没有组间分析")

    MoreThanList = [list1[i] for i in range(len(list1)) if mts1_sig[i] > 0.05]  # ['RR', 'HR', 'PR', 'QRS', 'QT']
    LessThanList = [list1[i] for i in range(len(list1)) if mts1_sig[i] < 0.05]
    MoreThanList_locate = [i for i in range(len(list1)) if mts1_sig[i] > 0.05]  # 大于0.05的变量位置 [0, 1, 2, 3, 4]
    LessThanList_locate = [i for i in range(len(list1)) if mts1_sig[i] < 0.05]  # 小于0.05的变量位置 [5]
    print(MoreThanList)
    print(MoreThanList_locate)

    tu1_sig = Get_UT1(df1, list1)  # UT表的总值,sig值总列表
    # 组内判断
    #print(tu1_sig)
    ZuNei_list = []
    if MoreThanList_locate:
        print('开始大于部分组内分析')
        More_tu1_sig = [tu1_sig[i] for i in MoreThanList_locate]  # UT表中大于0.05值
        for i in range(len(More_tu1_sig)):
            if More_tu1_sig[i][0] < 0.05:
                ZuNei_list.append(MoreThanList[i])

    if LessThanList_locate:
        print('开始小于部分组内分析')
        Less_tu1_sig = [tu1_sig[i] for i in LessThanList_locate]  # UT表中小于0.05值
        for i in range(len(Less_tu1_sig)):
            if Less_tu1_sig[i][1] < 0.05:
                ZuNei_list.append(LessThanList[i])
    ZuNei_list1= []
    for i in ZuNei_list:
        for j, k in enumerate(list1):
            if i == k:
                ZuNei_list1.append(j)
    ZuNei_list1.sort()
    #print(list1)
    #print(ZuNei_list1)
    ZuNei_list = [list1[i] for i in ZuNei_list1]
    number_ZuNei = len(ZuNei_list)
    print(ZuNei_list)
    if ZuNei_list:
        print('判断结束，有需要组内分析的变量')

        #print(ZuNei_list)
        group_mts_sig_total, group_mts_GG_total = Get_ZuNei_sig(df1, Group_name,ZuNei_list)
        #print(group_mts_sig_total)
        #print(group_mts_GG_total)
        # 球形检验sig列和绿房子列  都是2层 组别-指标
        group_ut_total = Get_ZuNei_sig2(df1, Group_name, ZuNei_list)
        print(group_ut_total)
        '''
        #print(group_ut_total)
        # UT表格对应数据表格,三层  组别-指标-分级
        pw_total = Get_ZuNei_pw(df1, Group_name, time, number_ZuNei)
        # pw表格两两比较表格 ，两层 组别，指标,每个指标时间减一个数

        for m in range(len(group_ut_total)):
            # 一：循环组别 m：1 2 3 4
            if group_mts_sig_total[m]:
                #print('以sig列进行判断')
                # 二：判断第一组，各组的sig是否存在
                # for i in range(len(group_mts_sig_total[m])):
                # 三：若存在 ,循环每个组别的每个指标1 2 3 4  ，判断是否小于0.05，将小于的指标位置放入列表
                group_less_than_list_locate = [i for i, j in enumerate(group_mts_sig_total[m]) if j < 0.05]
                if group_less_than_list_locate:
                    # sig中有小于0.05的值位置
                    # 四：循环这些小于0.05的位置，找到对应的UT表值与0.05比较，小于的记录下新位置
                    group_less_than_ut = [group_ut_total[m][h] for h in group_less_than_list_locate if
                                          group_ut_total[m][h][0] < 0.05]
                    if group_less_than_ut:
                        # 五：如果有指标还是小于0.05，则说明确实有差异，调用最后一张表，看到到底是哪个时间段有问题
                        # [m,s,ss+1] m那一组，s 该指标在zunei_list的位置，即该位置在组内列表的位置的便利在总指标便量的位置，ss+1 第几个时间段
                        for s in group_less_than_ut:
                            for ss in range(len(pw_total[m][s])):
                                if pw_total[m][s][ss] < 0.05:
                                    sig_list_ZuNei.append([list1.index(ZuJian_Less_list[i[0]])*time + i[1], (j+1)*3+1,1])
                    #print(sig_list_ZuJian)

            else:
                #print('没有sig，直接以大表格GG行进行判断')
                # 如果sig不存在，以GG列为准
                #print(group_mts_GG_total[m])
                #group_less_than_list_locate = [i for i, j in enumerate(group_mts_GG_total[m]) if j < 0.05]
                #print(group_less_than_list_locate)
                #if group_less_than_list_locate:
                #print('GG列有小于0.05的值，需要看下一个表GG行的值')
                #group_less_than_ut = [group_ut_total[m][h] for h in group_less_than_list_locate if
                #                     group_ut_total[m][h][1] < 0.05]
    
                if group_less_than_ut:
                    print('GG行也有小于0.05的值，可以看对比表找出差异时间了')
                    for s in group_less_than_ut:
                        for ss in range(len(pw_total[m][s])):
                            if pw_total[m][s][ss] < 0.05:
                                sig_list_ZuNei.append([m, list1.index(ZuNei_list[s]), ss + 1])
                                print('组内有差异')
                else:
                    print('组内没有差异')
            '''
        if len(ZuNei_list) ==1:
            ZuNei_list_1 = ZuNei_list
        else :
            ZuNei_list_1 = Get_Zunei_list(df1, Group_name, time,ZuNei_list)
        if ZuNei_list_1:
            #print(ZuNei_list_1)
            number = len(ZuNei_list_1)
            Pariwise_C = Get_ZuNei_pc(df1, Group_name, time, number)
            #print(Pariwise_C)
            list_1 = []

        for i,j in enumerate(group_ut_total):
            for n,m in enumerate(j):
                #print(m)
                #print(n)
                if m[1]<0.05:
                    for x, k in enumerate(Pariwise_C[i][n]):
                        #print( Pariwise_C[i][n])
                        if k < 0.01:
                            print(x)
                            print(list1.index(ZuNei_list_1[n]))
                            list_1.append([list1.index(ZuNei_list_1[n]) * time + x +1, (i + 1) * 3 - 2, 2])
                        if 0.01 < k < 0.05:
                            list_1.append([list1.index(ZuNei_list_1[n]) * time + x +1, (i + 1) * 3 - 2, 1])
                        #print(list_1)
        sig_list_ZuNei = list_1

            # 第一次循环 几个组
            #第二次循环，组中的几个指标
            # 第三次循环，指标中的对应值
           # for i, j in enumerate(Pariwise_C):
              #  for m, n in enumerate(j):
                   # for x, y in enumerate(n):
                        #if y < 0.01:
                        #    list_1.append([list1.index(ZuNei_list_1[m])*time+x+1, (i+1)*3-2, 2])
                        #if 0.01 < y < 0.05 :
                            #print(y)
                            #print(x)
                            #print(m)
                            #print(ZuNei_list_1[m])
                            #list_1.append([list1.index(ZuNei_list_1[m])*time+x+1, (i+1)*3-2, 1])
        print(list_1)

    else:
        print('没有进行组内判断分析')

    print(sig_list_ZuJian)
    return sig_list_ZuJian,sig_list_ZuNei

def xianzhuxing3(path_list, out_path):
    '''
    s: 获得实验的时间数
    group：获得实验的小组数
    group_name :组别的名字
    '''
    #print('start')
    xlsx_list = [fn for fn in listdir(path_list) if fn.endswith('.xlsx')]
    bb = list(map(lambda y: path_list + '\\' + y, xlsx_list))

    # df_col = ['检测指标', '组别']
    # df_col= df_col + Group_name

    for f in bb:
        # if any(each in f  for each in ['血压','RIP']):
        title = f[:-5].split('\\')[-1]
        print(f)
        df1 = pd.read_excel(f, header=None)
        df1_decimal = get_decimal(df1)
        if 'Oneway' not in list(df1[0]):
            Group_name, group = GetGRoup(f)
            df = select(df1,Group_name)
            df_col = ['指标', '时间段']
            df_col = df_col + Group_name
            get = get_get(df1)
            s,zhibiao = getTime(df)
            print(s)
            print(zhibiao)
            #s = s - 1
            sig_list_ZuJian, sig_list_ZuNei = judge(df1, Group_name, s, f,get, zhibiao)
            df_initial = select1(df1)
            df = clean(df, sig_list_ZuJian, sig_list_ZuNei, f, df1_decimal)
            df_4 = copy.copy(df)
            df11 = get_df2(df, s)
            df11.columns = df_col
            df22 = get_df1(df, s, Group_name)
            df33 = get_df3(df, Group_name, s)
            df44 = get_df4(df_4, s)
            print(Group_name)
            df44.columns = ['检测指标', '试验阶段', 'N'] + Group_name
            writer = pd.ExcelWriter(out_path + '\\' + title + 'ok.xlsx')
            df22.to_excel(writer, sheet_name='sheet1', index=None)
            df11.to_excel(writer, sheet_name='sheet2', index=None)
            df33.to_excel(writer, sheet_name='sheet3', index=None)
            df44.to_excel(writer, sheet_name='sheet4', index=None)
            writer.save()
            # if '细胞' in f:
            #     df_1, df_2 = count_cells(df_initial, path_list, title, s)
            #     writer1 =pd.ExcelWriter(out_path + '\\' + title + '检出率.xlsx')
            #     df_1.to_excel(writer1, sheet_name='sheet1', index=None)
            #     df_2.to_excel(writer1, sheet_name='sheet2', index=None)
            #     writer1.save()
            #
            #     print('end')

if __name__ == "__main__":
    path_list = r'K:\mashuaifei\死亡进行的显著性\11-15\操佳佳\\'
    df = xianzhuxing3(path_list, path_list)
