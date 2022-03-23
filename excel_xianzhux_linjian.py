import numpy as np
import pandas as pd
from os import listdir
import os
from pandas import Series,DataFrame


def easy_clean(df: pd.DataFrame, name: str, sigList: list, time, Group_name):
    #print(df)
    for i in range(1, df.shape[1], 3):
        if  name == '血液学':
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
        if name == '血生化':
            decimal_1 = [1, 1, 0, 1, 0, 0, 0, 0, 1, 1, 2, 2, 2, 1, 2, 1, 2, 0, 1, 2, 2]
            decimal_1_after = [val for val in decimal_1 for i in range(2)]
            for j in range(df.shape[0]):
                mean_decimal = decimal_1_after[j] + 1
                sd_decimal = decimal_1_after[j] + 2
                df.iloc[j, i] = str(('%.' + str(mean_decimal) + 'f') % df.iloc[j, i]) if df.iloc[j, i] != '' else ''
                df.iloc[j, i + 1] = str(('%.' + str(sd_decimal) + 'f') % df.iloc[j, i + 1]) if df.iloc[j, i + 1] != '' else ''
                df.iloc[j, i] = df.iloc[j, i] + '±' + df.iloc[j, i + 1]

        if name == '凝血' or name == '体重':
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
    #print(df)
    #print(sigList)
    #df = sig_work(df, sigList)
    #print(df)
    return df


def get_df1(df, time, Group_name):
    df_2 = easy_clean1(df, time, Group_name)
    return df_2


def get_df2(df: pd.DataFrame, time):
    for i in range(1, df.shape[1], 3):
        df[i + 2] = df[i + 2].apply(lambda x: 'n=' + str(x) if x != '' else '')
        df.drop(i + 1, axis=1, inplace=True)
    #print(df)
    df = pd.DataFrame(np.repeat(df.values, 2, axis=0))
    #print(df)
    for j in range(1, df.shape[0], 2):
        for i in range(1, df.shape[1], 2):
            df.iloc[j, i] = df.iloc[j, i+1]
    for i in range(2, df.shape[1], 2):
        df.drop(i, axis=1, inplace=True)

    df.insert(0, 'd', df[0])
    if '_' in df[0][1]:
        sep = '_'
        df[0] = df[0].apply(lambda x: x.split(sep, 1)[1])
        df['d'] = df['d'].apply(lambda x: x.split(sep, 1)[0])
    if '-' in df[0][1]:
        sep = '-'
        df[0] = df[0].apply(lambda x: x.split(sep, 1)[1])
        df['d'] = df['d'].apply(lambda x: x.split(sep, 1)[0])

    b = [i for i in range(0, df.shape[0], time*2)] # time 问题 即时间都有两行
    bb = [i for i in range(0, df.shape[0], 2)]
    a = [df['d'][i] if (i) in b else '' for i in range(0, len(df['d']))]
    aa = [df[0][i] if (i) in bb else '' for i in range(0, len(df[0]))]
    df['d'] = a
    df[0] = aa
    #print(df)
    #print(df_2)
    return df


def get_df3(df: pd.DataFrame, Group_name, s):
    print('开始  得到  表格 格式 3')
    print(df)
    try:
        var1 = [x.split('_')[0] for x in list(df[0])]
        var2 = [x.split('_')[1] for x in list(df[0])]
    except:
        var1 = [x.split('-')[0] for x in list(df[0])]
        var2 = [x.split('-')[1] for x in list(df[0])]
    var1_1 = sorted((set(var1)), key=var1.index)
    #var1_2 = var1_1*s

    var2_1 = sorted((set(var2)), key=var2.index)
    print(var2_1)
    var2_2 = sorted(var2, key= var2_1.index)
    #words = [var1_2[i] + '_' + var2_2[i] for i in range(len(var2_2))]
    #print(words)
    '''
    df[0] = df[0].astype('category')
    # inplace = True，使 recorder_categories生效
    df[0].cat.reorder_categories(words, inplace=True)
    # inplace = True，使 df生效
    df.sort_values(0, inplace=True)
    df = df.reset_index(drop=True)
    df[0] = df[0].apply(lambda x : x.split('_')[0])
    '''
    #df = df.T
    df2 = df.reset_index(drop=True)
    df2.drop(0, axis=1, inplace=True)
    df_columns = [j for j in Group_name for i in range(2)]
    #df_columns = [df_columns[j] if j%2 ==0 else df_columns[j]+'1' for j in range(len(df_columns))]
    #print(df_columns)
    # ['安慰剂对照组', '安慰剂对照组1', 'GR1802低剂量组', 'GR1802低剂量组1', 'GR1802中剂量组', 'GR1802中剂量组1', 'GR1802高剂量组', 'GR1802高剂量组1']
    #df_columns1 = [df_columns[i] if i%2 ==0 else '' for i in range(len(df_columns))]
    var2_3 = [str(i) + var2_1[i] for i in range(len(var2_1))]
    var1_2 = [str(i) + var1_1[i] for i in range(len(var1_1))]
    print(var2_3)
    print((var1_2))
    df = DataFrame(df2.values.tolist() ,columns=pd.Index(df_columns,name='zb'), index =pd.MultiIndex.from_product([var1_2,var2_3],names=['1','2']))
    df = df.T


    df = df.stack('2')
    a = (list(df.columns))
    df = df.reset_index()
    df.columns = [ '组别', '时间'] +a

    order = [ '组别', '时间'] +var1_2
    df =df [order]
    grouped_muti = df.groupby(['组别', '时间'])
    #print(grouped_muti.size())
    df1 = grouped_muti.get_group((Group_name[0], var2_3[0]))
    #def merge(df):
    for i in Group_name:
        for j in var2_3:
            df11 = grouped_muti.get_group((i,j))
            df1 =pd.concat([df1,df11],ignore_index=True)
    df1 = df1.drop(0,axis =0)
    df1 = df1.drop(1, axis=0)
    df1.columns = [ '组别', '时间'] + var1_1
    df1['时间'] = df1['时间'].apply(lambda x: x[1:] )

    print(df1)
    '''
    df0 = df.iloc[:, 0:int(df.shape[1] / s)]
    new_name = []
    for jj in range(int(df.shape[1] / s)):
        new_name.append(jj)  # 新索引   [0,1,2,3,4,5]
    for ii in range(1, s):
        df1 = df.iloc[:, (ii) * int(df.shape[1] / s): (ii+1) * int(df.shape[1] / s)]  # [6:12]  [12:18]
        df1.columns = new_name
        df0 = pd.concat([df0, df1], ignore_index=True)
    df = df0
    # 开始处理指标方面
    df_col = ['组别','检测时间'] + var1_1
    var3 = var2_1 * (len(Group_name))
    Group_name = [Group_name[int(i/(s*2))] if i%(s*2)==0 else ''for i in range(df.shape[0])]
    Group_name =  Group_name
    print(Group_name)
    var3 = [i for i in var3 for v in range(2)]
    var3 = [var3[i] if i % 2 == 0 else '' for i in range(len(var3))]
    df.insert(0, 'd', var3)
    df.insert(0, 'a', Group_name)
    df.columns = df_col
    '''
    return df1


def easy_clean1(df: pd.DataFrame, time: int, Group_name):
    print(df)
    if '_' in df[0][0]:
        var = [x.split('_')[0] for x in list(df[0])]
    # elif '-' in df[0][0]:
    #     var = [x.split('-')[0] for x in list(df[0])]
    else:
        var = list(df[0])
    print(var)
    var1 = list(set(var))
    var1.sort(key=var.index)  # 变量列表
    var = [var[v] if v % time == 0 else '' for v in range(len(var))]
    if '_' in df[0][0]:
        times_list1 = [x.split('_')[-1] for x in list(df[0])]
    elif '-' in df[0][0]:
        times_list1 = [x.split('-')[-1] for x in list(df[0])]
    else:
        times_list1 = list(df[0])
    times_list = list(set(times_list1))
    times_list.sort(key=times_list1.index)
    #print(var)
    #print(times_list)

    df = df.drop(0, axis=1)
    #group_n = df.shape[1]

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
    print(Group_name)
    var = [v for v in var1 for i in range(int(len((Group_name)*2)))]
    print(var)
    var = [var[v] if v % (len(Group_name)*2) == 0 else '' for v in range(len(var))]
    #var[0] = '检测指标'
    df = df.reset_index(drop=True)
    # print(df)
    print(var)
    print(len(var))
    #for i in range(1,df.shape[0],3):
        #df.drop(i, axis=0, inplace=True)

    df.insert(0, 'a', var)
    Group_name = [v if i % 2 == 0 else '' for v in Group_name for i in range(2)]
    list_group = Group_name * (int(df.shape[0] / (len(Group_name))))
    df.insert(1, 'b', list_group)
    df_col = ['检测指标', '组别'] + times_list
    print(df)
    print(df_col)
    print('saddddddddddddddddddddddddddddddddddddddddddd')
    df.columns = df_col
    #print(df)
    return df


def GetGRoup(file_list):
    df1 = pd.read_excel(file_list, header=None)
    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'Oneway'].index.tolist()[0] - 5]
    Group_name = list(df[0].dropna(axis=0))[1:]
    if 'Total' in Group_name:
        Group_name.remove('Total')
    group = len(Group_name)-1
    print(Group_name)
    return Group_name, group


def select(df1):
    df = df1[df1[df1[0] == 'Report'].index.tolist()[0] + 1:df1[df1[0] == 'Oneway'].index.tolist()[0] - 5]
    df = df.reset_index(drop= True)
    if 'Total' in list(df[0]):
        print('yes')
        df = df[0:df[df[0] == 'Total'].index.tolist()[0]]
    #print(df)
    df = df.drop([0, 1], axis=1)  # 删除前三列无用列
    df = df.dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    df = df.T
    df = df.reset_index(drop=True)
    df.replace(np.nan, '', inplace=True)
    #print(df)
    return df


def getTime(df):
    list1 = list(df[0])
    #print(list1)
    if '-'in list1[0]:
        s = int(len(list1) / len(set([x.split('-')[0] for x in list1])))
    elif '_' in list1[0]:
        s = int(len(list1) / len(set([x.split('_')[0] for x in list1])))
    else:
        s= int(len(list1))
    #print(s)
    return s


def xianzhuxing1(path_list, out_path):
    print('start')
    xlsx_list = [fn for fn in listdir(path_list) if fn.endswith('.xlsx')]
    bb = list(map(lambda y: path_list + '\\' + y, xlsx_list))
    Group_name, group = GetGRoup(bb[0])
    df_col = ['指标', '时间段']
    df_col = df_col + Group_name
    for f in bb:
        print(f)
        title = f[:-5].split('\\')[-1]
        if '血液学' in f:
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            #print(df)
            s = getTime(df)
            #sig_list = judge(df1, group)
            sig_list = []
            dff = easy_clean(df, '血液学', sig_list, s, Group_name)
            print(dff)
            #df_1 = get_df1(dff, s, Group_name)
            df_2 = get_df2(dff,  s,)
            df_3 = get_df3(df, Group_name, s)
            df_2.columns = df_col

        if '血生化' in f:
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            s = getTime(df)
            #sig_list = judge(df1, group)
            sig_list = []
            dff = easy_clean(df, '血生化', sig_list, s, Group_name)
            df_1 = get_df1(dff, s, Group_name)
            df_2 = get_df2(dff, s)
            df_3 = get_df3(df, Group_name,s )
            df_2.columns = df_col

        elif '凝血' in f or '体重' in f:
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            s = getTime(df)
            #sig_list = judge(df1, group)
            sig_list = []
            dff = easy_clean(df, '凝血', sig_list, s,Group_name)
            df_1 = get_df1(dff, s, Group_name)
            df_2 = get_df2(dff, s)
            df_3 = get_df3(df, Group_name, s)

        elif '脏器重量' in f or '脏脑比' in f:
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            s = getTime(df)
            #sig_list = judge(df1, group)
            sig_list = []
            df_1, df_2 = easy_clean(df, '脏脑比', sig_list, s,Group_name)
            df_1.columns = df_col

        elif '脏体比' in f:
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            s = getTime(df)
            #sig_list = judge(df1, group)
            sig_list = []
            df_1, df_2 = easy_clean(df, '脏体比', sig_list, s,Group_name)
            df_1.columns = df_col

        elif '摄食量' in f:
            df1 = pd.read_excel(f, header=None)
            df = select(df1)
            s = getTime(df)
            #sig_list = judge(df1, group)
            sig_list = []
            df_1,df_2 = easy_clean(df, '摄食量', sig_list, s,Group_name)
            df_1.columns = df_col
            #df_1.to_excel(out_path + '\\' + title + 'ok.xlsx', index=None)
        writer = pd.ExcelWriter(out_path + '\\' + title + 'ok.xlsx')
        #df_1.to_excel(writer, sheet_name='sheet1', index=None)
        df_2.to_excel(writer, sheet_name='sheet2', index=None)
        df_3.to_excel(writer, sheet_name='sheet3', index=None, header= None)
        writer.save()


if __name__ == "__main__":
    path_list = r'K:\mashuaifei\死亡进行的显著性\720苏\\'
    xianzhuxing1(path_list, path_list)


