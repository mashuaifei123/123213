import copy
import numpy as np
import pandas as pd
from os import listdir


def date_del(df):
    df.insert(1, 1_1, df['检测时间'])
    df[1_1].index = df[1_1].index + 1
    df['检测时间'] = df[1_1] + df['检测时间']
    df.drop([1_1], axis=1, inplace=True)
    df_mean = df.loc[[j for j in range(3, df.shape[0], 4)], :]
    # print(df_mean)
    df_mean = df_mean.dropna(subset=['检测时间'])
    df_mean['检测时间'] = df_mean['检测时间'].apply(lambda x: x[:-4])
    df_mean_list = df_mean.iloc[:, 2].tolist()
    time = list(set(df_mean_list))
    time.sort(key=df_mean_list.index)
    return df_mean, time


def Get_df2(df, time, CC):
    PDR_set_l = []
    for x, i in enumerate(time, 1):
        PDR_set_l.append((i, '{:02d}_{}'.format(x, i)))

    def id_to_idn(id, id_list):  # 根据及生成的列表加01- 用于排序
        for id_old in id_list:
            if id == id_old[0]:
                return id_old[1]

    df['检测时间'] = df[:]['检测时间'].map(lambda x: id_to_idn(str(x), PDR_set_l))
    # df.index = [df['动物编号'].tolist(), df['检测时间'].tolist()]
    # df.drop(['检测时间'], axis=1, inplace=True)
    df.sort_values(['动物编号', '动物ID'], inplace=True)
    df['检测时间'] = df['检测时间'].apply(lambda x: x.split('_')[1])
    d = []
    e = []
    for i in list(df['动物编号']):
        if i not in d:
            d.append(i)
        else:
            d.append('')
    for i in list(df['动物ID']):
        if i not in e:
            e.append(i)
        else:
            e.append('')
    df['动物ID'] = e
    df['动物编号'] = d
    if CC == '呼吸':
        df.drop(['最大峰值(g)', '最小谷值(g)'], axis=1, inplace=True)
        df.iloc[:, 3] = df.iloc[:, 3].apply(lambda x: str('%.0f' % (x))).astype(int)
        # df.iloc[:,2] = df.iloc[:, 2].apply(pd.to_numeric, errors='errors')
        df.iloc[:, 4] = df.iloc[:, 4].apply(lambda x: str('%.1f' % (x))).astype(float)
    if  CC =='心电':
        df.iloc[:, 3] = df.iloc[:, 3].apply(lambda x: str('%.0f' % (x))).astype(int)
        for i in range(3,df.shape[1]):
            df.iloc[:, i] = df.iloc[:, i].apply(lambda x: str('%.2f' % (x)))
            df.iloc[:, i] = df.iloc[:, i].apply(pd.to_numeric, errors='errors')
    if CC =='血压':
        for i in range(3,df.shape[1]):
            df.iloc[:, i] = df.iloc[:, i].apply(lambda x: str('%.0f' % (x)))
            df.iloc[:, i] = df.iloc[:, i].apply(pd.to_numeric, errors='errors')
    return df


def date_del_1(df):
    # 确保血压的最后一行中文解释是否删除
    df = df.dropna(axis=0)
    b = list(df['动物编号'])
    e = list(df['试验阶段'])
    ee = [i for i in e if '*' not in i]
    c = [i for i in range(len(b) - 1) if b[i] == b[i + 1]]
    # print(df)
    df.drop(df.index[c], inplace=True)
    df = df.reset_index(drop=True)
    # print(df)
    b = list(df['动物编号'])
    c = [i[0] for i in b]
    d = [i for i in range(len(c) - 1) if c[i] > c[i + 1]]
    # print(d)
    df_time = list(set(ee))
    df_time.remove('mean')
    df_time.sort(key=e.index)
    print(df_time)
    # print(df)
    # print(xueya_time)
    df.loc[0:d[0] + 1, '试验阶段'] = df_time[0]
    df.loc[d[-1] + 1:, '试验阶段'] = df_time[-1]

    for i in range(len(d) - 1):
        df.loc[d[i] + 1:d[i + 1], '试验阶段'] = df_time[i + 1]
    df.rename(columns={'试验阶段': '检测时间'}, inplace=True)
    return df,df_time



def date_del_one(df_mean):
    print(df_mean)
    df_mean_list = df_mean.iloc[:, 1].tolist()
    PDR_set = list(set(df_mean_list))
    PDR_set.sort(key=df_mean_list.index)  # 保留顺序转换列表->去重复集合->列表原始顺序排序
    PDR_set_l = []
    for x, i in enumerate(PDR_set, 1):
        PDR_set_l.append((i, '{:02d}_{}'.format(x, i)))

    def id_to_idn(id, id_list):  # 根据及生成的列表加01- 用于排序
        for id_old in id_list:
            if id == id_old[0]:
                return id_old[1]

    # print(PDR_set)
    # print(PDR_set_l)

    df_mean['检测时间'] = df_mean[:]['检测时间'].map(lambda x: id_to_idn(str(x), PDR_set_l))
    df_mean.index = [df_mean['动物编号'].tolist(), df_mean['检测时间'].tolist()]
    df_mean.drop(['动物编号'], axis=1, inplace=True)
    df_mean.drop(['检测时间'], axis=1, inplace=True)
    print(df_mean)
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
    df_spss_ok['Group'] = df_spss_ok['Group'].apply(pd.to_numeric, errors='errors')
    # print(df_spss_ok.dtypes)
    '''
    for k in range(2+len(PDR_set),df_spss_ok.shape[1]): 
        df_spss_ok.iloc[:,k] = df_spss_ok.iloc[:,k].apply(lambda x :str('%.2f'%x) )
        df_spss_ok.iloc[:,k] = df_spss_ok.iloc[:,k].apply(pd.to_numeric,errors='errors')
    '''
    second_name_list = CO1_name
    time = PDR_set
    # print(second_name_list)
    # print(time)
    return df_spss_ok, second_name_list, time


def decimal(df, time, CC):
    if CC == '心电':
        for k in range(2 + len(time), df.shape[1]):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.2f' % (x)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='errors')

    if CC == '呼吸':
        for k in range(2, 2 + len(time)):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.0f' % (x)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='errors')
        for k1 in range(2 + len(time), df.shape[1]):
            df.iloc[:, k1] = df.iloc[:, k1].apply(lambda x: str('%.1f' % (x)))
            df.iloc[:, k1] = df.iloc[:, k1].apply(pd.to_numeric, errors='errors')

    if CC == '血压':
        for k in range(2, df.shape[1]):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.0f' % (x + 0.000000000001)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='errors')


def Get_df1(df, time, CC):
    if CC == '呼吸':
        df.drop(['动物ID', '最大峰值(g)', '最小谷值(g)'], axis=1, inplace=True)
        df.iloc[:, 2] = df.iloc[:, 2].apply(lambda x: str('%.0f' % (x))).astype(int)
        # df.iloc[:,2] = df.iloc[:, 2].apply(pd.to_numeric, errors='errors')
        df.iloc[:, 3] = df.iloc[:, 3].apply(lambda x: str('%.1f' % (x))).astype(float)
    if CC =='心电':
        df.drop(['动物ID'], axis=1, inplace=True)
        df.iloc[:, 2] = df.iloc[:, 2].apply(lambda x: str('%.0f' % (x))).astype(int)
        for i in range(2,df.shape[1]):
            df.iloc[:, i] = df.iloc[:, i].apply(lambda x: str('%.2f' % (x)))
            df.iloc[:, i] = df.iloc[:, i].apply(pd.to_numeric, errors='errors')
    if CC =='血压':
        df.drop(['动物ID'], axis=1, inplace=True)
        for i in range(2,df.shape[1]):
            df.iloc[:, i] = df.iloc[:, i].apply(lambda x: str('%.0f' % (x)))
            df.iloc[:, i] = df.iloc[:, i].apply(pd.to_numeric, errors='errors')
    return df

def read_excel1(inpath, outpath):
    # 读取路径中的所有excel，放入列表等待处理
    print('start')
    xlsx_list = [fn for fn in listdir(inpath) if fn.endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_list_path = list(map(lambda y: inpath + '\\' + y, xlsx_list))
    print(xlsx_list_path)
    for i in xlsx_list_path:
        try:
            title = i[:-5].split('\\')[-1]
            if '心电' in i:
                xindian_F = pd.read_excel(i, skiprows=1, usecols='A:L')

                a = list(xindian_F.columns)
                df_mean, time = date_del(xindian_F)
                df_mean1 = copy.deepcopy(df_mean)
                df11 = Get_df1(df_mean, time, '心电')
                df22 = Get_df2(df_mean1, time, '心电')

                xindian_M = pd.read_excel(i, skiprows=1, usecols='N:Y')
                xindian_M.columns = a
                df_mean, time = date_del(xindian_M)
                df_mean1 = copy.deepcopy(df_mean)
                df33 = Get_df1(df_mean, time, '心电')
                df44 = Get_df2(df_mean1, time, '心电')

                print('心电finished')

            if '呼吸' in i:
                huxi = pd.read_excel(i, skiprows=1)
                table_length = int((huxi.shape[1] - 1) / 2)
                df_left = huxi.iloc[:, 0:table_length]
                df_right = huxi.iloc[:, table_length + 1:huxi.shape[1]]
                a = list(df_left.columns)

                print(a)
                if '检测时间' in a:
                    df_mean, time = date_del(df_left)
                    df_mean1 = copy.deepcopy(df_mean)
                    df11 = Get_df1(df_mean, time, '呼吸')
                    df22 = Get_df2(df_mean1, time, '呼吸')

                    df_right.columns = a
                    df_mean, time = date_del(df_right)
                    df_mean1 = copy.deepcopy(df_mean)
                    df33 = Get_df1(df_mean, time, '呼吸')
                    df44 = Get_df2(df_mean1, time, '呼吸')

                else:
                    a = a[1:]
                    new_column = ['呼吸_' + x for x in a]
                    new_column.insert(0, '动物编号')
                    df_left.columns = new_column
                    df_right.replace('/', '', inplace=True)
                    df_left.insert(0, 'Group', df_left[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
                    df_left['Group'] = df_left['Group'].apply(pd.to_numeric, errors='errors')
                    df_left.to_excel(outpath + '\\呼吸雄.xlsx', index=None)
                    df_right.columns = new_column
                    df_right.replace('/', '', inplace=True)
                    df_right.insert(0, 'Group', df_right[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
                    df_right['Group'] = df_right['Group'].apply(pd.to_numeric, errors='errors')
                print('呼吸finished')

            if '体温' in i:
                tiwen = pd.read_excel(i, skiprows=1)
                table_length = int((tiwen.shape[1] - 1) / 2)
                df_left = tiwen.iloc[:, 1:table_length - 1]

                df_right = tiwen.iloc[:, table_length + 1:tiwen.shape[1] - 2]
                a = list(df_left.columns[1:])
                new_column = ['体温℃_' + x for x in a]
                new_column.insert(0, '动物编号')
                # print(new_column)
                df_left.columns = new_column
                df_left.replace('/', '', inplace=True)
                df_left.insert(0, 'Group', df_left[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
                df_left['Group'] = df_left['Group'].apply(pd.to_numeric, errors='errors')
                df_left.to_excel(outpath + '\\体温雄.xlsx', index=None)
                df_right.columns = new_column
                df_right.replace('/', '', inplace=True)
                df_right.insert(0, 'Group', df_right[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
                df_right['Group'] = df_right['Group'].apply(pd.to_numeric, errors='errors')
                df_right.to_excel(outpath + '\\体温雌.xlsx', index=None)
                print('体温finished')

            if '血压' in i:
                xueya_F = pd.read_excel(i, skiprows=1, usecols='A:F')
                a = list(xueya_F.columns)
                df_mean, time = date_del_1(xueya_F)
                df_mean1 = copy.deepcopy(df_mean)
                df11 = Get_df1(df_mean, time, '血压')
                df22 = Get_df2(df_mean1, time, '血压')

                xueya_M = pd.read_excel(i, skiprows=1, usecols='H:M')
                xueya_M.columns = a
                df_mean, time = date_del_1(xueya_M)
                df_mean1 = copy.deepcopy(df_mean)
                df33 = Get_df1(df_mean, time, '血压')
                df44 = Get_df2(df_mean1, time, '血压')
                print('血压finished')
            writer = pd.ExcelWriter(outpath + '\\' + title + 'ok.xlsx')
            df11.to_excel(writer, sheet_name='sheet1', index=None)
            df22.to_excel(writer, sheet_name='sheet2', index=None)
            df33.to_excel(writer, sheet_name='sheet3', index=None)
            df44.to_excel(writer, sheet_name='sheet4', index=None)
            writer.save()
        except:
            pass


if __name__ == "__main__":
    in_path = r'K:\mashuaifei\excel之轮回无限\Fw_Fw_B2020008-G1014-01'
    read_excel1(in_path, in_path)
    print('all finished')