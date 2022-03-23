import numpy as np
import pandas as pd
from os import listdir
import os
import re
import zipfile
import copy

def date_del(df):

    df.insert(1, 1_1, df['检测时间'])
    df[1_1].index = df[1_1].index + 1
    df['检测时间'] = df[1_1] + df['检测时间']
    df.drop([1_1], axis=1, inplace=True)
    df_mean = df.loc[[j for j in range(3, df.shape[0], 4)], :]
    print(df_mean)
    df_mean = df_mean.dropna(subset=['检测时间'])
    df_mean['检测时间'] = df_mean['检测时间'].apply(lambda x: x[:-4])
    #print(df_mean)
    df_spss_ok, second_name_list, time, df_mean1 = date_del_one(df_mean)
    return df_spss_ok, second_name_list, time, df_mean1


def date_del_1(df):
    # 确保血压的最后一行中文解释是否删除

    df = df.dropna(axis=0)
    b = list(df['动物编号'])
    # print(b)
    e = list(df['试验阶段'])
    ee = [i for i in e if '*' not in i]
    c = [i for i in range(len(b) - 1) if b[i] == b[i + 1]]
    df = df.drop(df.index[c])
    df = df.reset_index(drop=True)
    b1 = list(df['动物编号'])
    # print('dasdasda')
    c = [i[0] for i in b1]
    d = [i for i in range(len(c) - 1) if c[i] > c[i + 1]]
    df_time = list(set(ee))
    df_time.remove('mean')
    df_time.sort(key=e.index)
    # print(df_time)
    # print(df)
    # print(xueya_time)
    df.loc[0:d[0] + 1, '试验阶段'] = df_time[0]
    df.loc[d[-1] + 1:, '试验阶段'] = df_time[-1]

    for i in range(len(d) - 1):
        df.loc[d[i] + 1:d[i + 1], '试验阶段'] = df_time[i + 1]
    df.rename(columns={'试验阶段': '检测时间'}, inplace=True)
    return df


def date_del_one(df_mean):
    #print(df_mean)
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

    df_mean1 = copy.deepcopy(df_mean)
    #print(df_mean1)
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
    df_spss_ok['Group'] = df_spss_ok['Group'].apply(pd.to_numeric, errors='coerce')
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
    return df_spss_ok, second_name_list, time, df_mean1


def decimal(df, time, CC):
    if CC == '心电':
        for k in range(2 + len(time), df.shape[1]):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.2f' % (x)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='coerce')

    if CC == '呼吸':
        print(df)
        for k in range(2, 2 + len(time)):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.0f' % (x)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='coerce')
        for k1 in range(2 + len(time), df.shape[1]):
            df.iloc[:, k1] = df.iloc[:, k1].apply(lambda x: str('%.1f' % (x)))
            df.iloc[:, k1] = df.iloc[:, k1].apply(pd.to_numeric, errors='coerce')

    if CC == '血压':
        for k in range(2, df.shape[1]):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.0f' % (x + 0.000000000001)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='coerce')

def decimal_1(df, time, CC):
    if CC == '心电':
        for k in range(3, df.shape[1]):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.2f' % (x)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='coerce')

    if CC == '呼吸':
        df.iloc[:, 2] = df.iloc[:, 2].apply(lambda x: str('%.0f' % (x)))
        df.iloc[:, 2] = df.iloc[:, 2].apply(pd.to_numeric, errors='coerce')
        for k in range(3, df.shape[1]):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.1f' % (x)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='coerce')

    if CC == '血压':
        for k in range(2, df.shape[1]):
            df.iloc[:, k] = df.iloc[:, k].apply(lambda x: str('%.0f' % (x + 0.000000000001)))
            df.iloc[:, k] = df.iloc[:, k].apply(pd.to_numeric, errors='coerce')




def split_excel(df,out_path, time:list, name:list,  first_name, sex,df1,df2):
    # 开始拆分excel
    print(df)
    if first_name =='体温':
        print(out_path + '\\' + first_name + sex + '.xlsx')
        df.to_excel(out_path + '\\' + first_name + sex + '.xlsx', index=None)
    else:
        name = [i.replace('（', '(') for i in name]
        name = [i.replace('）', ')') for i in name]
        name = [re.sub(u"\\(.*?\\)|\\{.*?}|\\[.*?]", "", i) for i in name]
        print(out_path + '\\' + first_name + sex + '.xlsx')
        #df.to_excel(out_path + '\\' + first_name + sex + '.xlsx', index=None)
        #print(int((df.shape[1]-2)/len(time)))
        writer1 = pd.ExcelWriter(out_path + '\\' + first_name + sex + '.xlsx')
        df .to_excel(writer1, sheet_name='spss格式', index=None)
        df1.to_excel(writer1, sheet_name='正常格式1', index=None)
        df2.to_excel(writer1, sheet_name='正常格式2', index=None)
        # 确定固定格式
        length = df.shape[0]
        df_solid = df.iloc[:, 0:2]
        df_leader = df.iloc[:,:2+len(time)]
        df_index = df.columns.tolist()[2:]
        df_index1 = [i.split('_')[0] for i in df_index]
        df_index2 = list(set(df_index1))
        df_index2.sort(key = df_index1.index)
        print(df_index2)

        df_leader.columns = ['Group','动物编号'] + time
        for i in range(1,int((df.shape[1]-2)/len(time))):
            df_concat = df.iloc[0:, i*len(time)+2:i*len(time)+len(time)+2]
            df_concat1  = pd.concat([df_solid, df_concat], axis=1) # 添上前缀group和编号
            df_concat1.columns = ['Group', '动物编号'] + time
            df_leader = pd.concat([df_leader,df_concat1],axis=0,ignore_index=True)
        df_index = [df_index2[int(i/length)] if i% int(length) ==0 else '' for i in range(df_leader.shape[0])]
        df_leader['Group'] = df_index
        df_leader =  df_leader.rename(columns = {'Group':'指标'})
        df_leader.to_excel(writer1, sheet_name='正常格式3', index=None)

        for i in range(int((df.shape[1]-2)/len(time))):
            df1 = df.iloc[:, i*len(time)+2:i*len(time)+len(time)+2]
            df2 = pd.concat([df_solid, df1], axis=1)
            #df2.to_excel(writer1, sheet_name='sheet'+str(i+2), index=None)
            df2.to_excel(writer1, sheet_name=name[i], index=None)
        writer1.save()


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


def read_excel(inpath, outpath):
    # 读取路径中的所有excel，放入列表等待处理
    print('start')
    xlsx_list = [fn for fn in listdir(inpath) if fn.endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_list_path = list(map(lambda y: inpath + '\\'+y, xlsx_list))
    print(xlsx_list_path)
    for i in xlsx_list_path:

        if '心电' in i:
            xindian_F = pd.read_excel(i, skiprows=1, usecols='B:L')
            a = list(xindian_F.columns)
            xindian_F_ok, second_name_list, time, xindian_F1 = date_del(xindian_F)
            decimal(xindian_F_ok, time, '心电')
            decimal_1(xindian_F1, time, '心电')
            xindian_F2 = copy.deepcopy(xindian_F1)
            xindian_F1['检测时间'] = xindian_F1['检测时间'].apply(lambda x: x[3:])

            xindian_F2.sort_values(['动物编号', '检测时间'], inplace=True)
            xindian_F2['检测时间'] = xindian_F2['检测时间'].apply(lambda x: x[3:])
            split_excel(xindian_F_ok, outpath, time, second_name_list, '心电', '雄', xindian_F1, xindian_F2)

            xindian_M = pd.read_excel(i, skiprows=1, usecols='O:Y')
            xindian_M.columns = a
            xindian_M_ok, second_name_list, time,xindian_M1 = date_del(xindian_M)
            decimal(xindian_M_ok, time, '心电')
            decimal_1(xindian_M1, time, '心电')
            xindian_M2 = copy.deepcopy(xindian_M1)
            xindian_M1['检测时间'] = xindian_M1['检测时间'].apply(lambda x: x[3:])
            xindian_M2.sort_values(['动物编号', '检测时间'], inplace=True)
            xindian_M2['检测时间'] = xindian_M2['检测时间'].apply(lambda x: x[3:])
            split_excel(xindian_M_ok, outpath, time,second_name_list,  '心电', '雌', xindian_M1, xindian_M2)
            print('心电finished')


        if '呼吸' in i:
            huxi = pd.read_excel(i, skiprows=1)
            table_length = int((huxi.shape[1] - 1) / 2)
            df_left = huxi.iloc[:, 1:table_length]
            df_right = huxi.iloc[:, table_length + 2:huxi.shape[1]]
            a = list(df_left.columns)

            print(a)
            if '检测时间' in a:
                print('huxi')
                huxi_F_ok, second_name_list, time, df3 = date_del(df_left)
                print('1')
                print(huxi_F_ok)
                decimal(huxi_F_ok, time, '呼吸')
                decimal_1(df3, time, '呼吸')
                aa = list(huxi_F_ok.columns)
                df3_1 = copy.deepcopy(df3)
                df3['检测时间'] = df3['检测时间'].apply(lambda x: x[3:])
                df3_1.sort_values(['动物编号', '检测时间'], inplace=True)
                df3_1['检测时间'] = df3_1['检测时间'].apply(lambda x: x[3:])
                split_excel(huxi_F_ok,outpath,time,second_name_list,  '呼吸', '雄', df3,df3_1)
                df_right.columns = a

                huxi_M_ok, second_name_list, time, df4 = date_del(df_right)
                decimal(huxi_M_ok, time, '呼吸')
                decimal_1(df4, time, '呼吸')
                df4_1 = copy.deepcopy(df4)
                df4['检测时间'] = df4['检测时间'].apply(lambda x: x[3:])
                df4_1.sort_values(['动物编号', '检测时间'], inplace=True)
                df4_1['检测时间'] = df4_1['检测时间'].apply(lambda x: x[3:])
                split_excel(huxi_M_ok,outpath,time,second_name_list,  '呼吸', '雌', df4, df4_1)

            else:
                a = a[1:]
                new_column = ['呼吸_' + x for x in a]
                new_column.insert(0, '动物编号')
                df_left.columns = new_column
                df_right.replace('/', '', inplace=True)
                df_left.insert(0, 'Group', df_left[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
                df_left['Group'] = df_left['Group'].apply(pd.to_numeric, errors='coerce')
                df_left.to_excel(outpath + '\\呼吸雄.xlsx', index=None)
                df_right.columns = new_column
                df_right.replace('/', '', inplace=True)
                df_right.insert(0, 'Group', df_right[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
                df_right['Group'] = df_right['Group'].apply(pd.to_numeric, errors = 'coerce')

                df_right.to_excel(outpath + '\\呼吸雌.xlsx', index=None)

            print('呼吸finished')

        if '体温' in i:

            tiwen = pd.read_excel(i, skiprows=1)
            print(tiwen)
            tiwen['panduan'] = tiwen['动物编号'].apply(lambda x:x[1])
            panduan = list(tiwen['panduan'])
            print(panduan)
            if 'F' in panduan and 'M' in panduan:
                if panduan[0] == panduan[1]:
                    tiwen = tiwen.drop('panduan',axis=1)
                    table_length = int((tiwen.shape[0]) / 2)
                    df_left = tiwen.iloc[:table_length, 1:]
                    df_right = tiwen.iloc[table_length:, 1:]
                else:
                    tiwen.sort_values(['panduan','动物编号'], inplace=True)
                    print(tiwen)
                    tiwen = tiwen.drop('panduan',axis=1)
                    table_length = int((tiwen.shape[0]) / 2)
                    df_left = tiwen.iloc[:table_length, 1:]
                    df_right = tiwen.iloc[table_length:, 1:]
            else:
                tiwen = tiwen.drop('panduan',axis=1)
                table_length = 2
                df_left = tiwen.iloc[:, 0:table_length]
                df_right = tiwen.iloc[:, table_length :tiwen.shape[1]]

            print(df_left)
            print(df_right)
            a = list(df_left.columns[1:])
            new_column = ['体温℃_' + x for x in a]
            new_column.insert(0, '动物编号')
            print(new_column)
            df_left.columns = new_column
            df_left.replace('/', '', inplace=True)
            df_left.insert(0, 'Group', df_left[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
            df_left['Group'] = df_left['Group'].apply(pd.to_numeric, errors='coerce')
            df_left.to_excel(outpath + '\\体温1.xlsx', index=None)
            print(df_right)
            df_right.columns = new_column
            df_right.replace('/', '', inplace=True)
            df_right.insert(0, 'Group', df_right[:]['动物编号'].map(lambda x: str(x).strip()[:1]))
            df_right['Group'] = df_right['Group'].apply(pd.to_numeric, errors='coerce')
            df_right.to_excel(outpath + '\\体温2.xlsx', index=None)
            print('体温finished')

        if '血压' in i:
            xueya_F = pd.read_excel(i, skiprows=1, usecols='B:F')
            a = list(xueya_F.columns)
            xueya_F = date_del_1(xueya_F)

            xueya_F1 = copy.deepcopy(xueya_F)
            xueya_F2 = copy.deepcopy(xueya_F)
            xueya_F1.sort_values('动物编号',inplace=True)

            xueya_F_ok, second_name_list, time,dasd = date_del_one(xueya_F)
            decimal(xueya_F_ok, time, '血压')
            decimal_1(xueya_F1, time, '血压')
            decimal_1(xueya_F2, time, '血压')
            split_excel(xueya_F_ok, outpath, time,second_name_list,  '血压', '雄', xueya_F1,xueya_F2)

            xueya_M = pd.read_excel(i, skiprows=1, usecols='I:M')
            xueya_M.columns = a
            xueya_M = date_del_1(xueya_M)

            xueya_M1 = copy.deepcopy(xueya_M)
            xueya_M2 = copy.deepcopy(xueya_M)
            xueya_M1.sort_values('动物编号',inplace=True)

            xueya_M_ok, second_name_list, time,dasd = date_del_one(xueya_M)
            decimal(xueya_M_ok, time, '血压')
            decimal_1(xueya_M1, time, '血压')
            decimal_1(xueya_M2, time, '血压')
            split_excel(xueya_M_ok,outpath,time,second_name_list,  '血压', '雌', xueya_M1, xueya_M2)
            print('血压finished')


    convert_zip(outpath, outpath)


if __name__ == "__main__":
    in_path = r'K:\mashuaifei\死亡进行的显著性\2022-02-22'
    read_excel(in_path, in_path)
    print('all finished')
