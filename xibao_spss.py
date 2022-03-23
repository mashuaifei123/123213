import pandas as pd
from os import listdir
import os
import re
import zipfile



def excel_spss(inpath, outpath):
    # 读取路径中的所有excel，放入列表等待处理
    print('start')
    xlsx_list = [fn for fn in listdir(inpath) if fn.endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_list_path = list(map(lambda y: inpath + '\\'+y, xlsx_list))
    for ii in xlsx_list_path:
        df_mean= pd.read_excel(ii)
        print(ii)
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
        print(df_mean)
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
        df_spss_ok['Group'] = df_spss_ok['Group'].apply(pd.to_numeric, errors='errors')
        second_name_list = CO1_name
        time = PDR_set

        for k in range(2 + len(time), df_spss_ok.shape[1]):
            df_spss_ok.iloc[:, k] = df_spss_ok.iloc[:, k].apply(lambda x: str('%.2f' % (x)))
            df_spss_ok.iloc[:, k] = df_spss_ok.iloc[:, k].apply(pd.to_numeric, errors='errors')
        #print(df_spss_ok)
        print(ii[:-5] + 'ok.xlsx')
        df_spss_ok.to_excel(ii[:-5] + 'ok.xlsx', index=None)

if __name__ == "__main__":
    in_path = r'K:\mashuaifei\Fw_A2020050-T012-01 健康检查 统计分析'
    excel_spss(in_path, in_path)
    print('all finished')