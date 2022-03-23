#  coding=utf-8
# 导入相关的库     
# 使用pytesseract时注意识别语言的两个文件以及别忘记训练.参考连接https://blog.csdn.net/a745233700/article/details/80175883
# 使用注意是阶段1-px的word文件名称.输入路径请带\
from PIL import Image
from os import listdir, path
from os import remove as osremove
import pytesseract
import win32com
import numpy as np
import pandas as pd
from win32com.client import Dispatch, constants, DispatchEx
from .huxi_to_table import df_to_xls
import pythoncom


def id_to_no(id, animal_noid):  # 根据ID返回NO
    for i in animal_noid:
        if id.strip()[1] != 'F' and id.strip()[1] != 'M':
            if id.strip() == str(i[1]):
                return str(i[0])
        if id.strip()[1] == 'F' or id.strip()[1] == 'M':
            return id.strip()
    return '异常ID-{}'.format(id.strip())
    


def no_to_id(id, animal_noid):  # 根据NO返回ID
    for i in animal_noid:
        if id.strip()[1] == 'F' or id.strip()[1] == 'M':
            if id.strip() == str(i[0]):
                return str(i[1])
        if id.strip()[1] != 'F' and id.strip()[1] != 'M':
            return id.strip()
    return '异常ID-{}'.format(id.strip())


def fenzu_read(path_xlsx_list):
    fenzu_fname = path_xlsx_list + '\\' + '分组表.xls'
    df_study = pd.read_csv(open(fenzu_fname), nrows=1)
    study_name = str(df_study.columns[0]).split('for ')[1]
    df_fenzu = pd.read_csv(open(fenzu_fname), skiprows=2)
    df_fenzu.reset_index()
    animal_no = df_fenzu['Study animal number'].tolist()
    animal_id = df_fenzu['Pretest number'].tolist()
    animal_noid = list(zip(animal_no, animal_id))
    return animal_noid, study_name


def png_to_csv1(Word1_path, Out_put_path):
    animal_noid, study_name = fenzu_read(Word1_path)
    png_to_csv_as(Word1_path, Out_put_path, animal_noid, study_name)


def Reset_index(df):
    col_li = list(df.columns.level[0])
    col_li = col_li[1:]
    col_li.reverse()
    col_li0 = list(df.columns.level[1])
    col_li.extend(col_li0)
    df.columns = col_li


def sex_group(df,index):
    df = df[~df['动物编号'].isin(['-'])]
    df['sex'] = df['动物编号'].apply(lambda x: x[1])
    df7_1 = df[df['sex'].isin(['M'])]
    df7_2 = df[df['sex'].isin(['F'])]
    df7_index = list(set(df["试验阶段"]))
    df7_1_index = list(set(df7_1["试验阶段"]))
    df7_2_index = list(set(df7_2["试验阶段"]))
    title_list4 = list(
        map(lambda x: '{}'.format(x, str(x)), [index + '_' + str(j)[2:] for j in df7_index]))
    title_list5 = list(
        map(lambda x: '{}'.format(x, str(x)), [index + '_' + str(j)[2:] for j in df7_1_index]))
    title_list6 = list(
        map(lambda x: '{}'.format(x, str(x)), [index + '_' + str(j)[2:] for j in df7_2_index]))
    print(df)
    df9 =   df.pivot_table(index = '动物编号',  columns = '试验阶段',values=index)
    df9_1 = df7_1.pivot_table(index='动物编号', columns='试验阶段', values=index)
    df9_2 = df7_2.pivot_table(index='动物编号', columns='试验阶段', values=index)

    for i in df7_1_index:
        df9_1[i] = df9_1[i].apply(lambda x: "%.1f" % x)
    for i in df7_2_index:
        df9_2[i] = df9_2[i].apply(lambda x: "%.1f" % x)
    df9 = df9.replace('nan','')
    df9_1 = df9_1.replace('nan', '')
    df9_2 = df9_2.replace('nan', '')
    df9.columns = title_list4
    df9_1.columns = title_list5
    df9_2.columns = title_list6
    return df9, df9_1, df9_2


def png_to_csv_as(Word1_path, Out_put_path, animal_noid, study_name):
    png_list = Word_to_html(Word1_path + '\\' )
    Word1_path = Word1_path + '\\'      
    P_or_D_list = []
    for pdk in png_list:
        P_or_D_list.append(pdk.replace(Word1_path,'').replace('.docx.files','').replace('\\',''))
    for path_png_list in range(len(png_list)):
        pngxlist = [fn for fn in listdir(png_list[path_png_list]) if fn.endswith('.png')]  # 生成一个列表根据路径添加*.png到列表
        data_list = []
        for i in range(len(pngxlist)):
            png_path_all = png_list[path_png_list] + pngxlist[i]
            data_list.append(img_crop_02(png_path_all, png_list, path_png_list).replace('MEO-44-01','').replace('1-','').replace('2-','').replace('3-',''))
            data_list.append(',')
            data_list.append(P_or_D_list[path_png_list])
            data_list.append(',')
            data_list.append(img_crop_01(png_path_all, png_list, path_png_list).replace('\n',','))
            data_list.append(',')
            # 每三行数据增加一个回车
            if  (i + 1) % 3 == 0 :
                data_list.append('\n')
            else:
                pass
            print("已完成", pngxlist[i], )
        file = open(Word1_path + P_or_D_list[path_png_list] + '.csv','w+')
        file.write('动物编号,试验阶段,最大呼气峰压,最小吸气谷压,呼吸频率(次/分),动物编号,试验阶段,最大呼气峰压,最小吸气谷压,呼吸频率(次/分),动物编号,试验阶段,最大呼气峰压,最小吸气谷压,呼吸频率(次/分),\n')
        for i in range(len(data_list)):
            s = str(data_list[i]).replace('[','').replace(']','') # 去除[],这两行按数据不同，可以选择
            s = s.replace('’','').replace(' ','')  # 去除单引号 空格
            file.write(s)
        file.close()
        print("文件已生成", Word1_path + P_or_D_list[path_png_list] + '.csv')
        #   删除使用过的文件
        osremove(png_list[path_png_list] + "temp01.png")
        osremove(png_list[path_png_list] + "temp02.png")
        
    #  导入合并
    csv_path = Word1_path
    #要拼接的文件夹及其完整路径，注意不要包含中文
    #拼接后要保存的文件路径
    SaveFile_Name = r'all.csv'      #合并后要保存的文件名
    #将该文件夹下的所有文件名存入一个列表
    csvlist = [fn for fn in listdir(csv_path) if fn.endswith('.csv')]
    csv_file_list = []
    for csv in csvlist:
        csv_file_list.append(csv_path + csv)
    #读取第一个CSV文件并包含表头
    df = pd.read_csv(open(csv_file_list[0],encoding='GBK'))   #编码默认UTF-8，若乱码自行更改
    df.drop(df.columns[-1],axis=1,inplace=True)   #  删除最后一行无意义
    #将读取的第一个CSV文件写入合并后的文件保存
    df.to_csv(csv_path + '\\' + SaveFile_Name,index=False,encoding='GBK')
    #循环遍历列表中各个CSV文件名，并追加到合并后的文件
    for i in range(1,len(csv_file_list)):
        df = pd.read_csv(open(csv_file_list[i],encoding='GBK'))
        df.drop(df.columns[-1],axis=1,inplace=True)
        df.to_csv(csv_path + '\\' + SaveFile_Name,index=False,header=False, mode='a+',encoding='GBK')

    df1 = pd.read_csv(csv_path + '\\' + SaveFile_Name,encoding='GBK')#这个会直接默认读取到这个Excel的第一个表单
    # df1.drop(df1.columns[-1],axis=1,inplace=True)   #  删除最后一行无意义
    #  以下为原始数据计算
    osremove(csv_path + '\\' + SaveFile_Name)
    Out_put_path = Out_put_path + '\\'
    df1.insert(4,'呼吸幅度(g)',df1['最大呼气峰压'] - df1['最小吸气谷压'])
    df1.insert(10,'呼吸幅度(g).1',df1['最大呼气峰压.1'] - df1['最小吸气谷压.1'])
    df1.insert(16,'呼吸幅度(g).2',df1['最大呼气峰压.2'] - df1['最小吸气谷压.2'])
    df2 = df1[['呼吸频率(次/分)','呼吸频率(次/分).1','呼吸频率(次/分).2']]
    df3 = df1[['最大呼气峰压','最大呼气峰压.1','最大呼气峰压.2']]
    df4 = df1[['最小吸气谷压','最小吸气谷压.1','最小吸气谷压.2']]
    df5 = df1[['呼吸幅度(g)','呼吸幅度(g).1','呼吸幅度(g).2']]
    df1['mean呼吸频率(次/分)'] = df2.mean(axis=1)
    df1['mean最大呼气峰压'] = df3.mean(axis=1)
    df1['mean最小吸气谷压'] = df4.mean(axis=1)
    df1['mean呼吸幅度(g)'] = df5.mean(axis=1)
    dfok = df1.drop(['动物编号.1','动物编号.2','试验阶段.1','试验阶段.2'],axis = 1)
    dfok['动物ID'] = dfok[:]['动物编号'].map(lambda x: no_to_id(str(x), animal_noid))  # 修改动物编号为 no
    dfok['动物编号'] = dfok[:]['动物编号'].map(lambda x: id_to_no(str(x), animal_noid))  # 修改动物编号列
    dfok.sort_values(by=['试验阶段', '动物编号'], axis=0, ascending=[True, True], inplace=True )  # 先依据试验阶段排序因为实验阶段有前标，在以动物编号排序axis0 纵向 ascending 升降排序，inplace=True 原文替换.
    df1 = dfok
    dfok.to_excel(Out_put_path + '\\' + '原始数据.xlsx')

    #  以下为个体原始数据
    df6 = df1[[u'动物编号',u'试验阶段',u'mean呼吸幅度(g)']]
    df7 = df1[[u'动物编号',u'试验阶段',u'mean呼吸频率(次/分)']]
    #  转置 长转宽需要替换不然顺序会不对
    df8, df8_1, df8_2 = sex_group(df6, 'mean呼吸幅度(g)')
    df8.to_excel(Out_put_path + '\\' + '个体数据统计mean呼吸幅度(g).xlsx')
    df9, df9_1, df9_2 = sex_group(df7, 'mean呼吸频率(次/分)')
    df9.to_excel(Out_put_path + '\\' + '个体数据统计mean呼吸频率.xlsx')

    df10 = df1[[u'动物编号',u'试验阶段',u'呼吸频率(次/分)',u'呼吸幅度(g)',]]
    df10.to_excel(Out_put_path + '总表mean呼吸频率呼吸幅度.xlsx')
    df_total = [dfok, df8_1, df8_2, df9_1, df9_2]
    df_to_xls(df_total, Out_put_path + '\\' + '原始数据ok.xlsx', Table_Title_A=study_name)  # 转换一个格式内容到xlsx


def img_crop_01(filepng, png_list, path_png_list):
    # 打开一张图
    img = Image.open(filepng)
    # 图片尺寸
    img_size = img.size
    h = img_size[1]  # 图片高度
    w = img_size[0]  # 图片宽度
    # 图片位置情况1
    x = 0.796 * w
    y = 0.215 * h
    w = 0.10 * w
    h = 0.04 * h
    '''
    # 参数需要的4个值
    x = 0.788 * w
    y = 0.228 * h
    w = 0.10 * w
    h = 0.04 * h
    '''
    # 剪裁图片
    region = img.crop((x, y, x + w, y + h))
    # 2值化
    L = region.convert('L') 
    # 保存图片
    L.save(png_list[path_png_list] + "temp01.png")
    # 返回pytesseract识别的数据.
    return pytesseract.image_to_string((png_list[path_png_list] + "temp01.png"), lang='zwp')    
def img_crop_02(filepng, png_list, path_png_list):
    # 打开一张图片
    img = Image.open(filepng)
    # 图片尺寸
    img_size = img.size
    h = img_size[1]  # 图片高度
    w = img_size[0]  # 图片宽度
    x = 0.002 * w
    y = 0.03 * h
    w = 0.3 * w
    h = 0.025 * h
    # 开始截取
    region = img.crop((x, y, x + w, y + h))
    # 保存图片
    L = region.convert('L')
    #L.show()
    # 保存图片
    L.save(png_list[path_png_list] + "temp02.png")
    # 返回pytesseract识别的数据.
    return pytesseract.image_to_string((png_list[path_png_list] + "temp02.png"), lang='meo')


def Word_to_html(Word_path):  #定义转换wordtohtml函数
    pythoncom.CoInitialize()
    docxlist = [fn for fn in listdir(Word_path) if fn.endswith('.docx')]  # 生成一个列表根据路径添加*.doc到列表
    wordtopng_list = []
    for i in docxlist:
        # w = win32com.client.Dispatch('Word.Application')
        #  或者使用下面的方法，使用启动独立的进程：
        w = win32com.client.DispatchEx('Word.Application')
        # 后台运行，显示程序界面，不警告
        w.Visible = 0 #这个至少在调试阶段建议打开，否则如果等待时间长的话，它至少给你耐心。。。
        w.DisplayAlerts = 0
        path = Word_path + i
        # 打开新的文件
        worddoc = w.Documents.Open(path) #这句话用来打开已有的文件，当然，在此之前你最好判断文件是否真的存在。。。
        #doc = w.Documents.Add() # 创建新的文档，我用的更多的是这个，因为我要的是创建、然后保存为。。
        htmlformat = Word_path + i + '.html'
        worddoc.SaveAs2(htmlformat,8) #  保存为html文件格式
        w.Quit()
        #  worddoc.Close()
        wordtopng_list.append(Word_path + i + '.files' + '\\')
    return wordtopng_list


if __name__ == '__main__':
    Word1_path = '文件夹'
    Out_put_path = '文件夹'
    png_to_csv1(Word1_path, Out_put_path)
