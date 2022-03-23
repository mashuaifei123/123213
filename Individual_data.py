'''
本程序将处理所有的个体数据问题
2020-08-04 主要是健康检查和细胞的个体数据
'''
import pandas as pd
import numpy as np
import os
import  copy
import  re
from os import listdir
from .spss_xlsx import read_excel
from .xibaotiqu import read_excel_xibao



def excel_froms(inpath,outpath):
    print('start finding!')
    xlsx_list = [i for i in listdir(inpath) if i.endswith('.xlsx')]
    xls_list = [i for i in listdir(inpath) if i.endswith('.xls')]
    xlsx_list_path = list(map(lambda x:  inpath+'\\'+ x, xlsx_list))
    print(xlsx_list)
    print(xls_list)
    if xlsx_list:
        read_excel(inpath, outpath)
    elif xls_list:
        read_excel_xibao(inpath,outpath)
    else:
        pass

if __name__ == "__main__":
    path = 'K:\mashuaifei\Fw_A2019034-T014-01'
    excel_froms(path,path)