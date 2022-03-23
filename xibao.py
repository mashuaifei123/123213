import numpy as np
import pandas as pd
from os import listdir
from .excel_xianzhux import *
from .excel_xianzhuxing3 import *



def xibao_xzx(path_list, out_path):
    print('start')
    print(out_path)
    xianzhuxing3(path_list, out_path)
    xianzhuxing1(path_list, out_path)



if __name__ == "__main__":
    path_list = r'K:\mashuaifei\死亡进行的显著性\细胞因子\\'
    df = xibao_xzx(path_list, path_list)
