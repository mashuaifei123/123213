import sys, fitz
import os
from os import listdir
import pandas as pd
import docx
from docx.shared import Inches


def get_pages(path):
    excellist = [fn for fn in listdir(path) if fn.endswith('.xlsx') if fn[0] != '~']
    bb = list(map(lambda y: pdfPath + '\\' + y, excellist))
    df = pd.read_excel(bb[0], header = None)
    data_dict = df.T.to_dict('list')
    for i in range(len(data_dict)):
        if len(data_dict[i]) > 6 or len(data_dict[i]) < 1:
            print('excel有问题')
            data_dict = {}
        else:
            for j in data_dict[i]:
                if j > 6 or j < 1:
                    print('excel有问题')
                    data_dict = {}
    print(data_dict)
    return data_dict


def pyMuPDF2_fitz(pdfPath, imagePath):
    data_dict = get_pages(pdfPath)
    if data_dict:
        pdflist = [fn for fn in listdir(pdfPath) if fn.endswith('.pdf') if fn[0] != '~']
        bb = list(map(lambda y: pdfPath + '\\' + y, pdflist))
        for j in bb:
            name = os.path.split(j[:-4])[1]
            print("imagePath=" + j)
            pdfDoc = fitz.open(j)
            image_path_list =[]
            if pdfDoc.pageCount == len(data_dict):
                for pg in range(pdfDoc.pageCount): # iterate through the pages
                    page = pdfDoc[pg]
                    rotate = int(0)
                    # 每个尺寸的缩放系数为9，这将为我们生成分辨率提高 9x9 倍的图像
                    zoom_x = 9
                    zoom_y = 9
                    mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate) # 缩放系数9在每个维度  .preRotate(rotate)是执行一个旋转

                    rect = page.rect                         # 截取页面大小
                    X_li = rect.br[0]/7578                   # 尺寸像素比  尺寸/像素
                    Y_li = rect.br[1]/5355
                    A1 = [(7578-7250)/2*X_li-1, (5355-4000)/2*Y_li-1]                     # 左边四个点
                    A2 = [(7578-7250)/2*X_li-1, ((5355-4000)/2+4000*1/3)*Y_li-1]
                    A3 = [(7578-7250)/2*X_li-1, ((5355-4000)/2+4000*2/3)*Y_li-1]
                    A4 = [(7578-7250)/2*X_li-1, (5355-(5355-4000)/2)*Y_li-1]

                    B1 = [((7578-7250)/2+7250*1/2)*X_li-1, (5355-4000)/2*Y_li-1]          # 中间四个点
                    B2 = [((7578-7250)/2+7250*1/2)*X_li-1, ((5355-4000)/2+4000*1/3)*Y_li-1]
                    B3 = [((7578-7250)/2+7250*1/2)*X_li-1, ((5355-4000)/2+4000*2/3)*Y_li-1]
                    B4 = [((7578-7250)/2+7250*1/2)*X_li-1, (5355-(5355-4000)/2)*Y_li-1]

                    C1 = [(7578-(7578-7250)/2) * X_li+0.3, (5355-4000)/2*Y_li+0.3]         # 右边四个点
                    C2 = [(7578-(7578-7250)/2) * X_li+0.3, ((5355-4000)/2+4000*1/3)*Y_li+0.3]
                    C3 = [(7578-(7578-7250)/2) * X_li+0.3, ((5355-4000)/2+4000*2/3)*Y_li+0.3]
                    C4 = [(7578-(7578-7250)/2)*X_li+0.3, (5355-(5355-4000)/2)*Y_li+0.3]

                    rect.A1 = rect.tl + A1         # 列表位置 转 图像位置
                    rect.A2 = rect.tl + A2
                    rect.A3 = rect.tl + A3
                    rect.A4 = rect.tl + A4

                    rect.B1 = rect.tl + B1
                    rect.B2 = rect.tl + B2
                    rect.B3 = rect.tl + B3
                    rect.B4 = rect.tl + B4

                    rect.C1 = rect.tl + C1
                    rect.C2 = rect.tl + C2
                    rect.C3 = rect.tl + C3
                    rect.C4 = rect.tl + C4
                    # mp = rect.tl + (rect.br - rect.tl) * 0.1
                    # mp = rect.tl + (rect.br - rect.tl) *1/5 # 矩形区域
                    # mp = rect.tl + (rect.br - rect.tl) * 0.5  #中心

                    clip1 = fitz.Rect(rect.A1, rect.B2)         # 想要截取的区域1
                    clip2 = fitz.Rect(rect.A2, rect.B3)
                    clip3 = fitz.Rect(rect.A3, rect.B4)
                    clip4 = fitz.Rect(rect.B1, rect.C2)
                    clip5 = fitz.Rect(rect.B2, rect.C3)
                    clip6 = fitz.Rect(rect.B3, rect.C4)
                    clip_all = [clip1, clip2, clip3, clip4, clip5, clip6]
                    for cc in data_dict[pg]:
                        pix = page.getPixmap(matrix=mat, alpha=False, clip=clip_all[cc-1]) # 将页面转换为图像
                        if not os.path.exists(imagePath):
                            os.makedirs(imagePath)
                        pix.writePNG(imagePath + '/' + name + 'images_%s' % (pg+1) + '_%s' % cc + '.png')  # 将图片写入指定的文件夹内
                        image_path_list.append(imagePath + '/' + name + 'images_%s' % (pg+1) + '_%s' % cc + '.png')
            else:
                pass
    doc2 = docx.Document()
    for i in image_path_list:
        doc2.add_picture(i, width=docx.shared.Inches(5))
    doc2.save(imagePath + '/' + name + '.docx')
    for root, dirs, files in os.walk(pdfPath):
        for name in files:
            if name.endswith(".png"):
                os.remove(os.path.join(root, name))

if __name__ == "__main__":
    pdfPath = 'K:\mashuaifei\转发_ PDF图片截取\\'
    imagePath = 'K:\mashuaifei\转发_ PDF图片截取\\'
    pyMuPDF2_fitz(pdfPath, imagePath)
