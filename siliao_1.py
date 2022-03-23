
import sys, fitz
import os
from os import listdir
import docx
from docx.shared import Inches

def pyMuPDF_fitz(pdfPath, imagePath):
    pdflist = [fn for fn in listdir(pdfPath) if fn.endswith('.pdf') if fn[0] != '~']
    bb = list(map(lambda y: pdfPath + '\\' + y, pdflist))
    for j in bb:
        name = os.path.split(j[:-4])[1]
        print("imagePath=" + j)
        pdfDoc = fitz.open(j)
    #for pg in range(pdfDoc.pageCount):
        page = pdfDoc[0]  # 选择第一页
        rotate = int(0)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
        # 此处若是不做设置，默认图片大小为：792X612, dpi=72
        zoom_x = 3 # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 3
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)

        if not os.path.exists(imagePath):  # 判断存放图片的文件夹是否存在
            os.makedirs(imagePath)  # 若图片文件夹不存在就创建

        pix.writePNG(imagePath + '/' +name+ 'images_%s.png' % pg)  # 将图片写入指定的文件夹内


def pyMuPDF2_fitz(pdfPath, imagePath):
    image_path_list = []
    # pdflist = [fn for fn in listdir(pdfPath) if fn.endswith('.pdf') if fn[0] != '~']
    # bb = list(map(lambda y: pdfPath + '\\' + y, pdflist))
    # for j in bb:
    name = os.path.split(pdfPath[:-4])[1]
    #print("imagePath=" + pdfPath)
    pdfDoc = fitz.open(pdfPath)
    # for pg in range(pdfDoc.pageCount): # iterate through the pages
    page = pdfDoc[-1]
    rotate = int(0)
    # 每个尺寸的缩放系数为3，这将为我们生成分辨率提高9倍的图像
    zoom_x = 2
    zoom_y = 2
    mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate) # 缩放系数9在每个维度  .preRotate(rotate)是执行一个旋转
    rect = page.rect                         # 页面大小
    mp = rect.tl + (rect.br - rect.tl)  *1/5# 矩形区域
    clip = fitz.Rect(mp, rect.tr)            # 想要截取的区域
    pix = page.getPixmap(matrix=mat, alpha=False, clip=clip) # 将页面转换为图像
    if not os.path.exists(imagePath):
        os.makedirs(imagePath)
    pix.writePNG(imagePath + '/' +name+ '.png' )  # 将图片写入指定的文件夹内
    # image_path_list.append(imagePath + '/' +name+ 'images_%s.png' % 0)
    # doc2 = docx.Document()
    # for i in image_path_list:
    #     doc2.add_picture(i, width=docx.shared.Inches(5))
    # doc2.save(imagePath + '/' + name + '.docx')

def pyMuPDF1_fitz(pdfPath, imagePath):
    image_path_list = []
    # pdflist = [fn for fn in listdir(pdfPath) if fn.endswith('.pdf') if fn[0] != '~']
    # bb = list(map(lambda y: pdfPath + '\\' + y, pdflist))
    # for j in bb:
    name = os.path.split(pdfPath[:-4])[1]
    #print("imagePath=" + pdfPath)
    pdfDoc = fitz.open(pdfPath)
    # for pg in range(pdfDoc.pageCount): # iterate through the pages
    page = pdfDoc[-2]
    rotate = int(0)
    # 每个尺寸的缩放系数为3，这将为我们生成分辨率提高9倍的图像
    zoom_x = 2
    zoom_y = 2
    mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate) # 缩放系数9在每个维度  .preRotate(rotate)是执行一个旋转
    rect = page.rect                         # 页面大小
    mp = rect.tl + (rect.br - rect.tl)  *1/5# 矩形区域
    clip = fitz.Rect(mp, rect.tr)            # 想要截取的区域
    pix = page.getPixmap(matrix=mat, alpha=False, clip=clip) # 将页面转换为图像
    if not os.path.exists(imagePath):
        os.makedirs(imagePath)
    pix.writePNG(imagePath + '/' +name+ '.png')  # 将图片写入指定的文件夹内



if __name__ == "__main__":
    pdfPath = 'K:\mashuaifei\图片识别\(无主题)\新建文件夹\\'
    imagePath = 'K:\mashuaifei\图片识别\(无主题)\新建文件夹\\'
    #pyMuPDF_fitz(pdfPath, imagePath)
    pyMuPDF2_fitz(pdfPath, imagePath)