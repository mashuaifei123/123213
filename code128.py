# 引入所需要的基本包
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128
from reportlab.graphics.barcode import eanbc, qr, usps
from reportlab.graphics.shapes import Drawing 
from reportlab.lib.units import mm, inch
from reportlab.graphics import renderPDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import os
pdfmetrics.registerFont(TTFont('pingbold', 'PingBold.ttf'))
pdfmetrics.registerFont(TTFont('ping', 'ping.ttf'))
pdfmetrics.registerFont(TTFont('msyhbd', 'msyhbd.ttf'))



# 生成PDF文件
class PDFGenerator:
    def __init__(self, file_path, filename):
        self.filename = filename
        self.file_path = file_path
        self.content_style = ParagraphStyle(name="ContentStyle", fontName="ping", fontSize=7, leading=7.5, spaceAfter=0,
                                            underlineWidth=0.3,  alignment=TA_LEFT,)
        self.content_style_C = ParagraphStyle(name="ContentStyle_c", fontName="ping", fontSize=5.7, leading=6.8, spaceAfter=0,
                                            underlineWidth=0.3,  alignment=TA_CENTER,)
        self.content_style_C_6236 = ParagraphStyle(name="ContentStyle_c6236", fontName="ping", fontSize=8, leading=16, spaceAfter=0,
                                            underlineWidth=0.3,  alignment=TA_CENTER,)
        self.content_style_C_2515 = ParagraphStyle(name="ContentStyle_c2515", fontName="msyhbd", fontSize=6, leading=7, spaceAfter=4,
                                            underlineWidth=0.3,  alignment=TA_CENTER,)

    def html_u(self, s):  # 修改字符串为下划线
        return '<u>{}</u>'.format(s)

    def genTaskPDF_tsm(self, home_data):
        story = []
        story.append(Paragraph("专题编号: " + self.html_u(home_data['studyno']), self.content_style))
        story.append(Paragraph("名称: " + self.html_u(home_data['name']), self.content_style))
        story.append(Paragraph("浓度: " + self.html_u(home_data['nongdu']) + " 数量: " + self.html_u(home_data['number']), self.content_style))
        story.append(Paragraph("贮存条件: " + self.html_u(home_data['zhucuntj']) + " 有效期至: " + self.html_u(home_data['youxiaoqi']), self.content_style))
        story.append(Paragraph("配制者: " + self.html_u(home_data['peizhizhe']) + " 配制日期: " + self.html_u(home_data['peizhidate']), self.content_style))
        story.append(Paragraph("成品编号: " + self.html_u(home_data['chengbsn']), self.content_style))
        story.append(Paragraph("备注: " + self.html_u(home_data['pristimano']), self.content_style))
        story.append(Spacer(1, 0.5 * mm))
        barcode_value = home_data['pristimano']
        barcode128 = code128.Code128(barcode_value, lquiet=0.1, barWidth=0.0115 * inch) # lquiet左边浮动距离 barWidth=0.0120 * inch 条条宽度
        story.append(barcode128)
        story.append(PageBreak())  # 分页符号
        return story
    
    def genTaskPDF_2612(self, home_data):  # 分析小标签
        story = []
        story.append(Paragraph(home_data['studyno'], self.content_style_C))
        story.append(Paragraph(home_data['v_2'], self.content_style_C))
        story.append(Paragraph(home_data['v_3'], self.content_style_C))
        story.append(PageBreak())  # 分页符号
        return story

    def genTaskPDF_2515(self, home_data):  # 临检小标签
        story = []
        story.append(Paragraph(home_data['studyno'], self.content_style_C_2515))
        story.append(Paragraph(home_data['v_2'] + home_data['v_3'], self.content_style_C_2515))
        story.append(Paragraph(home_data['v_4'], self.content_style_C_2515))
        story.append(PageBreak())  # 分页符号
        return story
    
    def genTaskPDF_6236(self, home_data):  # 分析大标签
        story = []
        story.append(Paragraph(home_data['studyno'], self.content_style_C_6236))
        story.append(Paragraph(home_data['tk'] + '&nbsp;&nbsp;&nbsp;&nbsp;' + home_data['v_3'], self.content_style_C_6236))
        story.append(Paragraph(home_data['v_4'] + '&nbsp;&nbsp;&nbsp;&nbsp;' + home_data['v_5'], self.content_style_C_6236))
        story.append(Paragraph(home_data['v_6'] + '&nbsp;&nbsp;&nbsp;&nbsp;' + home_data['v_7'], self.content_style_C_6236))
        story.append(Paragraph(home_data['v_8'], self.content_style_C_6236))
        story.append(PageBreak())  # 分页符号
        return story

    def savedoc(self, story, wight, chang):
        doc = SimpleDocTemplate(os.path.join(self.file_path, '{}.pdf'.format(str(self.filename))), pagesize=[wight * mm, chang * mm],
                                leftMargin=0.1 * mm, rightMargin=0.1 * mm, topMargin=0.1 * mm, bottomMargin=0.1 * mm)  # 设定大小注意带单位
        doc.build(story)
        
    def savedoc2612(self, story, wight, chang):
        doc = SimpleDocTemplate(os.path.join(self.file_path, '{}.pdf'.format(str(self.filename))), pagesize=[wight * mm, chang * mm],
                                leftMargin=0.1 * mm, rightMargin=0.1 * mm, topMargin=0.05 * mm, bottomMargin=0.1 * mm)  # 设定大小注意带单位
        doc.build(story)
    
    def savedoc6236(self, story, wight, chang):
        doc = SimpleDocTemplate(os.path.join(self.file_path, '{}.pdf'.format(str(self.filename))), pagesize=[wight * mm, chang * mm],
                                leftMargin=0.1 * mm, rightMargin=0.1 * mm, topMargin= 3 * mm, bottomMargin=0.1 * mm)  # 设定大小注意带单位
        doc.build(story)

    def savedoc2515(self, story, wight, chang):
        doc = SimpleDocTemplate(os.path.join(self.file_path, '{}.pdf'.format(str(self.filename))), pagesize=[wight * mm, chang * mm],
                                leftMargin=0.07 * mm, rightMargin=0.07 * mm, topMargin= 0.05 * mm, bottomMargin=0.07 * mm)  # 设定大小注意带单位
        doc.build(story)
    


def read_xlsx_pdf_tsm(path_xlsx_list, filepath):  # 用于读取Excel 生成相关文件
    xlsx_list = [fn for fn in os.listdir(path_xlsx_list) if fn.lower().endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_L_name = list(map(lambda y: path_xlsx_list + '\\' + y, xlsx_list))
    for i in xlsx_L_name:
        df = pd.read_excel(i, encoding='UTF-8', skiprows=0,)
        df.columns=['studyno', 'name', 'nongdu', 'number', 'zhucuntj', 'youxiaoqi', 'peizhizhe', 'peizhidate', 'chengbsn', 'pristimano']
        dfdice = df.to_dict(orient='records')
        filename = os.path.basename(i)[0:-5]  # 获取文件名称
        # print(dfdice)
        for i in dfdice:
            i['nongdu'] = str(i['nongdu'])
            i['pristimano'] = '{}{}'.format(i['studyno'][-8:], i['pristimano'][1:])  # 组装专题编号后八位加组别
        newpdf = PDFGenerator(filepath, filename)
        home_story = [i for b in [newpdf.genTaskPDF_tsm(x) for x in dfdice] for i in b]  # 序列解包 a=[[1,2,3],[1,2,3],[1,2,3]] [x for b in a for x in b]
        newpdf.savedoc(home_story, 50, 30)

def read_xlsx_pdf_2612(path_xlsx_list, filepath):  # 分析测试2612
    xlsx_list = [fn for fn in os.listdir(path_xlsx_list) if fn.lower().endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_L_name = list(map(lambda y: path_xlsx_list + '\\' + y, xlsx_list))
    for i in xlsx_L_name:
        df = pd.read_excel(i, encoding='UTF-8', skiprows=0,)
        df.columns=['studyno', 'v_2', 'v_3']
        dfdice = df.to_dict(orient='records')
        filename = os.path.basename(i)[0:-5]  # 获取文件名称
        newpdf = PDFGenerator(filepath, filename)
        home_story = [i for b in [newpdf.genTaskPDF_2612(x) for x in dfdice] for i in b]  # 序列解包 a=[[1,2,3],[1,2,3],[1,2,3]] [x for b in a for x in b]
        newpdf.savedoc2612(home_story, 26, 12)

def read_xlsx_pdf_3512(path_xlsx_list, filepath):  # 分析测试3512
    xlsx_list = [fn for fn in os.listdir(path_xlsx_list) if fn.lower().endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_L_name = list(map(lambda y: path_xlsx_list + '\\' + y, xlsx_list))
    for i in xlsx_L_name:
        df = pd.read_excel(i, encoding='UTF-8', skiprows=0,)
        df.columns=['studyno', 'v_2', 'v_3']
        dfdice = df.to_dict(orient='records')
        filename = os.path.basename(i)[0:-5]  # 获取文件名称
        newpdf = PDFGenerator(filepath, filename)
        home_story = [i for b in [newpdf.genTaskPDF_2612(x) for x in dfdice] for i in b]  # 序列解包 a=[[1,2,3],[1,2,3],[1,2,3]] [x for b in a for x in b]
        newpdf.savedoc2612(home_story, 35, 12)


def read_xlsx_pdf_6236(path_xlsx_list, filepath):  # 用于读取Excel 生成相关文件
    xlsx_list = [fn for fn in os.listdir(path_xlsx_list) if fn.lower().endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_L_name = list(map(lambda y: path_xlsx_list + '\\' + y, xlsx_list))
    for i in xlsx_L_name:
        df = pd.read_excel(i, encoding='UTF-8', skiprows=0,)
        df.columns=['studyno', 'tk', 'v_3', 'v_4', 'v_5', 'v_6', 'v_7', 'v_8']
        dfdice = df.to_dict(orient='records')
        filename = os.path.basename(i)[0:-5]  # 获取文件名称
        newpdf = PDFGenerator(filepath, filename)
        home_story = [i for b in [newpdf.genTaskPDF_6236(x) for x in dfdice] for i in b]  # 序列解包 a=[[1,2,3],[1,2,3],[1,2,3]] [x for b in a for x in b]
        newpdf.savedoc6236(home_story, 62, 36)


def read_xlsx_pdf_2515(path_xlsx_list, filepath):  # 用于读取Excel 生成相关文件
    xlsx_list = [fn for fn in os.listdir(path_xlsx_list) if fn.lower().endswith('.xlsx')]
    xlsx_list.sort()
    xlsx_L_name = list(map(lambda y: path_xlsx_list + '\\' + y, xlsx_list))
    for i in xlsx_L_name:
        df = pd.read_excel(i, encoding='UTF-8', skiprows=0,)
        df.columns=['studyno', 'v_2', 'v_3', 'v_4']
        dfdice = df.to_dict(orient='records')
        filename = os.path.basename(i)[0:-5]  # 获取文件名称
        newpdf = PDFGenerator(filepath, filename)
        home_story = [i for b in [newpdf.genTaskPDF_2515(x) for x in dfdice] for i in b]  # 序列解包 a=[[1,2,3],[1,2,3],[1,2,3]] [x for b in a for x in b]
        newpdf.savedoc2515(home_story, 25, 15)


if __name__ == "__main__":
    path_csv_list = r'C:\Users\admin\Desktop\22'
    read_xlsx_pdf_2515(path_csv_list, path_csv_list)
    print('ok')