#  coding=utf-8
import os # 设置虚拟路径,以供将py文件所在路径的类型文件全部导入
from os import listdir
import numpy as np
import pandas as pd
import win32com.client
from docx import Document  # docx用来操作word.docx,document用来新建空白文档
from docx.enum.text import WD_LINE_SPACING  # 设置段落的行间距
from docx.oxml.ns import qn  # 设置中文字体类型
from docx.shared import Cm, Pt  # 设置字体大小Pt磅值,设置行高Cm厘米
import time
import re
import pythoncom
from wordfindr import rest

class RemoteWord:
    def __init__(self, filename=None):
        self.xlApp = win32com.client.DispatchEx('Word.Application')
        self.xlApp.Visible = 0
        self.xlApp.DisplayAlerts = 0  # 后台运行，不显示，不警告
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                file_pass = True
                jishu = 0
                while file_pass:
                    try:
                        self.doc = self.xlApp.Documents.Open(filename, False, True)  # 第2个True只读打开文件.
                        file_pass = False
                    except:
                        jishu += 1
                        time.sleep(10)
                        if jishu == 5:
                            print('程序出现错误.请定位到文件打开部分.35行')
                            file_pass = False
            else:
                self.doc = self.xlApp.Documents.Add()  # 创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc = self.xlApp.Documents.Add()
            self.filename = ''

    def add_doc_end(self, string):
        '''在文档末尾添加内容'''
        rangee = self.doc.Range()
        rangee.InsertAfter('\n'+string)

    def add_doc_start(self, string):
        '''在文档开头添加内容'''
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string+'\n')

    def insert_doc(self, insertPos, string):
        '''在文档insertPos位置添加内容'''
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n'+string)

    def replace_doc(self, string, new_string):
        '''替换文字'''
        self.xlApp.Selection.Find.ClearFormatting()
        self.xlApp.Selection.Find.Replacement.ClearFormatting()
        self.xlApp.Selection.Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)

    def save(self):
        '''保存文档'''
        self.doc.Save()

    def save_as(self, filename):
        '''文档另存为'''
        self.doc.SaveAs(filename)

    def w_to_pdf(self, out_path:str):
        '''另存为pdf'''
        self.doc.SaveAs2(out_path, 17)
        self.xlApp.Documents.Close()

    def close(self):
        '''保存文件、关闭文件'''
        # self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()
    
    def Quit(self):
        '''退出出word'''
        self.xlApp.Quit()

    def PageSetup_Orientation(self, Find_string) -> int:
        '''查找文本内容所在章节或者页面并返回章节'''
        self.xlApp.Selection.Find.ClearFormatting()
        #  查找字符串查找完代表选中 注意匹配模式全文搜内容第78个布尔值
        self.xlApp.Selection.Find.Execute(Find_string, False, False, False, False, False, True, 1, True)
        # 可以返回查找文件所在章节
        page_number = self.xlApp.Selection.Information(2)
        return page_number
        # 把选中的内容所在章节设定为纵向
        # doc_s5 = self.xlApp.Selection.Sections[page_number]
        # PageSetup.PaperSize 
        # doc_s5.PageSetup.Orientation = 0
        # self.xlApp.Selection.Sections[4].PageSetup.Orientation = 0


def word_to_p(word_path, out_path):
    '''把一个word设定页面为A4 并把7000)所在章节设定为竖向'''
    # try:
    #     os.system('taskkill.exe /IM  WINWORD.EXE /F')
    # except:
    #     pass
    # finally:
    #     pass
    print(out_path)
    pythoncom.CoInitialize()
    docxlist = [fn for fn in listdir(word_path) if fn.endswith('.docx') if fn[0] != '~']
    bb = list(map(lambda y: word_path + '\\' + y, docxlist))
    for j in bb:
        print(bb)
        rest(j)
        pdfname = out_path + '\\'+(os.path.split(j)[1]).split('.')[0] + '.pdf'
        print(pdfname)
        # document.close()

        dpos_n_l = set()
        listfind = [' 一般状态观察个体数据 ',
                    ' 给药后一般状态观察个体数据 ',
                    ' 眼科检查个体数据 ',
                    #'Coagulation (CA-7000) 凝血个体数据',
                    'Coagulation (CA 7000)   凝血指标检测结果个体数据',
                    # 'Individual Animal Generalized Measurement by Parameter Report',
                    ' 糖化血红蛋白指标个体数据 ',
                    #' 尿沉渣个体数据 ',
                    #'尿沉渣指标检查结果个体数据',
                    'Urine Sediment   尿沉渣指标检查结果个体数据',
                    ' 生殖相关激素指标个体数据 ',
                    ' 兽医观察个体数据 ',
                    ' Body Weight Report with Individual Values ',
                    '体重测定结果个体数据',
                    ' Body Weight Gain Report with Individual Values ',
                    ' Average Food Consumption Report with Individual Values ',
                    ' Organ Weight Report with Individual Values ',
                    ' Individual Female Reproduction with Individual Fetal Data ',
                    ' Individual Female Reproduction Data ',
                    ' Individual Fetal Observation Report ',
                    ' Individual Fetal Skeletal Observation  Mean Report ',
                    ' Organ Weight to Terminal Body Weight Report with Individual Values ',
                    ' Organ Weight to Brain Weight Report with Individual Values ',
                    ' Daily Clinical Signs Summary Report ',
                    ' Individual Female Reproduction  Data with Individual Fetal Data ',
                    ' 免疫毒性指标个体数据 ',
                    ' Clinical Signs - Post Dose Summary Report ',
                    ' 体重变化幅度（%）',
                    ' 摄食量变化幅度（%）',
                    ' 临床病理指标变化幅度（%）',
                    ' 脏器重量及脏体比/脏脑比（%）',
                    ' 流式细胞指标检测结果个体数据 ',
                    ' 血生化指标检测结果个体数据 ',
                    #' 凝血指标个体数据 ',
                    '凝血指标检测结果个体数据 ',
                    #' 尿液指标检测结果个体数据 ',
                    ' 尿液指标检测结果个体数据 ',
                    #' 血液学指标检测结果个体数据 ',
                    ' 血液学指标检测结果个体数据 ',
                    #' 尿沉渣指标个体数据 ',
                    #'尿沉渣指标检查结果个体数据',
                    ' 摄食量(g/day/animal)检测结果统计汇总 ',
                    ' 体重(g)检测结果统计汇总 ',
                    '体重(g)测定结果统计汇总',
                    '体重测定结果（kg）统计汇总',
                    '摄食量测定结果（g/day）统计汇总',
                    'Average Food Consumption Summary Report',
                    ' 体重(kg)检测结果统计汇总 ',
                    'Body Weight  Summary Report',
                    ' 流式细胞指标检测结果统计汇总 ',
                    ' 血生化指标检测结果统计汇总 ',
                    ' 凝血指标检测结果统计汇总 ',
                    ' 血液学指标检测结果统计汇总 ',
                    ' 摄食量变化幅度（%） ',
                    ' 体重变化幅度（%） ',
                    ' 血液学指标变化幅度（%） ',
                    ' 凝血指标变化幅度（%） ',
                    ' 血生化指标变化幅度（%） ',
                    ' 流式细胞指标变化幅度（%） ',
                    #' 体重检测结果个体数据 ',
                    ' 体重测定结果个体数据 ',
                    #' 摄食量检测结果个体数据 ',
                    ' 摄食量测定结果个体数据 ',
                    ' 尿生化指标检测结果统计汇总 ',
                    #' 尿生化指标检测结果个体数据 ',
                    ' 尿生化指标测定结果个体数据 ',
                    ]

        '''wordapi操作查找章节'''

        doc = RemoteWord(j)  # 初始化一个doc对象 win32aip
        for o in listfind:
            Selec_info = doc.PageSetup_Orientation(o)
            dpos_n_l.add(Selec_info)
        doc.close()
        print(dpos_n_l)

        '''python-docx设定word'''
        document = Document(j)  # python-docx
        dsec_len = len(document.sections)  # word共有多少章节

        '''设定纸张依据节返回值'''
        # 设定没张纸的大小所有第单数节为横向
        for i in range(2, dsec_len, 2):
            section = document.sections[i]
            section.header_distance = Cm(0.85)  # 设定页眉上页边距
            section.page_height, section.page_width = Cm(21.0), Cm(29.7)

        # section.orientation = WD_ORIENT.LANDSCAPE  # 设定章节横纵向
        for i in dpos_n_l:  # 循环一个找到的节 集合设定页面方向
            try:
                section = document.sections[i-1]
                section.header_distance = Cm(0.85)
                new_width, new_height = section.page_height, section.page_width
                section.page_width, section.page_height  = new_width, new_height
            except:
                print('转换方向错误')
        # 设定没张纸的大小所有第双数数节为纵向
        for i in range(1, dsec_len, 2):
            section = document.sections[i]
            section.header_distance = Cm(0.85)
            section.page_height, section.page_width = Cm(29.7), Cm(21.0)

        # 第1也就是0节修改为竖向--Find不到的内容会返回1所以最后在设定第一张
        section = document.sections[0]
        section.header_distance = Cm(0.85)
        section.page_height, section.page_width = Cm(29.7), Cm(21.0)

        # 保存word到pdf
        document.save(out_path + '\\'+ 'temp.docx')
        doc2 = RemoteWord(out_path+  '\\'+ 'temp.docx')
        # out_path = os.path.join(out_path, '{}.pdf'.format(docxlist[0][:-5]))
        print(pdfname)
        doc2.w_to_pdf(pdfname)
        doc2.Quit()
        os.remove(out_path+ '\\'+ 'temp.docx')



if __name__ == "__main__":
    # time_start = time.clock()
    word_path = r'K:\mashuaifei\Fw_A2020019-T014-01个体'
    out_path = r'K:\mashuaifei\Fw_A2020019-T014-01个体'
    word_to_p(word_path, out_path)
    # print('恭喜你搞定了总耗时', time.clock() - time_start)
