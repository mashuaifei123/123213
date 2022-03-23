import os
import win32com.client
 
 
# 处理Word文档的类
 
class RemoteWord:
  def __init__(self, filename=None):
      self.xlApp=win32com.client.DispatchEx('Word.Application')
      self.xlApp.Visible=0
      self.xlApp.DisplayAlerts=0    #后台运行，不显示，不警告
      if filename:
          self.filename=filename
          if os.path.exists(self.filename):
              self.doc=self.xlApp.Documents.Open(filename)
          else:
              self.doc = self.xlApp.Documents.Add()    #创建新的文档
              self.doc.SaveAs(filename)
      else:
          self.doc=self.xlApp.Documents.Add()
          self.filename=''
 
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
 
  def replace_doc(self,string,new_string):
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
 
  def close(self):
      '''保存文件、关闭文件'''
      self.save()
      self.xlApp.Documents.Close()
      self.xlApp.Quit()
 
def rest(path_word):
    # taskill_allword = input("请确认已关闭word确认请回车:")
    # os.popen('taskkill.exe /IM  WINWORD.EXE /F')
    # path_word = input('请输入word文件路径:')
    # path_cfg = input('请输入cfg文件路径:')
    #path_word = 'K:\mashuaifei\pristima_wp\Re_Re_A2020002-T005-01-20200903-094951\ViewFile个体.docx'
    # path_cfg = ''
    print(path_word)
    doc = RemoteWord(path_word)  # 初始化一个doc对象
    # list_Find_R = [line for line in open(path_cfg)]
    # list_Find_R_NEW = []
    # for p in range(len(list_Find_R)):
    #     list_Find_R_NEW.append(list_Find_R[p].strip("\n"))   # 删除/n回车符号添加到新列表
    # list1 = list_Find_R_NEW[::2]
    # list2 = list_Find_R_NEW[1::2]
    # list1 = [('动物编号', 'Animal#'),
    #                 ('症状', 'Sign'),
    #                 ('出现天数', 'Days Seen'),
    #                 ('总天数', 'Total Days Seen'),
    #                 ]
    list1 = [('体重检测结果个体数据', '体重测定结果个体数据'),
             ('体重(g)检测结果统计汇总', '体重(g)测定结果统计汇总'),
             ('摄食量检测结果个体数据','摄食量测定结果个体数据'),
             ('体温检测结果个体数据','体温测定结果个体数据'),
             ('呼吸指标测定结果个体数据','呼吸指标测定结果个体数据'),
             ('心电图测定结果个体数据','心电图检查结果个体数据'),
             ('血压指标测定结果个体数据','血压指标测定结果个体数据'),
              # ('血液学指标测定结果个体数据','血液学指标检测结果个体数据'),
               ('凝血指标个体数据','凝血指标检测结果个体数据'),
              # ('血生化指标测定结果个体数据','血生化指标检测结果个体数据'),
                 ('尿液指标检测结果个体数据','尿液指标测定结果个体数据'),
                  ('尿沉渣指标个体数据','尿沉渣指标检查结果个体数据'),
                   ('脏器重量(g)检测结果个体数据','脏器重量（g）个体数据'),
                    ('脏体比（%）个体数据','脏体比（%）计算结果个体数据'),
                    ('脏脑比（%）个体数据','脏脑比（%）计算结果个体数据'),
             ( '体重(kg)检测结果统计汇总','体重测定结果（kg）统计汇总'),
             ('摄食量(g/day/animal)检测结果统计汇总','摄食量测定结果（g/day）统计汇总 '),
             ('脏器重量(g)检测结果','脏器重量(g)')
                ]
    for i in list1:
        doc.replace_doc(i[0],i[1])  # 替换文本内容
        # print("已经检查:",i)
        # print("剩余:",len(list1)-i)
    doc.close()
    print("转换完成")

if __name__ == "__main__":
    word_path = r'K:\mashuaifei\pristima_wp\1\1.docx'
    rest(word_path)