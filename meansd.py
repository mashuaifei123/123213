import pymssql
import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from openpyxl.styles import Font
from os import listdir  # 设置虚拟路径,以供将py文件所在路径的类型文件全部导入
import win32com
from docx import Document  # docx用来操作word.docx,document用来新建空白文档
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL  # 单元格对齐方式位置模块
from docx.enum.text import WD_ALIGN_PARAGRAPH  # 段落,表格对齐方式位置模块
from docx.enum.text import WD_LINE_SPACING  # 设置段落的行间距
from docx.enum.section import WD_ORIENT  # 章节的方向 PORTRAIT 纵向　LANDSCAPE 横向
from docx.enum.section import WD_SECTION  # 增加新章节 NEW_PAGE 下一页
from docx.oxml.ns import qn  # 设置中文字体类型
from docx.shared import Cm, Pt  # 设置字体大小Pt磅值,设置行高Cm厘米
from win32com.client import Dispatch, constants
import matplotlib.pyplot as plt
import matplotlib
import time
import re
import zipfile


