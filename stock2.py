import pandas as pd
import jieba
import jieba.analyse
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator

# 正常显示画图时出现的中文和负号
from pylab import mpl
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置字体为黑体，解决Matplotlib中文乱码问题
plt.rcParams['axes.unicode_minus'] = False  # 解决Matplotlib坐标轴负号'-'显示为方块的问题

import akshare as ak

# stock_df = ak.stock_zh_index_spot()    # 新浪财经
# print(stock_df)

# import akshare as ak
# # stock_a_pe_df = ak.stock_a_pe(market="sh")
# # print(stock_a_pe_df)


# indicator_name_list = ak.stock_sina_lhb_detail_daily(trade_date="20200730", symbol="返回当前交易日所有可查询的指标")
# print(indicator_name_list)  # 输出当前交易日可以查询的指标
# stock_sina_lhb_detail_daily_df = ak.stock_sina_lhb_detail_daily(trade_date="202101-21", symbol="涨幅偏离值达7%的证券")
# print(stock_sina_lhb_detail_daily_df)

# print(ak.stock_a_lg_indicator(stock="all") )
#
# stock_a_indicator_df = ak.stock_a_lg_indicator(stock="600089")
# print(stock_a_indicator_df)

df = ak.js_news(indicator='最新资讯')
mylist = list(df.content.values)
# 对标题内容进行分词（即切割为一个个关键词）

text = ''.join(mylist)

blacklist = ['责任编辑', '\n', '\t', '也', '上', '后', '前',
             '为什么', '再', ',', '认为', '12', '美元',
             '以及', '因为', '从而', '但', '像', '更', '用',
             '“', '这', '有', '在', '什么', '都', '是否', '一个'
    , '是不是', '”', '还', '使', '，', '把', '向', '中',
             '新', '对', ' ', ' ', u')', '、', '。', ';',
             '之后', '表示', '%', '：', '?', '...', '的', '和',
             '了', '将', '到', ' ', u'可能', '2021', '怎么',
             '从', '年', '今天', '要', '并', 'n', '《', '为',
             '月', '号', '日', '大', '如果', '哪些',
             '北京时间', '怎样', '还是', '应该', '这个',
             '这么', '没有', '本周', '哪个', '可以', '有没有']

# 设置blacklist黑名单过滤无关词语
for word in jieba.cut(text):
    if word in blacklist:
        continue
    if len(word) < 2:  # 去除单个字的词语
        continue
d = ''.join(text)
backgroud_Image = plt.imread('K:\mashuaifei\image幻想乡\B.jpg')
wc = WordCloud(
    background_color='white',
    # 设置背景颜色
    mask=backgroud_Image,
    # 设置背景图片
    font_path=r"c:\windows\fonts\simsun.ttc",
    # 若是有中文的话，这句代码必须添加
    max_words=2000,  # 设置最大现实的字数
    stopwords=STOPWORDS,  # 设置停用词
    max_font_size=150,  # 设置字体最大值
    random_state=30)
# 生成词云
wc.generate(d)
plt.figure(figsize=(12, 12), facecolor='w', edgecolor='k')
plt.imshow(wc)
# 是否显示x轴、y轴下标
plt.title('新闻标题词云\n(2021年1月26日)', fontsize=18)
plt.axis('off')
plt.show()
