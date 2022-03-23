import tushare as ts
import pandas as pd
import matplotlib.pyplot as plt
from pyecharts import Map
# 正常显示画图时出现的中文和负号
from pylab import mpl

mpl.rcParams['font.sans-serif'] = ['SimHei']
mpl.rcParams['axes.unicode_minus'] = False

# 设置token
token = '1290830128b2bc2a394f077db4fb96b5c7feff3b05e6a01dae5254be'
# ts.set_token(token)
pro = ts.pro_api(token)

basic = pro.stock_basic()
basic.to_csv(r"K:\mashuaifei\stock\basics_data.csv", encoding='gbk')
# print(basic.iloc[:3,:8])
area = basic.groupby('area')['name'].count()

area['广东'] = area['广东'] + area['深圳']
area.drop(['深圳'], inplace=True)
# print(area.sort_values(ascending=False)[:20])
d = dict(area)
province = list(d.keys())
value = list(d.values())

# map = Map("中国上市公司分布", title_color="#fff",
#           title_pos="center", width=1200,
#           height=600,background_color='#404a59')
# map.add("", province, value,visual_range=[min(value),max(value)],
#        is_label_show=True,maptype='china',
#         visual_text_color='#000',label_pos="center",
#         is_visualmap=True)
# map.render(path="中国上市公司分布.html")

basics_data = pro.daily_basic(ts_code='', trade_date='20180726',
                              fields='ts_code,trade_date,turnover_rate,volume_ratio,pe,pb')

d = ['name', 'industry', 'pe', 'pb', 'totalAssets',
     'esp', 'rev', 'profit', 'gpr', 'npr']
df = basics_data[d]
print(df.head(3))
