
import pandas as pd
df1= pd.read_excel('K:\mashuaifei\骨髓\骨髓\骨髓\LB_BW.xlsx')
df1= df1.reset_index(drop=True)

list1=['SQYuanShiLinBa', 'SQYouZhiLinBa', 'SQChengShuLinBa', 'SQYiXingLinBa', 'SQSiJianLinBa',
      'SQYuanShiDanHe', 'SQYouZhiDanHe', 'SQChengShuDanHe','SQYuanShiJiang','SQYouZhiJiang',
       'SQChengShuJiang','SLSSZhaoYou','SLSSWanYou','SLSSGanZhuangHe','SLSSFenYeHe',
      'SLSJZhaoYou','SLSJWanYou','SLSJGanZhuangHe','SLSJFenYeHe']
for i in list1:
    df1[i].fillna(0, inplace=True)

df1.eval('SQLinBaTotal = SQYuanShiLinBa + SQYouZhiLinBa + SQChengShuLinBa + SQYiXingLinBa + SQSiJianLinBa' , inplace=True)
df1.eval('SQDanHeTotal= SQYuanShiDanHe + SQYouZhiDanHe + SQChengShuDanHe' , inplace=True)
df1.eval('SQJiangTotal= SQYuanShiJiang + SQYouZhiJiang + SQChengShuJiang' , inplace=True)
df1.eval('SLShisuanxingli= SLSSZhaoYou+ SLSSWanYou+ SLSSGanZhuangHe+ SLSSFenYeHe' , inplace=True)
df1.eval('SJShijianxingli= SLSJZhaoYou+ SLSJWanYou+ SLSJGanZhuangHe+ SLSJFenYeHe' , inplace=True)
df = pd.DataFrame(df1[['PatientName','SendTime', 'SuiTotal', 'SuiLiHongBi', 'SuiQiTaTotal',
                       'SQYuanShiXue', 'SuiLiTotal', 'SLYuanShiLi', 'SLZhaoYouLi', 'SLZXZhongYou', 'SLZXWanYou',
                       'SLZXGanZhuangHe', 'SLZXFenYeHe', 'SuiHongTotal', 'SHYuanShiHong', 'SHZhaoYouHong',
                       'SHZhongYouHong', 'SHWanYouHong', 'SHYuanJuHong', 'SHZhaoJuHong',
                       'SHZhongJuHong', 'SHWanJuHong', 'SuiJHTotal', 'SQLinBaTotal', 'SQDanHeTotal', 'SQJiangTotal',
                       'SLShisuanxingli', 'SJShijianxingli', 'CheckResult','CheckSee']])
list2= ['animal','date','总数', '粒红比', '其它',
        '原始血细胞', '粒系', '原始粒', '早幼粒', '中中粒', '中晚粒',
        '中杆粒', '中分叶核',  '红系','原红', '早幼红',
        '中幼红', '晚幼红', '原巨红', '早巨红',
        '中巨红', '晚巨红', '淋巴', '单核', '巨核细胞', '浆细胞' ,
        '嗜酸性粒细胞', '嗜碱性粒细胞', 'CheckResult','Morphological_description' ]
df.columns= list2
df.dropna(axis=1, how='all', inplace=True)
for i in list(df.columns[4:-2]):
    df[i] = df.apply(lambda x:  x[i]*100 / x['总数'] ,axis= 1  )
df = df.fillna(0)
df['date'] = df['date'].apply(lambda x: str(x)[:10])
df['Description'] = None
df['Diagnosis'] = None
df['zs']=None
df.insert(list(df.columns).index('CheckResult'),'增生',df['zs'])
df.drop(['zs'],axis=1,inplace=True)
df.to_excel('K:\mashuaifei\骨髓\骨髓\骨髓\LB_BW21.xlsx')