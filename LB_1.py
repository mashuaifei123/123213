
import pandas as pd
import re
# 读取form表，生成字典对应。需要form表中试验的关于时间流程类参数
df_form = pd.read_excel(r'K:\mashuaifei\骨髓\骨髓\骨髓\form_A2019028-T014-01.xlsx', sheet_name='study_info')
df_form.drop('Chinese', axis=1, inplace=True)
df_form['info'] = df_form['info'].apply(lambda x: str(x).replace('.', '-') if '.' in str(x) else x)
df_form.dropna(axis=0, how='any', inplace=True)
df_dict = df_form.set_index("item").to_dict()['info']

# 读取form表2，需要内定的术语对照表，进行merge
df2 = pd.read_stata(r'K:\mashuaifei\骨髓\骨髓\骨髓\ct_lb.dta')

# 读取已经手动填完数据的表格
df1 = pd.read_excel('K:\mashuaifei\骨髓\骨髓\骨髓\LB_BoneMarrow.xlsx')
df1.dropna(axis=1, how='all', inplace=True)
list1 = list(df.columns)
list2 = ['date', 'animal', '总数', 'CheckResult', 'Morphological_description']
list3 = list1[3:-2]
df1 = df1.melt(id_vars=list2, var_name='index', value_name='LBORRES')
# 拼接 form表2，扩充数据
df = pd.merge(df1, df2, how='left', on='index')
# 重命名从CAT到STRESU的变量名
df.columns = list(df.columns[:list(df.columns).index('CAT')]) \
             + ['LB' + str(c) for c in df.columns[list(df.columns).index('CAT'):list(df.columns).index('STRESU') + 1]] \
             + list(df.columns[list(df.columns).index('STRESU') + 1:])
# 根据form表1中的 DOSING1：给药期第一天
#             DOSING2 ：给药期第一天
# 计算数据采集时间 ，分别看雌雄是否同时开始给药，如果有dosing2，说明不同时开始，计算天数需要分雌雄。如果没有dosing2，说明同时开始给药，直接计算。
df['start1'] = df_dict['DOSING1']
anmail = df['animal']
if 'DOSING2' in df_dict:
    df['start2'] = df_dict['DOSING2']
    df['LBDY'] = df.apply(lambda x: pd.to_datetime(x['date']) - pd.to_datetime(x['start1']) if x['animal'][1] == 'F' \
        else pd.to_datetime(x['date']) - pd.to_datetime(x['start2']), axis=1).map(lambda x: x.days) + 1
else:
    df['LBDY'] = df.apply(lambda x: pd.to_datetime(x['date']) - pd.to_datetime(x['start1']), axis=1).map(
        lambda x: x.days) + 1
# 根据LBDY生成的研究天数LBNOMDY，整数值型
# 同样的计划研究时间天数VISITDY，整数值型
#  LBNOMLBL：天数+day ：DAY 28
df = df.rename(columns={'date': 'LBDTC'})
df['LBNOMDY'] = df['LBDY']
df['VISITDY'] = df['LBDY']
df['LBNOMLBL'] = df['LBNOMDY'].apply(lambda x: 'DAY ' + str(x))
df.dropna(subset=['LBORRES'], inplace=True)
# 实验的完成状态：LBSTAT ,在LBORRES里出现这两个英文语句视为没有完成该动物的实验。
# 没完成的理由 ，LBreason：没完成就有理由，需要从LBORRES中找，完成为空。
# LBORRES：测量结果或发现，
df['LBSTAT'] = df['LBORRES'].apply(lambda
                                       x: 'NOT DONE' if x == 'No marrow fluid was extracted' or x == 'Material quality is not good,Smear quality is not good,Uncountable' else '')
df['LBREASND'] = df.apply(lambda x: x['LBORRES'] if x['LBSTAT'] == "NOT DONE" else '', axis=1)
df['LBORRES'] = df.apply(
    lambda x: str(int(x['LBORRES'])) + ' of 5' if x['LBTESTCD'] == 'HYPERPLA' and x['LBORRES'] != '' else x['LBORRES'],
    axis=1)

df['DOMAIN'] = 'LB'
df['STUDYID'] = df_dict['STUDYID']
df['USUBJID'] = df['STUDYID'] + '-' + df['animal']
df['LBSTRESC'] = df['LBORRES']
# 'LBSTRESC 和LBSTRESN 都是结果的标准化，基于LBORRES，只是不同格式的展现，C即character，N即numeric。
df['LBSTRESN'] = pd.to_numeric(df['LBSTRESC'], errors='coerce')

# 根据LBDY实验的进行日期，比较form表1种的时间和 数据时间是否一致，判断是否按照计划执行 ，如果是按照计划，为空，不是计划则Y
# 'DOSDUR': 'P49D', 'RECSAC': 'P56D',
df['LBUSCHFL'] = df['LBDY'].apply(lambda x: 'Y' if x != int(re.findall('\d+', df_dict['DOSDUR'])[0]) + 1
                                                   and x != int(re.findall('\d+', df_dict['DOSDUR'])[0]) + int(
    re.findall('\d+', df_dict['RECSAC'])[0]) + 1 else '')
# 对LBDY的补充说明吧，对准时执行的实验说明
df['LBTPT'] = df['LBDY'].apply(
    lambda x: 'Anatomy on the first day of recovery' if x == int(re.findall('\d+', df_dict['DOSDUR'])[0]) + 1
    else 'Day after the end of recovery period' if x == int(re.findall('\d+', df_dict['DOSDUR'])[0]) + int(
        re.findall('\d+', df_dict['RECSAC'])[0]) + 1
    else '')

# 需要将 指标中的‘CL_DESC’，'CL_MORPH'放到末尾
df['LBTPTNUM'] = df['LBTPT'].apply(lambda x: 1 if x == 'Anatomy on the first day of recovery'
else 2 if x == 'Day after the end of recovery period' else '')
df['LBTPTNUM'] = pd.to_numeric(df['LBTPTNUM'])
df['lbtest_num'] = df['LBTESTCD'].apply(lambda x: 99 if x == 'CL_DESC' else 100 if x == 'CL_MORPH' else 0)
df = df.sort_values(by=['USUBJID', 'LBDY', 'lbtest_num'], ascending=[True, True, True])
df.reset_index(drop=True, inplace=True)
# 生成一列分组数据，其实就是groupby。
AA = df['USUBJID']
d = 1
lbseq = [1]
for i in range(1, len(AA)):
    if AA[i] == AA[i - 1]:
        d += 1
        lbseq.append(d)
    else:
        d = 1
        lbseq.append(d)
df['LBSEQ'] = lbseq

## 判断需要哪些列进行导出

variable = list(df.columns)
if 'LBSTAT' in variable:
    if 'LBUSCHFL' in variable:
        df = pd.DataFrame(df[['STUDYID', 'DOMAIN', 'USUBJID', 'LBSEQ', 'LBTESTCD', 'LBTEST', 'LBCAT', 'LBORRES',
                              'LBORRESU', 'LBSTRESC', 'LBSTRESN', 'LBSTRESU', 'LBSTAT', 'LBREASND', 'LBSPEC',
                              'LBUSCHFL', 'VISITDY', 'LBDTC', 'LBDY', 'LBNOMDY', 'LBNOMLBL', 'LBTPT', 'LBTPTNUM']])
    else:
        df = pd.DataFrame(df[['STUDYID', 'DOMAIN', 'USUBJID', 'LBSEQ', 'LBTESTCD', 'LBTEST', 'LBCAT', 'LBORRES',
                              'LBORRESU', 'LBSTRESC', 'LBSTRESN', 'LBSTRESU', 'LBSTAT', 'LBREASND', 'LBSPEC',
                              'VISITDY', 'LBDTC', 'LBDY', 'LBNOMDY', 'LBNOMLBL', 'LBTPT', 'LBTPTNUM']])
else:
    if 'LBUSCHFL' in variable:

        df = pd.DataFrame(df[['STUDYID', 'DOMAIN', 'USUBJID', 'LBSEQ', 'LBTESTCD', 'LBTEST', 'LBCAT', 'LBORRES',
                              'LBORRESU', 'LBSTRESC', 'LBSTRESN', 'LBSTRESU', 'LBSPEC',
                              'LBUSCHFL', 'VISITDY', 'LBDTC', 'LBDY', 'LBNOMDY', 'LBNOMLBL', 'LBTPT', 'LBTPTNUM']])
    else:
        df = pd.DataFrame(df[['STUDYID', 'DOMAIN', 'USUBJID', 'LBSEQ', 'LBTESTCD', 'LBTEST', 'LBCAT', 'LBORRES',
                              'LBORRESU', 'LBSTRESC', 'LBSTRESN', 'LBSTRESU', 'LBSPEC',
                              'VISITDY', 'LBDTC', 'LBDY', 'LBNOMDY', 'LBNOMLBL', 'LBTPT', 'LBTPTNUM']])
# 拆分字符过长问题

lborres1 = []
for i in df['LBORRES']:
    if len(str(i)) > 200:
        list4 = [a.span()[0] for a in re.finditer(' ', i)]
        # print(list4)
        for j in range(len(list4)):
            if list4[j] > 200:
                s = list4[j - 1]
                lborres1.append(i[s:])
                break
    else:
        lborres1.append('')
if len(set(lborres1)) > 1:
    df['LBORRES'] = df['LBORRES'].apply(lambda x: str(x)[:s])
    df.insert(8, 'LBORRES1', lborres1)

lbstresc = []
for i in df['LBSTRESC']:
    if len(str(i)) > 200:
        list4 = [a.span()[0] for a in re.finditer(' ', i)]
        # print(list4)
        for j in range(len(list4)):
            if list4[j] > 200:
                s = list4[j - 1]
                lbstresc.append(i[s:])
                break
    else:
        lbstresc.append('')
if len(set(lbstresc)) > 1:
    df['LBSTRESC'] = df['LBSTRESC'].apply(lambda x: str(x)[:s])
    df.insert(11, 'LBSTRESC1', lbstresc)

df['LBORRES'] = df['LBORRES'].astype(str)
df['LBSTRESC'] = df['LBSTRESC'].astype(str)
df.to_stata('K:\mashuaifei\骨髓\骨髓\骨髓\lb_bw1.dta')

