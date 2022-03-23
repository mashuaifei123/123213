import urllib.request, http.cookiejar
import urllib.parse
import json
import pandas as pd
import sqlite3


# from .views import save_data_to_model

def pachong():
    cookie = http.cookiejar.CookieJar()
    handler = urllib.request.HTTPCookieProcessor(cookie)
    opener = urllib.request.build_opener(handler)
    logurl = "https://eam.cti-cert.com/common/login.do"
    logdata = {'account': '48502', 'password': '123', 'isFirst': 'true'}
    postdata = urllib.parse.urlencode(logdata).encode("utf-8")
    request = opener.open(logurl, data=postdata)
    header = {
        'user-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36'}
    reslll = opener.open('https://eam.cti-cert.com/storeAccount/search.do?txnCode=STORE_ACCOUNT_SEARCH')
    ccc = reslll.read().decode()
    data_list = json.loads(ccc)
    datas = data_list["results"]
    df = pd.DataFrame(datas)
    return df


df = pachong()
# save_data_to_model(df)
print(df)
# new_item = sqlite3.connect("CKinformation.db")
# print('Opened database successfully')
# df = pachong()
# items = df.to_dict('records')
# for item in items:
#     new_item.deptName = item['deptName']
#     new_item.orgName = item['orgName']
#     new_item.sortName = item['sortName']
#     new_item.houseId = item['houseId']
#     new_item.num = item['num']
#     new_item.taxTotal = item['taxTotal']
#     new_item.deptId = item['deptId']
#     new_item.factoryName = item['factoryName']
#     new_item.materialCode = item['materialCode']
#     new_item.materialId = item['materialId']
#     new_item.type = item['type']
#     new_item.orgId = item['orgId']
#     new_item.houseName = item['houseName']
#     new_item.unit = item['unit']
#     new_item.total = item['total']
#     new_item.sortId = item['sortId']
#     new_item.name = item['name']
#     new_item.partNo = item['partNo']
#     new_item.save()
# print('成功写入')
# new_item.close()
