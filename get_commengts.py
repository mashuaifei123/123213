import requests
from bs4 import BeautifulSoup


def getHtml(url):  # 下载网页源代码
    header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; LCTE; rv:11.0) like Gecko'}
    try:
        r = requests.get(url, headers=header)
        r.encoding = 'utf-8'
        # print(r.status_code)
        r.raise_for_status()
        return r.text
    except:
        getHtml(url)


def run(data):
    html = getHtml("http://guba.eastmoney.com/list," + data['share_code'] + ",f_" + str(data['page']) + ".html")
    # getAndStoreInf(html, data['share_code'])
    print('-------------page-------------' + str(data['page']))
