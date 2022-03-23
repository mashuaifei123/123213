from selenium import webdriver
import requests


def get_cookie_from_network():
    c = requests.cookies.RequestsCookieJar()
    url_login = 'https://eam.cti-cert.com/'
    driver = webdriver.PhantomJS(
        executable_path=r'K:\mashuaifei\translation-谷歌\phantomjs-2.1.1-windows\phantomjs-2.1.1-windows\bin/phantomjs.exe')
    driver.get(url_login)
    driver.find_element_by_xpath('//input[@type="text"]').send_keys('48502')
    driver.find_element_by_xpath('//input[@type="password"]').send_keys('123')

    driver.find_element_by_xpath('//input[@type="submit"]').click()

    cookie_list = driver.get_cookies()
    print(cookie_list)


cc = get_cookie_from_network()
print(cc)
