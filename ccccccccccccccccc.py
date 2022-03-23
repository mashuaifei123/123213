# coding:utf-8
import requests


class InsertOrders:
    # 初始化设置session和headers
    def __init__(self):
        self.session = requests.session()
        # headers 信息用fiddler抓包获取
        self.headers = {
            "Referer": "http://192.168.x.xxx/xxx/xxx/login.html",
            "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
        }

    # 登录电商平台，参数从表格取出，表格数据为数据库查询出符合购买要求的客户登录id和密码
    def login(self, username, password):
        url = "http://192.168.x.xxx/xxx/xxx/login"  # 登录url
        data = {
            "USERNAME": username,
            "PASSWORD": password
        }
        print(username, password)
        login_result = self.session.post(url, data=data, headers=self.headers)
        return login_result.json()

    # 查找价格
    def findProjectPrice(self):
        find_price_url = "http://192.168.x.xxx/xxx/xxx/findCommodityDetail"  # 查找价格url
        data = {
            "modelId": "pigiron_Z14",
            "syFacilityTypeId": "CK001",
            "priceListItemId": "10000",
            "partyId": "12674",
            "_init_modelId": "pigiron_Z14",
            "_init_partyId": "12674",
            "_init_syFacilityTypeId": "CK001"
        }
        find_price_result = self.session.post(find_price_url, data=data, headers=self.headers).json()
        cashPrice = find_price_result["baseInfo"]["cashPrice"]
        return cashPrice

    # 加入购物车
    def addShopCar(self):
        shop_car_url = "http://192.168.x.xxx/xxx/xxx/insertShopCar"  # 加入购物车url
        price = self.findProjectPrice()
        data = {
            "status": "LAST",
            "priceListItemId": "10000",
            "price": price,
            "payMethod": "PAYTYPE1",
            "otherMethod": "",
            "sendMethod": "CK001",
            "getMethod": "SETTLEMENT1",
            "buyMethod": "CARTYPE001/1",
            "totalCarNum": "1",
            "totalNum": "32"
        }
        addshopcar_result = self.session.post(shop_car_url, data=data, headers=self.headers).json()
        shop_car_id = addshopcar_result["shopCarId"]
        return shop_car_id

    # 提交订单
    def insertOrders(self):
        insert_order_url = "http://192.168.x.xxx/xxx/xxx/insertOrder"  # 提交订单url
        json_arr_str = self.addShopCar()
        data = {
            "useDisStyle": "NO",
            "jsonArrStr": "[" + "'%s'" % json_arr_str + "]",
            "futureDate": "2018-11-12"
        }
        insertorder_result = self.session.post(insert_order_url, data=data, headers=self.headers).json()
        return insertorder_result


if __name__ == '__main__':

    need_list = []
    username = ["username"]
    password = ["password"]
    order = InsertOrders()
    order.login(username, password)
    # 循环取出列表中的每一个值（字典）
    for user_info in need_list:
        result = order.insertOrders()
        result_code = result['orderCode']
        print("下单成功,订单编号为：{}".format(result_code))
