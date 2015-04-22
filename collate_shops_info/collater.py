#coding=utf-8
import xlrd
import urllib, urllib2
try:
    import json
except ImportError:
    import simplejson as json       # pylint: disable-msg=F0401
    
GEOCODE_QUERY_GOOGLE_URL = 'http://maps.googleapis.com/maps/api/geocode/json?'

GEOCODE_QUERY_BAIDU_URL = 'http://api.map.baidu.com/geocoder?'
BAIDU_APP_KEY= 'E81fa5f3afe8a077938013718853a8dc'

FILE_PATH = '/Users/shaoli/Desktop/shop_list.txt'
WRONG_FILE_PATH = '/Users/shaoli/Desktop/wrong_shop_list.txt'

ZOOC_INFO_FILE_PATH = '/Users/shaoli/Desktop/finally_shops_info/zooc.xls'
LANCY_INFO_FILE_PATH = '/Users/shaoli/Desktop/finally_shops_info/lancy.xls'
MM_INFO_FILE_PATH = '/Users/shaoli/Desktop/finally_shops_info/mm.xls'
LIME_INFO_FILE_PATH = '/Users/shaoli/Desktop/finally_shops_info/lime.xls'

class ShopInfo:
    def __init__(self):
        self.title = ''
        self.brand = ''
        self.city = ''
        self.province = ''
        self.address = ''
        self.number = ''
        self.lat = 0
        self.lng = 0
    
    def show_shop(self):
        print 'Province: %s' % self.province
        print 'Brand: %s' % self.brand
        print 'City: %s' % self.city
        print 'Title: %s' % self.title
        print 'Address: %s' % self.address
        print 'Number: %s' % self.number
        print '%f, %f' % (self.lat, self.lng)
        print self.get_info_dic()
        print '==========================='
        
    def get_info_dic(self):
        return {'province': self.province,
                'brand': self.brand,
                'city': self.city,
                'title':  self.title,
                'address': self.address,
                'number': self.number,
                'lat': self.lat,
                'lng': self.lng,
                }

shop_list = []

def readShopList():
    file_str = FILE_PATH
    shop_list_file = open(file_str, 'r')
    shop_list_str = shop_list_file.read()
    shop_list = eval(shop_list_str)
    shop_list_file.close()
    return shop_list

def saveShoplist():
    file_str = FILE_PATH
    wrong_file_str = WRONG_FILE_PATH
    wrong_address_shops = []
    shops = []
    for shop in shop_list:
        shops.append(shop)
#        if shop['lat'] == 0:
#            wrong_address_shops.append(shop)
#        else:
#            shops.append(shop)
    shop_list_file = open(file_str, 'w')
    shop_list_str = str(shops)
    shop_list_file.write(shop_list_str)
    shop_list_file.close()
    
    wrong_shop_list_file = open(wrong_file_str, 'w')
    wrong_shop_list_str = str(wrong_address_shops)
    wrong_shop_list_file.write(wrong_shop_list_str)
    wrong_shop_list_file.close()

class MapKit(object):
    def get_latlng_by_address(self, address):
        pass
    
    def fetch_json(self, query_url, params={}, headers={}):       # pylint: disable-msg=W0102
        """Retrieve a JSON object from a (parameterized) URL.
        
        :param query_url: The base URL to query
        :type query_url: string
        :param params: Dictionary mapping (string) query parameters to values
        :type params: dict
        :param headers: Dictionary giving (string) HTTP headers and values
        :type headers: dict 
        :return: A `(url, json_obj)` tuple, where `url` is the final,
        parameterized, encoded URL fetched, and `json_obj` is the data 
        fetched from that URL as a JSON-format object. 
        :rtype: (string, dict or array)
        """
        encoded_params = urllib.urlencode(params)    
        url = query_url + encoded_params
        request = urllib2.Request(url, headers=headers)
        response = urllib2.urlopen(request)
        return (url, json.load(response))

class BaiduMapKit(MapKit):
    def get_latlng_by_address(self, address):
        params = {
            'address':  address,
            'output':   'json',
            'key':       BAIDU_APP_KEY,
        }
        url, response = self.fetch_json(GEOCODE_QUERY_BAIDU_URL, params)
        if response['status'] == 'OK' and response['result']:
            location = response['result']['location']
            return location['lat'], location['lng']
        return 0, 0

class GoogleMapKit(MapKit):
    def get_latlng_by_address(self, address):
        params = {
                'address':  address,
                'sensor':   'false',
            }
        url, response = self.fetch_json(GEOCODE_QUERY_GOOGLE_URL, params)
        if response['status'] == 'OK':
            location = response['results'][0]['geometry']['location']
            return location['lat'], location['lng']
        return 0, 0

class MapKitFactory(object):
    def getMapKit(self, name):
        if 'baidu'  == name:
            return BaiduMapKit()
        elif 'google' == name:
            return GoogleMapKit()
        
        return MapKit()

def get_latlng_by_address(address):
    f = MapKitFactory()
    m = f.getMapKit('baidu')
    return m.get_latlng_by_address(address)

def parse_zooc():
    wb = xlrd.open_workbook(ZOOC_INFO_FILE_PATH)
    table = wb.sheets()[0]
    nrows = table.nrows
    for i in range(2, nrows-1):
        shop = ShopInfo()
        shop.province = table.row(i)[0].value.strip()
        shop.city = table.row(i)[1].value.strip()
        shop.title = table.row(i)[2].value.strip()
        address, number = table.row(i)[3].value.split(u'电话')
        shop.address = address.strip()
        shop.number = number.strip()
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'mo'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
        
    table = wb.sheets()[1]
    nrows = table.nrows
    for i in range(2, nrows-1):
        shop = ShopInfo()
        shop.province = table.row(i)[0].value.strip()
        shop.city = table.row(i)[0].value.strip()
        shop.title = table.row(i)[1].value.strip()
        shop.address = table.row(i)[4].value.strip()
        shop.number = table.row(i)[3].value.strip()
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'mo'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
    
    table = wb.sheets()[2]
    nrows = table.nrows
    for i in range(5, nrows-1):
        shop = ShopInfo()
        shop.province = table.row(i)[0].value.strip()
        shop.city = table.row(i)[1].value.strip()
        shop.title = table.row(i)[2].value.strip()
        address, number = table.row(i)[3].value.split(u'电话')
        shop.address = address.strip()
        shop.number = number.strip()
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'mo'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())

def parse_lancy():
    wb = xlrd.open_workbook(LANCY_INFO_FILE_PATH)
    table = wb.sheets()[0]
    nrows = table.nrows
    for i in range(2, nrows-4):
        shop = ShopInfo()
        shop.province = table.row(i)[3].value.strip()
        shop.city = table.row(i)[2].value.strip()
        shop.title = table.row(i)[1].value.strip()
        shop.address = table.row(i)[5].value.strip()
        shop.number = table.row(i)[4].value
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'lf25'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
        
def parse_lime():
    wb = xlrd.open_workbook(LIME_INFO_FILE_PATH)
    table = wb.sheets()[0]
    nrows = table.nrows
    for i in range(2, nrows):
        shop = ShopInfo()
        shop.province = table.row(i)[1].value.strip()
        shop.city = table.row(i)[2].value.strip()
        shop.title = table.row(i)[3].value.strip()
        shop.address = table.row(i)[5].value.strip()
        shop.number = table.row(i)[4].value
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'lf'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
        
def parse_mm():
    wb = xlrd.open_workbook(MM_INFO_FILE_PATH)
    table = wb.sheets()[0]
    nrows = table.nrows
    for i in range(1, nrows):
        shop = ShopInfo()
        shop.province = table.row(i)[0].value.strip()
        shop.city = table.row(i)[1].value.strip()
        shop.title = table.row(i)[2].value.strip()
        shop.address = table.row(i)[3].value.strip()
        shop.number = table.row(i)[4].value
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'mm'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
        
    table = wb.sheets()[1]
    nrows = table.nrows
    for i in range(1, nrows):
        shop = ShopInfo()
        shop.province = table.row(i)[1].value.strip()
        shop.city = table.row(i)[2].value.strip()
        shop.title = table.row(i)[3].value.strip()
        shop.address = table.row(i)[4].value.strip()
        shop.number = table.row(i)[5].value
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'mm'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
        
    table = wb.sheets()[2]
    nrows = table.nrows
    for i in range(1, nrows):
        shop = ShopInfo()
        shop.province = table.row(i)[0].value.strip()
        shop.city = table.row(i)[1].value.strip()
        shop.title = table.row(i)[2].value.strip()
        shop.address = table.row(i)[4].value.strip()
        shop.number = table.row(i)[3].value
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'mm'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
        
    table = wb.sheets()[3]
    nrows = table.nrows
    for i in range(1, nrows):
        shop = ShopInfo()
        shop.province = table.row(i)[0].value.strip()
        shop.city = table.row(i)[1].value.strip()
        shop.title = table.row(i)[2].value.strip()
        shop.address = table.row(i)[4].value.strip()
        shop.number = table.row(i)[3].value
        lat, lng = get_latlng_by_address(shop.address.encode('utf-8'))
        shop.lat = lat
        shop.lng = lng
        shop.brand = 'mm'
        shop.show_shop()
        shop_list.append(shop.get_info_dic())
        
if __name__ == "__main__":
# parse data
    parse_zooc()
    parse_lancy()
    parse_lime()
    parse_mm()
    saveShoplist()

# change data type to sql
#    shop_list = readShopList()
#    for shop in shop_list:
#        print "('%d','%s','%s','%s','%s','%s','%s','%f','%f',1)," \
#        % (shop_list.index(shop), shop['title'], shop['number'], shop['province'], shop['city'], shop['address'], shop['brand'], shop['lat'], shop['lng'])

# modify wrong data
#    wrong_data = [u'合肥市宿州路4号合肥金鹰购物中心', u'昆明市白搭路金格百货', u'温州市鹿城区荷花路银泰百货'
#                  , u'天津泰达市民文化广场友谊名都', u'天津市河西区乐园道乐天百货'
#                  , u'西安市经开区凤城五路世纪金花赛高购物中心', u'大庆市大商新玛特纬二路新玛特'
#                  , u'南京市玄武区中山路德基广场', u'南京汉中路金鹰国际商城'
#                  , u'南京市草场门大街龙江新城市广场', u'太原市小店区亲贤北街梅元百盛'
#                  , u'太原市长风街燕莎友谊商场', u''
#                  , u'西安市劳动南路大唐西市高业管理有限公司', u'南宁市清秀区民族大道中段梦之岛百货'
#                  , u'沈阳市和平区青年大街沈阳华润中心万象城', u'济南市历下区天地坛街贵和购物中心'
#                  , u'济南市市中区英雄山路八一银座', u'郑州市花园路新玛特'
#                  , u'郑州市新玛特购物广场', u'靖江市驥江西路天一百货购物中心'
#                  , u'北京朝阳区七圣中街百盛商场', u'北京市西城区长安商场'
#                  , u'北京市海淀区复兴路翠微大厦', u'北京市中关村大街当代商城'
#                  , u'北京市海淀区双安商场', u'北京市东城区王府井百货大楼'
#                  , u'北京市海淀区燕莎友谊商城', u'北京市朝阳区燕莎友谊商城'
#                  , u'北京市崇文区崇文门外大街新世界百货', u'北京朝阳区七圣中街太阳宫百盛'
#                  , u'天津第一大街友谊名都', u'济南市市中区经十路鲁商广场'
#                  , u'鞍山铁东区天兴百盛购物中心', u'大连中山区新玛特购物休闲广场'
#                  , u'武汉市汉口解放大道庄胜SOGO', u'襄阳长虹路武商购物中心'
#                  , u'宜昌市夷陵大道丹尼斯购物广场', u'连云港新浦区九龙大世界'] 
#    wrong_data = [u'沈阳市和平区青年大街沈阳华润中心万象城',u'济南市历下区天地坛街贵和购物中心',
#                  u'济南市八一银座',u'河南省郑州市花园路38号新玛特（郑州）总店 国贸店 三楼朗姿专柜',
#                  u'河南省郑州市二七路200号郑州新玛特购物广场二店（金博大）三楼朗姿专柜',u'靖江市驥江西路299号靖江市天一百货购物中心有限公司2楼朗姿专柜',
#                  u'天津经济技术开发区第一大街86号友谊名都市民广场店女服组2层莱茵专柜',u'济南市市中区经十路19288号鲁商广场内玉函银座二层2层MM专柜',
#                  u'辽宁省鞍山市铁东区二道街88号鞍山天兴百盛购物中心二层MM专柜',u'',
#                  u'辽宁省大连市中山区青三街1#新玛特购物休闲广场1号中心2FMM专柜',u'',]
#    for address in wrong_data:
#        lat, lng = get_latlng_by_address(address.encode('utf-8'))
#        print address
#        print lat, lng
        