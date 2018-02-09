# coding:utf-8
import xlrd
import xlwt
import requests
import urllib
import math
import re
import json
import time
import configparser
 
pattern_x = re.compile(r'"x":(".+?")')
pattern_y = re.compile(r'"y":(".+?")')
 
 
def mercator2wgs84(mercator):
    # key1=mercator.keys()[0]
    # key2=mercator.keys()[1]
    point_x = mercator[0]
    point_y = mercator[1]
    x = point_x / 20037508.3427892 * 180
    y = point_y / 20037508.3427892 * 180
    y = 180 / math.pi * (2 * math.atan(math.exp(y * math.pi / 180)) - math.pi / 2)
    return (x, y)
 
def get_config(product):
    config = configparser.SafeConfigParser()
    config.read('key.conf')
    return config.get("key", product)
 
def get_mercator(addr):
    quote_addr = urllib.parse.quote(addr.encode('utf8'))
    # city = urllib.parse.quote(u'兰州市'.encode('utf8'))
    # province = urllib.parse.quote(u'甘肃省'.encode('utf8'))
    # if quote_addr.startswith(city) or quote_addr.startswith(province):
    #     pass
    # else:
    #     quote_addr = quote_addr
    # s = urllib.parse.quote(u'山东省'.encode('utf8'))
    key = get_config('baidu')
    url = "http://api.map.baidu.com/geocoder/v2/?address=%s&output=json&ak=" + key
    api_addr = url % (
        quote_addr
        )
    req = requests.get(api_addr)
    
    content = json.loads(req.text)
    
    # x = re.findall(pattern_x, content)
    # y = re.findall(pattern_y, content)
    x = content['result']['location']['lng']
    y = content['result']['location']['lat']
    
    if x:
    #     # x = x[0]
    #     # y = y[0]
    #     # x = x[1:-1]
    #     # y = y[1:-1]
    #     x = float(x)
    #     y = float(y)
        location = (x, y)
    else:
        location = ()
    return location
 
 
def run():
    data = xlrd.open_workbook('positions.xls')
    rtable = data.sheets()[0]
    # nrows = rtable.nrows
    values = rtable.col_values(0)
    
    workbook = xlwt.Workbook()
    wtable = workbook.add_sheet('data', cell_overwrite_ok=True)
    row = 0
    for value in values:
        mercator = get_mercator(value)
        # print(mercator)
        # break
        if mercator:
            # wgs = mercator2wgs84(mercator)
            wgs = mercator
        else:
            wgs = ('NotFound', 'NotFound')
        print("%s,%s,%s" % (value, wgs[0], wgs[1]))
        wtable.write(row, 0, value)
        wtable.write(row, 1, wgs[0])
        wtable.write(row, 2, wgs[1])
        row = row + 1
        time.sleep(1)
    
    workbook.save('data.xls')
    input("positions.xls文件中的地址经纬度已获取完毕，并写入同目录下的data.xls文件中，注意此坐标系为百度bd09ll（百度经纬度坐标），点击任意键退出")
 
 
if __name__ == '__main__':
    run()