#coding:utf-8

import urllib2
import time
import xlwt
import threading
from bs4 import BeautifulSoup

class AQI(object):
    def __init__(self):
        self.url = 'http://www.pm25x.com'
        self.headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36'}

    # 从主页获取所有城市信息
    def getCitys(self):
        try:
            response = urllib2.urlopen(self.url)
        except urllib2.HTTPError as e:
            print 'HTTPError:', e
            return None
        homePageCode = response.read().decode('utf-8', 'ignore')
        soup = BeautifulSoup(homePageCode, 'html.parser')
        citylistBody = soup.find('dl', attrs = {'class': 'citylist'})
        soup2 = BeautifulSoup(str(citylistBody), 'html.parser')
        citylist = {}
        for cityLink in soup2.find_all('a'):
            city = cityLink.get_text()
            href = cityLink.get('href')
            citylist[city] = href
        return citylist

    # 获取城市的AQI值
    def getAQIs(self, citylist):
        aqis = {}
        for city, href in citylist.items():
            try:
                # 伪装成浏览器，防止访问次数过多出现502 Bad Gateway错误
                req = urllib2.Request(url = self.url + href, headers = self.headers)
                response = urllib2.urlopen(req)
            except urllib2.HTTPError as e:
                print 'HTTPError:', e
            cityPageCode = response.read().decode('utf-8', 'ignore')
            soup = BeautifulSoup(cityPageCode, 'html.parser')
            dataBody = soup.find('div', attrs = {'id': 'rdata'})
            soup2 = BeautifulSoup(str(dataBody), 'html.parser')
            aqivalue = soup2.find('div', attrs={'class': 'aqivalue'}).text
            aqileveltext = soup2.find('div', attrs={'class': 'aqileveltext'}).text
            aqilevel = soup2.find('span').text
            if aqivalue.isdigit(): #过滤个别抓不到数据的城市， 如 泰州
                aqis[city] = (int(aqivalue), aqileveltext, aqilevel)
                print u'已获取', city, u'的AQI信息：', aqivalue, aqileveltext, aqilevel
        return sorted(aqis.items(), key=lambda t: t[1][0], reverse=True)    # 返回list，元素为tuple
        #return OrderedDict(sorted(aqis.items(), key=lambda t: t[1][0], reverse=True))  # 返回OrderedDict

    def getAQI(self, city, href, aqis):
        try:
            req = urllib2.Request(url=self.url + href, headers=self.headers)
            response = urllib2.urlopen(req)
        except urllib2.HTTPError as e:
            print 'HTTPError:', e
            return
        cityPageCode = response.read().decode('utf-8', 'ignore')
        soup = BeautifulSoup(cityPageCode, 'html.parser')
        dataBody = soup.find('div', attrs = {'id': 'rdata'})
        soup2 = BeautifulSoup(str(dataBody), 'html.parser')
        aqivalue = soup2.find('div', attrs={'class': 'aqivalue'}).text
        aqileveltext = soup2.find('div', attrs={'class': 'aqileveltext'}).text
        aqilevel = soup2.find('span').text
        if aqivalue.isdigit(): #过滤个别抓不到数据的城市， 如 泰州
            aqis[city] = (int(aqivalue), aqileveltext, aqilevel)
            print u'已获取', city, u'的AQI信息：', aqivalue, aqileveltext, aqilevel

    def getAQIsThreads(self, citylist):
        aqis = {}
        threads = []
        for city, href in citylist.items():
            t = threading.Thread(target=self.getAQI(city, href, aqis))
            threads.append(t)
        for t in threads:
            t.start()
        for t in threads:
            t.join()    # 等待线程结束或超时
        return sorted(aqis.items(), key=lambda t: t[1][0], reverse=True)

    # 保存AQI信息
    def saveAQIs(self, aqis):
        workbook = xlwt.Workbook(encoding='utf-8')
        booksheet = workbook.add_sheet('城市AQI指数', cell_overwrite_ok=True)
        styleHead = xlwt.easyxf('font: bold 1; align: horiz center')
        styleLeft = xlwt.easyxf('align: horiz left')
        for j, data in enumerate(('城市', 'AQI指数', '空气质量', '等级')):
            booksheet.write(0, j, data, styleHead)  # 标题加粗居中
        for i, cityInfo in enumerate(aqis):
            booksheet.write(i+1, 0, cityInfo[0])
            booksheet.write(i+1, 1, cityInfo[1][0], styleLeft)    # 数字默认右对齐，改为左对齐
            booksheet.write(i+1, 2, cityInfo[1][1])
            booksheet.write(i+1, 3, cityInfo[1][2])
        t = time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime())
        fileName = 'AQIs_' + t + '.xls'
        workbook.save(fileName)

    def start(self):
        citylist = self.getCitys()
        #aqis = self.getAQIs(citylist)
        aqis = self.getAQIsThreads(citylist)
        self.saveAQIs(aqis)

if __name__ == '__main__':
    AQI().start()
