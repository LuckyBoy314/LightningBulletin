# -*- coding: utf-8 -*-
import re
import urllib2
import urllib
import urlparse
from bs4 import BeautifulSoup
import cookielib

loginurl = 'http://122.224.174.179:8088/login.jsp'

class Login(object):

    def __init__(self):
        self.name = ''
        self.passwprd = ''


        self.cj = cookielib.LWPCookieJar()
        self.opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(self.cj))
        urllib2.install_opener(self.opener)

    def setLoginInfo(self,username,password):
        '''设置用户登录信息'''
        self.name = username
        self.pwd = password

    def login(self):
        '''登录网站'''
        loginparams = {'uname':self.name, 'password':self.pwd}
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36'}
        req = urllib2.Request(loginurl, urllib.urlencode(loginparams),headers=headers)
        response = urllib2.urlopen(req)
        print response.getcode()
        self.operate = self.opener.open(req)


def get_target_urls(html_cont):
    target_urls = set()

    soup = BeautifulSoup(html_cont, "lxml", from_encoding="gb2312")
    # 前缀 http://122.224.174.179:8088/
    # <a href="details.jsp?cid=20160323133024179293859zj" target="_blank">查看</a>

    links = soup.find_all('a', string="查看")
    for link in links:
        tar_url = link["href"]
        full_url = urlparse.urljoin("http://122.224.174.179:8088/", tar_url)
        target_urls.add(full_url)
    return target_urls


def download(url, headers):
    if url is None:
        return

    req = urllib2.Request(url, headers=headers)
    response = urllib2.urlopen(req)
    if response.getcode() != 200:
        return None
    return response.read()


def collect_data(url, headers):
    data = {}
    html_cont = download(url, headers)
    soup = BeautifulSoup(html_cont, "lxml", from_encoding="utf-8")

    # 时间和申请内容
    node = soup.select("#grid > tbody:nth-of-type(1) > tr:nth-of-type(9) > td:nth-of-type(2)")

    text = node[0].get_text()
    if text.find(u"雷") == -1:
        return None

    # ***去除转行符号
    text_strip = re.sub(re.compile("\s+"), " ", text)
    data["date_and_accident"] = text_strip

    # 地点
    node = soup.select("#grid > tbody:nth-of-type(1) > tr:nth-of-type(5) > td:nth-of-type(2)")
    data["place"] = node[0].get_text()

    # 单位（个人）
    node = soup.select("#grid > tbody:nth-of-type(1) > tr:nth-of-type(2) > td:nth-of-type(2)")
    data["entity"] = node[0].get_text()

    # 受损明细
    tags = soup.find_all(id="grid")
    tds = tags[2].find_all("td")

    if len(tds) > 4:
        data["loss"] = tds[4].string  # "受损物:",
        data["loss_amount"] = tds[5].string  # 受损数量
        data["direct_economic_loss"] = tds[6].string  # 直接经济损失
        data["redirect_economic_loss"] = tds[7].string  # 间接经济损失
    else:
        data["loss"] = data["loss_amount"] = data["direct_economic_loss"] = data["redirect_economic_loss"] = url
    # 人员伤亡
    tds = tags[1].find_all("td")
    data["casualty"] = tds[3].string  # 人员伤亡情况

    return data


def crawl(headers,n,outfile):
    page_num = 1

    with open(outfile, 'w') as fout:
        fout.write("时间和内容@地点@受损单位@受损物@受损数量@直接经济损失@间接经济损失@人员伤亡情况")
        fout.write("\n")
        while page_num < n:

            new_url = "http://122.224.174.179:8088/list.jsp?currentpage=%d" % page_num + "&city=%E7%BB%8D%E5%85%B4%E5%B8%82"

            print "crawling 第%d个网页" % page_num

            html_cont = download(new_url, headers)

            target_urls = get_target_urls(html_cont)

            for url in target_urls:
                try:
                    tmp_data = collect_data(url, headers)
                    if tmp_data is None:
                        continue
                    stringline = tmp_data["date_and_accident"] + "@" + tmp_data["place"] + "@" + tmp_data["entity"] + "@" + \
                                 tmp_data["loss"] + "@" + tmp_data["loss_amount"] + "@" + tmp_data[
                                     "direct_economic_loss"] + "@" + \
                                 tmp_data["redirect_economic_loss"] + "@" + tmp_data["casualty"]

                    fout.write(stringline.encode('utf-8'))
                    fout.write("\n")
                except:
                    print "Error:", url
            page_num += 1

#todo 如何自动获取coookies

if __name__ == "__main__":
    userlogin = Login()
    username = 'shaoxing_city'
    password = '123456'
    userlogin.setLoginInfo(username,password)
    userlogin.login()

    cookie = userlogin.cj
    for item in cookie:
        cookie_item = item.name + "=" + item.value

    print cookie_item
    headers = {"cookie":  cookie_item}
    n = 10#爬取的网页数目
    outfile = "out.txt"#输入文件
    crawl(headers,n,outfile)
