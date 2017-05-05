#coding:utf-8

import sys
import re
import urllib2
import urllib
import requests
import cookielib


#####################################################
#登录人人
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
        self.operate = self.opener.open(req)
        thePage = response.read()

if __name__ == '__main__':
    userlogin = Login()
    username = 'shaoxing_city'
    password = '123456'
    userlogin.setLoginInfo(username,password)
    userlogin.login()

    cookie = cookielib.CookieJar()
    # 利用urllib2库的HTTPCookieProcessor对象来创建cookie处理器
    handler = urllib2.HTTPCookieProcessor(cookie)
    # 通过handler来构建opener
    opener = urllib2.build_opener(handler)
    url ="http://122.224.174.179:8088/index.jsp"
    response = opener.open(url)
    for item in cookie:
        cookie_item = item.name + "=" + item.value
    print cookie_item