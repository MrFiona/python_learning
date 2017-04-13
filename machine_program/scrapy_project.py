#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import urllib
import urlparse
import urllib2
from urllib2 import URLError
from bs4 import BeautifulSoup

def download(url, user_agent='wswp', proxy=None, num_retries=2):
    print 'Begain Download the url:[ %s ]', url
    #设置代理proxy,设定一个默认的用户代理'wswp'
    headers = {'User-agent':user_agent}
    request = urllib2.Request(url, headers=headers)
    opener = urllib2.build_opener()
    if proxy:
        proxy_params = {urlparse.urlparse(url).scheme:proxy}
        opener.add_handler(urllib2.ProxyBasicAuthHandler(proxy_params))
    try:
        response = urllib2.urlopen(request, timeout=500).read()
    except URLError as e:
        print 'Download Error:', e.reason
        print e.errno, e.message, e.args, e.reason, e.strerror
        response = None
        if num_retries > 0:
            if hasattr(e, 'code') and 500 <= e.errno <= 600:
                download(url, user_agent, proxy, num_retries - 1)
    return response

def crawl_sitemap(url):
    sitemap = download(url)
    links = re.findall('<loc>(.*?)</loc>', sitemap)
    for link in links:
        html = download(link)
        print html

def get_links(html):
    webpage_regex = re.compile(r'<a[^>]+href=["\'](.*?)["\']',re.IGNORECASE)
    return webpage_regex.findall(html)

def link_crawler(seed_url, link_regex):
    crawl_queue = [seed_url]
    seen = set(crawl_queue)
    while crawl_queue:
        url = crawl_queue.pop()
        html = download(url)
        for link in get_links(html):
            if re.match(link_regex, link):
                link = urlparse.urljoin(seed_url, link)
                if link not in seen:
                    seen.add(link)
                    crawl_queue.append(link)

# link_crawler('http://example.webscraping.com/index', '/(index|view)')
# crawl_sitemap('http://example.webscraping.com/sitemap.xml')
# m = download('http://httpstat.us/500')
# print m

response = urllib2.urlopen('http://www.iqiyi.com/').read()
# print response.read()

# print re.findall(r'<li class=.*>', response)

soup = BeautifulSoup(response, 'html.parser')

i = 0
# for li in soup.find_all('li'):
#     i += 1
    # print 'i = %d：%s' %(i, li)

import cookielib, urllib2
# req = urllib2.Request('https://mail.qq.com/', 'dasdqwe')
# req.add_header('User-Agent', 'Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)')
# ckjar = cookielib.MozillaCookieJar('ff.txt')
# ckproc = urllib2.HTTPCookieProcessor(ckjar)
# opener = urllib2.build_opener(ckproc)
# f = opener.open(req)
# htm = f.read()
# f.close()
# ckjar.save(ignore_discard=True, ignore_expires=True)

#创建Opener----------------------------
#创建cookie对象
cookie = cookielib.CookieJar()
#创建COOKIE处理程序
cookieProc = urllib2.HTTPCookieProcessor(cookie)
#创建opener
opener = urllib2.build_opener(cookieProc)
#安装到urlopen()(这里也可以不用install_opener)
urllib2.install_opener(opener)

#发起请求------------------------------
#设置请求参数
postdata = {
    'username':'name',
    'password':'password',
    'type':1
    }
postdata = urllib.urlencode(postdata)

#注：这里还可以这样写：
# 直接把 postdata = 'username=xxx@163.com&password=xxxx&type=1'

#设置请求header
header = {'User-Agent':'Mozilla/5.0 (Windows NT 6.3; Win64; x64; rv:50.0) Gecko/20100101 Firefox/50.0'}
#Request
req = urllib2.Request(
     url = 'https://mail.qq.com/cgi-bin/frame_html?sid=xkUypUMyVnCxfDlh&r=e5af7affffbee1c656db5d7cc27fbd17',
     data = postdata,
     headers = header
    )
#请求
res = urllib2.urlopen(req).read()

#如果上面代码没有install_opener()，则这里可用:opener.open(req).read()方法来获取内容

print res


