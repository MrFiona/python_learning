#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import urllib
import urlparse
import urllib2
from urllib2 import URLError, urlopen, Request
def download(url, user_agent='wswp', proxy=None, num_retries=2):
    print 'Begain Download the url:[ %s ]', url
    #设置代理proxy,设定一个默认的用户代理'wswp'
    headers = {'User-agent':user_agent}
    request = Request(url, headers=headers)
    opener = urllib2.build_opener()
    if proxy:
        proxy_params = {urlparse.urlparse(url).scheme:proxy}
        opener.add_handler(urllib2.ProxyBasicAuthHandler(proxy_params))
    try:
        response = urlopen(request, timeout=500).read()
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
m = download('http://httpstat.us/500')
print m