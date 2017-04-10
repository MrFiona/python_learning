#!/usr/bin/env python
# -*- coding: utf-8 -*-
# User    : apple
# File    : multi_process.py
# Author  : MrFiona 一枚程序员
# Time     : 2017-03-19 22:40

import requests
from requests.exceptions import HTTPError, ConnectionError
from bs4 import BeautifulSoup,NavigableString
from multiprocessing import Process, JoinableQueue
import time
#AyouBlog类#
# get_page_url函数获得所有博客的URL
class AyouBlog():
    def __init__(self):
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:48.0) Gecko/20100101 Firefox/48.0",}
        self.s = requests.session()
    def get_page_url(self):
        urls_set = []
        url="http://blog.csdn.net/u013055678?viewmode=contents"
        try:
            html = self.s.get(url, headers=self.headers)
        except HTTPError as e:
            # print(str(e))
            return str(e)
        except ConnectionError as e:
            # print(str(e))
            return str(e)
        try:
            # soup = BeautifulSoup(html.content, "lxml")
            # page_div = soup.find_all("span", {"class": "link_title"})
            # for url in page_div:
            #     a_url = "http://blog.csdn.net"+url.find("a").attrs["href"]
            #     urls_set.add(a_url)
            k = 0
            for i in ['http://www.iqiyi.com/'] * 100:
                k += 1
                urls_set.append(i)
                print k
        except AttributeError as e:
            # print(str(e))
            return str(e)
        return urls_set
#ThreadUrl继承进程类#run函数将JoinableQueue中的URL逐个取出,然后打开,取得博客详细页面的标题
class ThreadUrl(Process):
    def __init__(self, queue):
        super(ThreadUrl, self).__init__()
        self.queue = queue
        self.s = requests.session()
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:48.0) Gecko/20100101 Firefox/48.0",}
    def run(self):
        while not self.queue.empty():
            host = self.queue.get()
            try:
                html = self.s.get(host, headers=self.headers)
                print html
            except HTTPError as e:
                print(str(e))
                return str(e)
            except ConnectionError as e:
                print(str(e))
                return str(e)
            try:
                soup = BeautifulSoup(html.content, "lxml")

                # class_div = soup.find("span",{"class":"link_title"})
                # print((class_div.text).strip())
            except AttributeError as e:
                # print(str(e))
                return str(e)
            except NavigableString as e:
                # print(str(e))
                return str(e)
            self.queue.task_done()

def main():
    #进程列表
    worker_list = list()
    #创建队列
    queue = JoinableQueue()
    #将URL放进队列
    p = AyouBlog()
    for url in p.get_page_url():
        # print(url)
        queue.put(url)
    #开多进程
    for i in range(3):
        t = ThreadUrl(queue)
        worker_list.append(t)
        t.start()
        #队列清空后再执行其它
        queue.join()
        #进程关闭(这个是不是多余啊?)
        for w in worker_list:
            w.terminate()
if __name__=="__main__":
    start = time.time()
    main()
    print("Elapsed Time: %s" % (time.time() - start))
