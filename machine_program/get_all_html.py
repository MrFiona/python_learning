#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-03-21 09:40
# Author  : MrFiona
# File    : get_all_html.py
# Software: PyCharm Community Edition

import os
import re
import time
import codecs
import ghost
import urllib2
import chardet
import threading
import collections
from bs4 import BeautifulSoup
from extract_data import GetAnalysisData

#'https://dcg-oss.intel.com/ossreport/auto'
class GetUrlFromHtml(object):
    def __init__(self, html_url_pre=None, file_path = os.getcwd() + os.sep + 'report_html'):
        if not html_url_pre:
            raise UserWarning('Please send the url parameter!!!')
        if (html_url_pre and isinstance(html_url_pre, str) and html_url_pre.isspace()):
            raise UserWarning('All parameters cannot be an empty string!!!')
        if not isinstance(html_url_pre, str):
            raise UserWarning('All parameters must be a string!!!')

        self.html_url_pre = html_url_pre
        self.file_path = file_path
        self.department_list = []
        self.stage_type_list = []
        self.date_list = []
        self.url_info_list = []

    #递归创建目录函数  私有函数
    def _recursive_create_dir(self, dir_path=None):
        if not dir_path:
            raise UserWarning('请输入需要创建的目录路径!!!')
        #检测目录是否存在
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)

    #将指定目录中的文件写入数据  私有函数
    def _write_info(self, file_path, write_message):
        # print file_path
        try:
            if not file_path:
                raise UserWarning('请输入需要创建的目录路径!!!')
            #当输入的是目录而不是文件路径时则报错
            if os.path.isdir(file_path) and not os.path.isfile(file_path):
                raise UserWarning('您输入的待写入文件是目录!!!')
            with codecs.open(file_path, 'w', encoding='utf-8') as f:
                # print 'qqq', type(write_message)
                f.writelines(write_message.decode( chardet.detect(write_message)['encoding']))
        #当上一级以及以上目录不存在时则会报错
        except IOError as e:
            raise UserWarning('请检查 [ %s ] 目录是否存在!!!' % os.path.split(file_path)[0])

    #提取Html中指定的tag信息列表
    def _get_html_tag(self, html, tag):
        # print 'html', html
        soup = BeautifulSoup(html, 'html.parser')
        tag = soup.find_all(re.compile(tag))
        # print 'tag', tag
        # 去除Parent Directory这个无效选项
        tag.pop(0)
        # 按时间倒序输出
        tag.sort(reverse=True)
        return tag

    #将当前html写入对应路径的文件，返回含有li吧标签的信息列表
    def _common_get_info(self, html, create_html_name):
        # print 'tttt', html
        li_list = []
        try:
            # 写入html,需考虑多级路径的情况下处理，将路径分隔符替换为下划线,在连接日期的时候会出现%也要替换为下划线
            create_html_name = re.sub('[%_]', '_', create_html_name)
            response_html = urllib2.urlopen(html, timeout=6).read()
            # 递归创建存储信息的目录
            if create_html_name == 'auto':
                file_path = self.file_path
            else:
                file_path = self.file_path + os.sep + create_html_name
            self._recursive_create_dir(file_path)
            # print 'wwwwwwww', create_html_name
            self._write_info(file_path + os.sep +  '_'.join(create_html_name.split(os.sep)) + '.html', response_html)
            # 提取html信息
            li_list = self._get_html_tag(response_html, 'li')
            # print 'wee', li_list
        except urllib2.HTTPError as e:
            print e
            print '访问出错,请检查地址 %s 是否有效!!!' % html
            return
        except urllib2.URLError as e:
            print e
            print '超时'
            return
        return li_list

    #获取主页面html中的部门信息列表
    def get_department_html_info(self):
        li_list = self._common_get_info(self.html_url_pre, 'auto')
        if not li_list:
            return []
        department_info_list = []
        # print li_list
        for li in li_list:
            # print li
            if 'Purley-FPGA/' in li.encode('utf-8'):
                # print 'ddddd'
                department_info_list.append('Purley-FPGA')
            if 'Bakerville/' in li.encode('utf-8'):
                # print 'eeeee'
                department_info_list.append('Bakerville')
        department_info_list.sort(reverse=True)
        return department_info_list

    #获取部门html中的阶段信息列表
    def get_stage_type_html_info(self, department_name):
        li_list = self._common_get_info(self.html_url_pre + department_name, department_name)
        if not li_list:
            return []
        stage_info_list = []
        for li in li_list:
            if 'Silver/' in li.encode('utf-8'):
                stage_info_list.append('Silver')
            if 'Gold/' in li.encode('utf-8'):
                stage_info_list.append('Gold')
            if 'BKC/' in li.encode('utf-8'):
                stage_info_list.append('BKC')
            stage_info_list.sort(reverse=True)
        return stage_info_list

    #获取阶段html中的日期信息列表
    def get_date_html_info(self, department_name, stage_type_name):
        li_list = self._common_get_info(self.html_url_pre + department_name + '/' + stage_type_name, department_name + os.sep + stage_type_name)
        if not li_list:
            return []
        # 循环提取后续url前缀和文件名前缀
        url_pre_list = []
        file_pre_list = []
        li_list = list(li_list)
        for pre in li_list:
            href = re.search('<li><a\s+href=(?P<name>.*)>(?P<year>.*)</a></li>', str(pre))
            url_pre_list.append(href.groupdict()['name'].replace("\"""", ""))
            file_pre_list.append(href.groupdict()['year'])
            # print href.groupdict()

        # print li_list
        # print url_pre_list
        # print file_pre_list
        url_pre_list.sort(reverse=True)
        return url_pre_list

    #获取日期Html中的数据url地址
    def get_result_data_url(self, department_name, stage_type_name, date_name):
        li_list = self._common_get_info(self.html_url_pre + department_name + '/' + stage_type_name + '/' + date_name, department_name + os.sep + stage_type_name + os.sep + date_name)
        # print li_list
        if not li_list:
            return None
        file_name = None
        for li in li_list:
            file = re.findall('<li><.*>\s(.*?)</a></li>', li.encode('utf-8'))
            # print file
            if file[0].endswith('.html'):
                file_name = file[0]
                # print 'qqq', file_name
        return file_name

    def write_html_by_multi_thread(self):
        thread_list = []
        for url in self.url_info_list:
            thread = threading.Thread(target=GetAnalysisData, args=(url, ))
            thread_list.append(thread)
        for t in thread_list:
            t.start()

    #将url写进文件
    def save_url_info(self):
        with open(self.file_path + os.sep + 'url_info.txt', 'w') as f:
            f.write('\n'.join(self.url_info_list))

    #获取所有的类型的数据
    def get_all_type_data(self, get_info_not_save_flag=False):
        #获取部门列表
        self.department_list = self.get_department_html_info()
        if not self.department_list:
            print '部门列表为空!!!'
            return
        for department in self.department_list:
            #获取部门中的测试类型列表
            self.stage_type_list = self.get_stage_type_html_info(department)
            if not self.stage_type_list:
                print '测试类型列表为空!!!'
                return
            for stage in self.stage_type_list:
                self.date_list = self.get_date_html_info(department, stage)
                if not self.date_list:
                    print '日期列表为空!!!'
                    return
                for date in self.date_list:
                    file_name = self.get_result_data_url(department, stage, date.replace('/', ''))
                    if not file:
                        print 'html文件不存在!!!'
                        return
                    html_suffix = '/'.join([str(department), str(stage), str(date.replace('/', '')), str(file_name)])
                    # GetAnalysisData(self.html_url_pre + html_suffix)
                    #将url添加到url列表
                    self.url_info_list.append(self.html_url_pre + html_suffix)
                    print 'vvv', self.html_url_pre + html_suffix
                    #将源码写进文件
                    # object.save_html()
        #将url列表信息写进文件
        self.save_url_info()

    #获取部门, 测试类型, 日期列表
    def get_department_stage_date_list(self):
        print self.department_list, self.stage_type_list, self.date_list
        return self.department_list, self.stage_type_list, self.date_list

if __name__ == '__main__':
    start = time.time()
    object = GetUrlFromHtml('https://dcg-oss.intel.com/ossreport/auto/')
    object.get_all_type_data()
    #多线程写文件
    object.write_html_by_multi_thread()
    object.get_department_stage_date_list()
    print '用时: [ %d ]' %(time.time() - start)
    # hw = object.get_hw_configuration_info('BasinFalls', 'Gold111', '')
    # object.get_department_html_info()
    # object.get_stage_type_html_info('Purley-FPGA')
    # object.get_date_html_info('Purley-FPGA', 'Silver')
    # object.get_result_data_url('Purley-FPGA', 'Silver', '2017%20WW03')








    # row_lable_list = ['D2', 'E2', 'F2', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2']
    # common_formula_lable_list = [['D', 'Q'], ['E', 'R'], ['F', 'S'], ['G', 'T'], ['H', 'U'],
    #                              ['I', 'V'], ['J', 'W'], ['K', 'X'], ['L', 'Y'], ['M', 'Z']]
    #
    # for i in range(len(row_lable_list)):
    #     common_formula = '=OR(IF(OR(ISBLANK({0}3),ISBLANK({1}3)),FALSE,NOT(EXACT({0}3,{1}3))),IF(OR(ISBLANK({0}4),ISBLANK({1}4)),FALSE,NOT(EXACT({0}4,{1}4))),IF(OR(ISBLANK({0}5),ISBLANK({1}5)),FALSE,NOT(EXACT({0}5,{1}5))),IF(OR(ISBLANK({0}6),ISBLANK({1}6)),FALSE,NOT(EXACT({0}6,{1}6))),IF(OR(ISBLANK({0}7),ISBLANK({1}7)),FALSE,NOT(EXACT({0}7,{1}7))),IF(OR(ISBLANK({0}8),ISBLANK({1}8)),FALSE,NOT(EXACT({0}8,{1}8))),IF(OR(ISBLANK({0}9),ISBLANK({1}9)),FALSE,NOT(EXACT({0}9,{1}9))))'
    #     insert_row_formula_list = '=OR(IF(OR(ISBLANK({0}3),ISBLANK({1}3)),FALSE,NOT(EXACT({0}3,{1}3))),IF(OR(ISBLANK({0}4),ISBLANK({1}4)),FALSE,NOT(EXACT({0}4,{1}4))),IF(OR(ISBLANK({0}5),ISBLANK({1}5)),FALSE,NOT(EXACT({0}5,{1}5))),IF(OR(ISBLANK({0}6),ISBLANK({1}6)),FALSE,NOT(EXACT({0}6,{1}6))),IF(OR(ISBLANK({0}7),ISBLANK({1}7)),FALSE,NOT(EXACT({0}7,{1}7))),IF(OR(ISBLANK({0}8),ISBLANK({1}8)),FALSE,NOT(EXACT({0}8,{1}8))),IF(OR(ISBLANK({0}9),ISBLANK({1}9)),FALSE,NOT(EXACT({0}9,{1}9))))'.format(
    #         common_formula_lable_list[i][0], common_formula_lable_list[i][1])
    #     print insert_row_formula_list
    #     worksheet.write_row('C2', ['Changed from previous configuration?'], insert_row_formula_list)