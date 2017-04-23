#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-17 10:48
# Author  : MrFiona
# File    : extract_bakerville_data.py
# Software: PyCharm Community Edition

import os
import sys
import re
import copy
import codecs
import chardet
import socket
import ghost  #需要安装依赖包：Pyside或者PyQt4   安装pip install pyside即可
import urllib2
import collections
import HTMLParser
import collections
import htmlentitydefs
from bs4 import BeautifulSoup
from cache_mechanism import DiskCache

reload(sys)
sys.setdefaultencoding('utf-8')

class GetAnalysisData(object):
    def __init__(self, data_url, get_info_not_save_flag=False, cache=None):
        self.data_url = data_url
        self.cache = cache
        try:
            result = None
            if self.cache:
                try:
                    result = self.cache[self.data_url]
                except KeyError:
                    pass

            if result is None:
                result = urllib2.urlopen(self.data_url).read()
                if self.cache:
                    self.cache[self.data_url] = result
            # print 'data_url:\t', self.data_url
            self.save_file_name = os.path.split(self.data_url)[-1].split('.')[0] + '_html'
            temp_path = self.data_url.split('auto')[1].strip('/').split('/')[:-1] #去掉最后一个html文件名
            self.grandfather_path = temp_path[0]              #部门类型级别
            self.father_path = temp_path[1]                   #测试类型级别
            self.son_path = re.sub('[%]', '_', temp_path[-1]) #日期级别
            self.save_file_path = os.getcwd() + os.sep + 'report_html' + os.sep +'%s' % os.sep.join([ self.grandfather_path, self.father_path, self.son_path ])
            url_split_list = self.data_url.split('/')[-2]
            self.date_string = url_split_list[-4:]
        except IndexError, e:
            raise UserWarning('请检查url: [ %s ] 是否输入完整!!!' %self.data_url)
        #保存文件标记开启，将html代码保存为文件
        if get_info_not_save_flag:
            self.save_html()

    def save_html(self):
        try:
            from test import FilterTag
            filters = FilterTag()
            html = urllib2.urlopen(self.data_url).read().decode('utf-8')
            response_parser = HTMLParser.HTMLParser()
            html = response_parser.unescape(html)
            html = filters.filterHtmlTag(html)
            html = filters.replaceCharEntity(html)
            # 将源代码写入文件
            # print type(html)
            with codecs.open(self.save_file_path + os.sep + self.save_file_name, 'w', encoding='utf-8') as file:
                file.write(html)
        except urllib2.HTTPError as e:
            print e
            print '访问出错,请检查地址 %s 是否有效!!!' % self.data_url
            return
        except urllib2.URLError as e:
            print e
            print '超时'
            return

    # 去除特殊字符策略函数
    def _remove_non_alphanumeric_characters(self, object_string_list):
        # 对提取的字符串列表进行清洗，统一组合格式：空格分隔  Mon Apr 10 14:00:27 2017
        for ele in range(len(object_string_list)):
            object_string_list[ele] = re.sub('[\s]', 'MrFiona', object_string_list[ele])
            object_string_list[ele] = re.sub('[^\w\"\'\.\_\-\>\[\]\(\)\@\~\,\*]', '', object_string_list[ele])
            temp_ele_list = re.split('MrFiona', object_string_list[ele])
            ele_list = [effective for effective in temp_ele_list if len(effective) != 0]
            object_string_list[ele] = ' '.join(ele_list)
        return object_string_list

        # 提取通用的部分

    def _common_regex(self, data_type, bkc_flag=False, header_list=None, cell_data_list=None):
        if not header_list:
            header_list = []
        if not cell_data_list:
            cell_data_list = []
        # print data_type, bkc_flag
        if not os.path.exists(self.save_file_path):
            raise UserWarning('%s 目录不存在!!!' % self.save_file_path)
        # BKC标记开启，自动更换为文件BKC地址
        if bkc_flag:
            print self.save_file_path
            self.save_file_path = self.save_file_path.replace('Silver', 'BKC')
            # print self.save_file_path
            # if os.path.exists(self.save_file_path):
            if not os.path.exists(self.save_file_path):
                # raise UserWarning('%s 目录不存在!!!' % self.save_file_path)
                return [], [], []
            search_file_list = os.listdir(self.save_file_path)
            # print search_file_list
            bkc_file_name = ''
            for file in search_file_list:
                match = re.match(r'\d+', file)
                if match:
                    bkc_file_name = file
            if len(bkc_file_name) == 0:
                print 'BKC文件不存在'
                return [], [], []
            self.save_file_name = bkc_file_name

        # 读取文本中的html代码
        with codecs.open(self.save_file_path + os.sep + self.save_file_name, 'r', encoding='utf-8') as f:
            data = f.readlines()
        data = ''.join(data)
        # data = self.cache[self.data_url]
        # 提取HW Configuration部分的代码
        regex = re.compile(r'<span class="sh2">&nbsp; %s </span>(.*?)<div class="panel-heading">' % data_type,
                           re.S | re.M)
        header = re.findall(regex, data)
        string_data = ''.join(header)
        # 提取所有的tr部分
        soup_tr = BeautifulSoup(string_data, 'html.parser')
        tr_list = soup_tr.find_all(re.compile('tr'))
        # 增加保险措施，通过关键字提取文本html实现
        if not tr_list:
            # print '提取的tr_list为空!开始通过文本提取匹配内容......'
            file_path = self.save_file_path + os.sep + self.save_file_name
            # print file_path
            if bkc_flag:
                print '\033[43mBKC链接:\t\033[0m', self.data_url.replace('Silver', 'BKC')
            fread = codecs.open(self.save_file_path + os.sep + self.save_file_name, 'r', 'utf-8')
            fwrite = codecs.open(self.save_file_path + os.sep + 'temp.txt', 'wb', 'utf-8')
            pre_keyword = data_type
            next_keyword = 'panel collapse-report-panel'
            flag = False
            for line in fread:
                if flag:
                    fwrite.write(line)
                if pre_keyword in line:
                    fwrite.write(line)
                    flag = True
                if next_keyword in line:
                    flag = False
            fread.close()
            fwrite.close()

            with codecs.open(self.save_file_path + os.sep + 'temp.txt', 'r', 'utf-8') as f:
                lines = f.readlines()

            lines = ''.join(lines)
            # print lines, type(lines), chardet.detect(lines.encode('utf-8'))
            soup = BeautifulSoup(lines, 'html.parser')
            tr_list = soup.find_all('tr')
            # print '文本提取匹配内容结束......'

        return tr_list, header_list, cell_data_list

        # 获取hw有效的列名列表

    def _get_hw_effective_header_list(self, effective_header_list):
        header_list = []
        # 提取有效的行数
        effective_num_list = []
        header_regex = re.compile('[Ss]ystem(?P<name>\(?.*\)?)', re.M | re.S)
        element_regex = re.compile('\d+')
        range_num_regex = re.compile('\d+?\,?~\d+')
        #判断周表头列是否含有SKX-DE
        SKX_DE_FLAG = False
        for label in range(len(effective_header_list)):
            temp_num_list = []
            # print effective_header_list[label]
            temp_string_list = re.split(header_regex, effective_header_list[label])
            remove_line_break(temp_string_list)
            self._remove_non_alphanumeric_characters(temp_string_list)
            header_list.append(temp_string_list[0])
            #进一步提取表头列数
            #排除类似这种情况 SKX-DE(001,~004)(Cycling Test)
            if temp_string_list[0].startswith('SKX-DE'):
                SKX_DE_FLAG = True
                continue
            #1、数字之间逗号相隔且不含油连续数字符号~
            if '~' not in temp_string_list[0]:
                signal_num_list = re.findall(element_regex, temp_string_list[0])
                if signal_num_list:
                    signal_num_list = [int(ele) for ele in signal_num_list]
                    temp_num_list.extend(signal_num_list)
            #含有连续数字符号~  (1,~5)  (1~5)  1~5
            else:
                if ',' not in temp_string_list[0]:
                    #不含有逗号只含有~
                    range_num_list = re.findall('\d+', temp_string_list[0])
                    if range_num_list:
                        #转换为离散的数字列
                        temp_change_list = [nu for nu in range(int(range_num_list[0]), int(range_num_list[1]) + 1)]
                        temp_num_list.extend(temp_change_list)
                #既含有逗号又含有~ (1,~5) 1,~5 (1,3~6) 1,2~5
                else:
                    #去括号
                    temp_string = re.sub('[()]', '', temp_string_list[0])
                    judge_string_list = re.findall('\d+~\d+', temp_string)
                    #说明是 (1,2~7)这样的情况 ~左右数字齐全
                    # print 'judge_string_list:\t',judge_string_list
                    if judge_string_list:
                        temp_list = []
                        split_string_list = [ele for ele in re.split(',', temp_string) if len(ele) != 0]
                        for deal_num in split_string_list:
                            #2~6
                            if '~' in deal_num:
                                range_num_list = re.findall('\d+', deal_num)
                                temp_list = [nu for nu in range(int(range_num_list[0]), int(range_num_list[1]) + 1)]
                            #单个数字 2
                            else:
                                temp_list.append(deal_num)
                        temp_num_list.extend(temp_list)
                    #说明是 (1,~5)这种情况
                    else:
                        range_num_list = re.findall('\d+', temp_string)
                        if range_num_list:
                            # 转换为离散的数字列
                            temp_change_list = [nu for nu in range(int(range_num_list[0]), int(range_num_list[1]) + 1)]
                            temp_num_list.extend(temp_change_list)
            if temp_num_list:
                effective_num_list.append(temp_num_list)

            # print 'temp_num_list:\t', temp_num_list
        # print 'header_list:\t', header_list
        # print 'effective_num_list:\t',effective_num_list
        effective_header_list = [int(ele) for child_element in effective_num_list for ele in child_element]
        effective_header_list = sorted(effective_header_list, key=lambda x: x)
        #SKX_DE标记开启,自动填充至12 表示12列数据
        if SKX_DE_FLAG:
            for i in range(len(effective_header_list)+1, 12+1):
                effective_header_list.append(i)
        effective_header_list = ['System ' + str(ele) for ele in effective_header_list]
        # print 'effective_num_list:\t', effective_num_list
        # print 'effective_header_list:\t', effective_header_list

        return SKX_DE_FLAG, effective_header_list, effective_num_list

    # 判断插入到cell_data_list的次数并插入数据
    # 根据表头列数字和cell_data_list进一步处理,生成新的cell_data_list,effective_num_list
    def _insert_numers_to_cell_data_list(self, SKX_DE_FLAG, effective_num_list, cell_data_list):
        cell_data_direction_dict = collections.OrderedDict()
        for ele in range(len(effective_num_list)):
            for k in effective_num_list[ele]:
                temp_cell_data_list = []
                for j in range(len(cell_data_list)):
                    temp_cell_data_list.append(cell_data_list[j][ele])
                cell_data_direction_dict[k] = temp_cell_data_list
        cell_data_direction_dict = sorted(cell_data_direction_dict.iteritems())
        x, effective_cell_data_list = zip(*cell_data_direction_dict)
        effective_cell_data_list = [[ele[i] for ele in effective_cell_data_list] for i in range(len(cell_data_list))]
        #重复拷贝信息的次数
        repeat_append_num = 12 - len(effective_cell_data_list[0])
        #不管是否含有SKX_DE，只要小于12列后面的列就用最后一列信息补充
        if repeat_append_num > 0:
            for m in range(len(cell_data_list)):
                for nu in range(repeat_append_num):
                    effective_cell_data_list[m].append(cell_data_list[m][-1])

        return effective_cell_data_list

    def get_hw_data(self, data_type, bkc_flag=True):
        tr_list, header_list, cell_data_list = self._common_regex(data_type, bkc_flag)
        if not tr_list:
            return self.date_string, [], [], []
        # 循环处理tr中的代码，提取出excel表头信息，存放在表头列表header_list
        td_list = []
        effective_header_list = []

        soup_tr = BeautifulSoup(str(tr_list[0]), 'html.parser')
        th_list = soup_tr.find_all('th')
        td_list = soup_tr.find_all('td')
        effective_th_or_td_list = th_list if th_list else td_list
        for th in effective_th_or_td_list[1:]:
            soup_th = BeautifulSoup(str(th), 'html.parser')
            th_string_list = list(soup_th.strings)
            # 去掉换行符元素
            remove_line_break(th_string_list)
            effective_header_list.append(th_string_list[0])
        max_length = len(effective_th_or_td_list)

        for tr in tr_list[1:]:
            temp = []
            add_header_flag = False
            #tr中可能是th也可能是td标签,最后两个tr是合并项
            soup_tr = BeautifulSoup(str(tr), 'html.parser')
            #有可能出现tr出即出现td又出现th的情况
            th_list = soup_tr.find_all('th')
            td_list = soup_tr.find_all('td')
            effective_th_or_td_list = th_list if th_list else td_list
            compare_length = len(effective_th_or_td_list)
            if compare_length < max_length and not ( td_list and th_list):
                add_header_flag = True
            if not add_header_flag:
                soup_string = BeautifulSoup(str(effective_th_or_td_list[0]), 'html.parser')
                string_list = list(soup_string.strings)
                string_list = remove_line_break(string_list)
                string_list = self._remove_non_alphanumeric_characters(string_list)
                header_list.append(string_list[0])

            if not (td_list and th_list):
                if compare_length < max_length:
                    for_list = effective_th_or_td_list
                else:
                    for_list =effective_th_or_td_list[1:]
            else:
                for_list = td_list
            for ele in for_list:
                cell_soup = BeautifulSoup(str(ele), 'html.parser')
                cell_string_list = list(cell_soup.strings)
                cell_string_list = remove_line_break(cell_string_list)
                cell_string_list = self._remove_non_alphanumeric_characters(cell_string_list)
                temp.append(cell_string_list[0])
            cell_data_list.append(temp)

        # print effective_header_list
        # 提取有效的表头列数
        SKX_DE_FLAG, effective_header_list, effective_num_list = self._get_hw_effective_header_list(effective_header_list)
        cell_data_list = self._insert_numers_to_cell_data_list(SKX_DE_FLAG, effective_num_list, cell_data_list)

        print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        print '\033[32meffective_header_list:\t\033[0m', effective_header_list, len(effective_header_list)
        return self.date_string, effective_header_list, header_list, cell_data_list

def get_url_list_by_keyword(pre_keyword, back_keyword, key_url_list=None):
    if not key_url_list:
        key_url_list = []

    f = open(os.getcwd() + os.sep + 'report_html' + os.sep + 'url_info.txt')
    for line in f:
        if pre_keyword in line and back_keyword in line:
            key_url_list.append(line.strip('\n'))
    # print key_url_list, len(key_url_list)
    return key_url_list

def remove_line_break(object_string_list):
    try:
        while 1:
            object_string_list.remove('\n')
    except ValueError:
        pass

    try:
        while 1:
            object_string_list.remove('')
    except ValueError:
        pass

    return object_string_list


if __name__ == '__main__':
    import time
    start = time.time()
    # url = 'https://dcg-oss.intel.com/ossreport/auto/Bakerville/Silver/2017%20WW15/5848_Silver.html'
    url_list = get_url_list_by_keyword('Bakerville', 'Silver')
    # print url_list, len(url_list)
    cache = DiskCache()
    for url in url_list:
        obj = GetAnalysisData(url, get_info_not_save_flag=False, cache=cache)
        # obj.save_html()
        obj.get_hw_data('HW Configuration')
    print time.time() - start