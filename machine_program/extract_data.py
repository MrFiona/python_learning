#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-03-21 09:40
# Author  : MrFiona
# File    : extract_data.py
# Software: PyCharm Community Edition

import os
import sys
import re
import codecs
import urllib2
import HTMLParser
import htmlentitydefs
from bs4 import BeautifulSoup
from cache_mechanism import DiskCache
from public_use_function import verify_validity_url, FilterTag

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

    #获取url详细信息
    def get_info_detail(self):
        return self.data_url

    def decodeHtmlEntity(self, origHtml, decodedEncoding=""):
        decodedEntityName = re.sub('&(?P<entityName>[a-zA-Z]{2,10});',
                                   lambda matched: unichr(htmlentitydefs.name2codepoint[matched.group("entityName")]),
                                   origHtml)
        decodedCodepointInt = re.sub('&#(?P<codePointInt>\d{2,5});',
                                     lambda matched: unichr(int(matched.group("codePointInt"))), decodedEntityName)
        decodedCodepointHex = re.sub('&#x(?P<codePointHex>[a-fA-F\d]{2,5});',
                                     lambda matched: unichr(int(matched.group("codePointHex"), 16)),
                                     decodedCodepointInt)
        decodedHtml = decodedCodepointHex
        if (decodedEncoding):
            decodedHtml = decodedHtml.encode(decodedEncoding, 'ignore')
        return decodedHtml

    def save_html(self):
        try:
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
            return

    #去除特殊字符策略函数
    def _remove_non_alphanumeric_characters(self, object_string_list):
        # 对提取的字符串列表进行清洗，统一组合格式：空格分隔  Mon Apr 10 14:00:27 2017
        for ele in range(len(object_string_list)):
            object_string_list[ele] = re.sub('[\s]', 'MrFiona', object_string_list[ele])
            object_string_list[ele] = re.sub('[^\w\"\'\.\_\-\>\[\]\(\)\@\~\/\*]', '', object_string_list[ele])
            temp_ele_list = re.split('MrFiona', object_string_list[ele])
            ele_list = [effective for effective in temp_ele_list if len(effective) != 0]
            object_string_list[ele] = ' '.join(ele_list)
        return object_string_list

    #提取通用的部分,如果bkc_flag=True并且BKC和Gold都没有数据则取Silver数据 优先级：BKC > Gold > Silver
    def _common_regex(self, data_type, bkc_flag=False, replace_flag=False, header_list=None, cell_data_list=None):
        Silver_Gold_BKC_string = 'Silver'
        if not header_list:
            header_list = []
        if not cell_data_list:
            cell_data_list = []
        # 如果bkc_flag = True并且BKC没有数据则取Silver数据,需将变量恢复原值
        if replace_flag and not bkc_flag:
            self.save_file_path = self.save_file_path.replace('BKC', 'Silver')
            self.save_file_path = self.save_file_path.replace('Gold', 'Silver')
            self.data_url = self.data_url.replace('BKC', 'Silver')
            self.data_url = self.data_url.replace('Gold', 'Silver')
            self.save_file_name = self.save_file_name.replace('BKC', 'Silver')
            self.save_file_name = self.save_file_name.replace('Gold', 'Silver')
            Silver_Gold_BKC_string = 'Silver'

        # if not os.path.exists(self.save_file_path):
        #     raise UserWarning('%s 目录不存在!!!' % self.save_file_path)
        #BKC标记开启，自动更换为文件BKC地址
        if bkc_flag:
            self.save_file_path = self.save_file_path.replace('Silver', 'BKC')
            self.data_url = self.data_url.replace('Silver', 'BKC')
            Silver_Gold_BKC_string = 'BKC'
            # 如果BKC无数据则取Gold数据
            if not os.path.exists(self.save_file_path):
                # raise UserWarning('%s 目录不存在!!!' % self.save_file_path)
                self.save_file_path = self.save_file_path.replace('BKC', 'Gold')
                self.data_url = self.data_url.replace('BKC', 'Gold')
                Silver_Gold_BKC_string = 'Gold'
                if not os.path.exists(self.save_file_path):
                    return '', [], [], []

            search_file_list = os.listdir(self.save_file_path)
            bkc_file_name = ''
            for file in search_file_list:
                match = re.match(r'\d+', file)
                if match:
                    bkc_file_name = file
            if len(bkc_file_name) == 0:
                print 'BKC以及Gold文件不存在'
                return '', [], [], []
            self.save_file_name = os.path.split(self.data_url)[-1].split('.')[0] + '_html'
            self.save_file_name = bkc_file_name

        # 读取文本中的html代码
        # with codecs.open(self.save_file_path + os.sep + self.save_file_name, 'r', encoding='utf-8') as f:
        #     data = f.readlines()
        # data = ''.join(data)
        data = self.cache[self.data_url]
        # 提取HW Configuration部分的代码
        regex = re.compile(r'<span class="sh2">&nbsp; %s </span>(.*?)<div class="panel-heading">' % data_type, re.S | re.M)
        header = re.findall(regex, data)
        string_data = ''.join(header)
        # 提取所有的tr部分
        soup_tr = BeautifulSoup(string_data, 'html.parser')
        tr_list = soup_tr.find_all(re.compile('tr'))

        #增加保险措施，通过关键字提取文本html实现
        if not tr_list:
            # if bkc_flag:
            #     print '\033[43mBKC链接:\t\033[0m', self.data_url.replace('Silver', 'BKC')
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
            soup = BeautifulSoup(lines, 'html.parser')
            tr_list = soup.find_all('tr')

        return Silver_Gold_BKC_string, tr_list, header_list, cell_data_list

    #获取hw有效的列名列表
    def _get_hw_effective_header_list(self, effective_header_list):
        # 提取有效的行数
        effective_num_list = []
        prefix_suffix_list = []
        for label in range(len(effective_header_list)):
            temp_string_list = re.split('[\D+]', effective_header_list[label])
            temp = re.split('[\d+~]', effective_header_list[label])
            # 去除空字符串
            try:
                space_num = temp.count('')
                for num in range(space_num):
                    temp.remove('')
            except ValueError as e:
                print e
                raise UserWarning('删除空格出现错误!!!')
            prefix_suffix_list.append(temp)
            temp_list = list(set(temp_string_list))[1:]
            for i in range(len(temp_list)):
                temp_list[i] = int(temp_list[i])
                temp_list.sort()
            effective_num_list.append(temp_list)

        effective_temp_list = []
        for pre, head in zip(prefix_suffix_list, effective_num_list):
            effective_temp_list.append([pre, head])

        effective_header_list = []
        for node_list in effective_temp_list:
            try:
                range_col = node_list[1]
                left_half_section = range_col[0]
                try:
                    right_half_section = range_col[1] + 1
                except IndexError:
                    right_half_section = left_half_section + 1

                for col in range(left_half_section, right_half_section):
                    if len(node_list[0]) == 1:
                        effective_header_list.append(node_list[0][0] + str(col))
                    if len(node_list[0]) > 1:
                        effective_header_list.append(node_list[0][0] + str(col) + node_list[0][-1])
            except (ValueError, IndexError) as e:
                print e
        return effective_header_list, effective_num_list

    #判断插入到cell_data_list的次数并插入数据
    def _insert_numers_to_cell_data_list(self, cell_data_list):
        try:
            effective_header_list, effective_num_list = self._get_hw_effective_header_list(cell_data_list[0])
            for node_list in effective_num_list:
                left_half_section = node_list[0]
                try:
                    right_half_section = node_list[1] + 1
                except IndexError:
                    right_half_section = left_half_section + 1
                for col in range(left_half_section-1, right_half_section-2):
                    for i in range(len(cell_data_list)):
                        if col >= len(cell_data_list[i]) - 1:
                            cell_data_list[i].append( cell_data_list[i][-1] )
                        elif col < len(cell_data_list[i]) - 1:
                            cell_data_list[i].insert( col, cell_data_list[i][col] )
        except (ValueError, IndexError) as e:
            print e
        return cell_data_list

    # @do_cprofile('./mkm_run_hw.prof')
    def get_hw_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)

        if not tr_list:
            return Silver_Gold_BKC_string, self.date_string, [], [], []
        # 循环处理tr中的代码，提取出excel表头信息，存放在表头列表header_list
        effective_header_list = []
        for t in tr_list:
            soup_td = BeautifulSoup(str(t), 'html.parser')
            # 排除其他非td杂项
            if soup_td.td != None:
                # 排除td中无字符串的情况
                if soup_td.td.strings != None:
                    left_string_list = list(soup_td.td.strings)
                    # 适配，去除字符串列表中的多余字符   最终列表长度为1
                    if '\n' in left_string_list:
                        left_string_list.remove('\n')
                    data = left_string_list[0].encode('utf-8')
                    # 去除header_list元素中换行符
                    data = data.replace('\n', '')
                    # 排除字符串部分为非字母数字的td
                    header_list.append(data)
            td_data_list = soup_td.find_all('td')
            # 获取td中的字符串即待插入excel表中单元格中的数据部分
            temp = []
            # print td_data_list
            for td in td_data_list:
                get_td_data = BeautifulSoup(str(td), 'html.parser')
                td_string_list = list(get_td_data.td.strings)
                break_num = td_string_list.count('\n')
                try:
                    for i in range(break_num):
                        td_string_list.remove('\n')
                except IndexError:
                    pass
                td_data = td_string_list[0].encode('utf-8')
                # 去除cell_data_list元素中的非字母编码部分
                td_data = re.sub('[\xc2\xa0\xc3\x82]', '', td_data)
                temp.append(td_data)
                # 获取有效的列表明列表
                if t == tr_list[0]:
                    if td != td_data_list[0]:
                        effective_header_list.append(td_data)
                        effective_header_list, effective_num_list = self._get_hw_effective_header_list(effective_header_list)
            if temp:
                temp = self._remove_non_alphanumeric_characters(temp)
                cell_data_list.append(temp[1:])
            #适配html表列名对应行中含有标签不含有td标签
            if t == tr_list[0]:
                # print tr_list[0]
                if not effective_header_list:
                    th_soup = BeautifulSoup(str(tr_list[0]), 'html.parser')
                    th_list = th_soup.find_all('th')
                    # print th_list
                    th_list = th_list[1:]
                    for word in th_list:
                        string_soup = BeautifulSoup(str(word), 'html.parser')
                        th_strings_list = list(string_soup.strings)
                        th_strings_list = [ effective_string for effective_string in th_strings_list if len(effective_string) > 1 ]
                        effective_header_list.append(th_strings_list[0])
                        cell_data_list.insert(0, effective_header_list)
                effective_header_list, effective_num_list = self._get_hw_effective_header_list(effective_header_list)

        cell_data_list = self._insert_numers_to_cell_data_list(cell_data_list)
        header_list = self._remove_non_alphanumeric_characters(header_list)
        if cell_data_list:
            cell_data_list[0] = effective_header_list
        # header_cell_data_info = collections.OrderedDict()
        # [header_cell_data_info.update({header_list[key]: cell_data_list[key]}) for key in range(len(header_list))]
        header_list = header_list[ -(len(cell_data_list) -1): ]
        try:
            if header_list[0].startswith('Ingredient'):
                header_list = header_list[1:]
        except IndexError:
            pass
        cell_data_list = cell_data_list[ -(len(header_list)): ]
        # 行坐标列表
        # print effective_header_list, len(effective_header_list)
        # 列坐标列表
        # print '\033[32meffective_header_list:\t\033[0m', effective_header_list, len(effective_header_list)
        # print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        return Silver_Gold_BKC_string, self.date_string, effective_header_list, header_list, cell_data_list

    def judge_silver_bkc_func(self, data_type, bkc_flag):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self._common_regex(data_type, bkc_flag)
        # 如果bkc_flag = True并且BKC没有数据则取Silver数据
        if not tr_list and bkc_flag:
            Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self._common_regex(data_type, bkc_flag=False, replace_flag=True)
        return  Silver_Gold_BKC_string, tr_list, header_list, cell_data_list

    # @do_cprofile('./mkm_run_sw.prof')
    def get_sw_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)
        if not tr_list:
            return Silver_Gold_BKC_string, 0, self.date_string, [], [], [], [], []
        soup_header = BeautifulSoup(str(tr_list[0]), 'html.parser')
        header_list = list(soup_header.strings)
        num = header_list.count('\n')
        for j in range(num):
            header_list.remove('\n')
        for i in range(len(header_list)):
            header_list[i] = header_list[i].replace('\n', '')
        #排除部分周会有更多的列
        header_list = header_list[:4]
        header_length = len(header_list)
        difference_length = len(header_list) - len(header_list[:4])
        #左边的第一列
        left_col_list_1 = []
        #左边的第二列
        left_col_list_2 = []
        head_constant = ''
        head_constant_1 = ''
        url_list = []
        for t in tr_list[1:]:
            qstring =''
            temp = []
            temp_url_list = []
            soup_tr = BeautifulSoup(str(t), 'html.parser')
            td_list = soup_tr.find_all('td')
            if difference_length >= 1:
                td_list = td_list[:-difference_length]
            #默认情况是正常列数
            td_list = td_list
            actual_td_list = td_list
            num = len(td_list)
            #第一种情况：左边出现行和列都合并的情况，此时解析时第一个td是多余的，去掉
            if num == header_length + 1:
                soup_5 = BeautifulSoup(str(t), 'html.parser')
                string_5_list = list(soup_5.strings)
                nu = string_5_list.count('\n')
                for i in range(nu):
                    string_5_list.remove('\n')
                head_constant_1 = string_5_list[0]
                actual_td_list = td_list[1:]

            if num == header_length + 2:
                actual_td_list = td_list[2:]
                soup_6 = BeautifulSoup(str(t), 'html.parser')
                string_6_list = list(soup_6.strings)
                nu = string_6_list.count('\n')
                for i in range(nu):
                    string_6_list.remove('\n')
                head_constant, head_constant_1 = string_6_list[0], string_6_list[1]

            if num == header_length:
                soup_4 = BeautifulSoup(str(t), 'html.parser')
                string_4_list = list(soup_4.strings)
                nu = string_4_list.count('\n')
                for i in range(nu):
                    string_4_list.remove('\n')
            # print actual_td_list
            for td in actual_td_list:
                soup = BeautifulSoup(str(td), 'html.parser')
                string_list = list(soup.strings)
                #正则取匹配url链接
                obj_list = re.findall('<a href="(.*?)">', str(td), re.M|re.S)
                if obj_list:
                    #逐个添加
                    for url in obj_list:
                        url = url.split()[0].replace("\"", "")
                        temp_url_list.append(url)
                #非空
                if string_list:
                    for ele in string_list:
                        if ele == '\n':
                            string_list.remove('\n')
                    for i in range(len(string_list)):
                        string_list[i] = string_list[i].replace('\n', '')
                # print 'string_list;\t', string_list
                # string_list[0] = string_list[0].replace('\n', '')
                #可能会含有多个的url链接，需要添加整个列表string_list，防止丢失数据
                for ele in string_list:
                    ele = ele.strip(' ')
                    if u'\xa0' in ele:
                        ele = ele.replace(u'\xa0', '')
                        qstring += ele
                        continue
                    temp.append(ele)
                temp.append(qstring)
            #添加每一行的url列表
            # print 'temp:\t', temp
            # print 'temp_url_list:\t', temp_url_list
            if temp_url_list:
                url_list.append(temp_url_list)
            left_col_list_1.append(head_constant)
            left_col_list_2.append(head_constant_1)
            num = temp.count('')
            for ele in range(num):
                temp.remove('')
            if temp:
                if temp[1] == temp[-2]:
                    temp.remove(temp[-2])
                #消除特殊字符带来的影响
                temp = self._remove_non_alphanumeric_characters(temp)

                if '20WW42' in self.data_url:
                    cell_data_list.append(temp[2:-2])
                else:
                    cell_data_list.append(temp)

        #去除多余的空格和换行
        for i in range(len(left_col_list_1)):
            left_col_list_1[i] = left_col_list_1[i].replace(' ', '')
            left_col_list_1[i] = left_col_list_1[i].replace('\n', '')
        for i in range(len(left_col_list_2)):
            left_col_list_2[i] = left_col_list_2[i].replace(' ', '')
            left_col_list_2[i] = left_col_list_2[i].replace('\n', '')
        #第40周特殊
        if len(header_list) > len(cell_data_list[0]):
            header_list = header_list[ -len(cell_data_list[0]): ]

        #保证cell_data数据有列对齐
        for k in range(len(cell_data_list)):
            if len(cell_data_list[k]) > len(header_list):
                cell_data_list[k] = cell_data_list[k][ -(len(header_list)): ]
        # print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        # print '\033[32mleft_col_list_1:\t\033[0m', left_col_list_1
        # print '\033[32mleft_col_list_2:\t\033[0m', left_col_list_2
        # print '\033[31murl_list:\t\033[0m', url_list, len(url_list)
        return Silver_Gold_BKC_string, header_length, self.date_string, url_list, header_list, cell_data_list, left_col_list_1, left_col_list_2
        #返回表头列表，行单元格信息列表，左边两列信息列表

    def get_ifwi_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)

        if not tr_list:
            return Silver_Gold_BKC_string, self.date_string, [], []
        for tr in tr_list:
            soup = BeautifulSoup(str(tr), 'html.parser')
            th_list = soup.find_all('th')
            if tr == tr_list[0]:
                #提取表头
                for th in th_list:
                    b = re.findall('>(.*?)<', str(th))
                    header_list.append(b[0])
            if not th_list:
                td_string_list = list(soup.strings)
                num = td_string_list.count('\n')
                for j in range(num):
                    td_string_list.remove('\n')
                #去掉首位字符串开头的空格,排除空格的影响 后续会排序
                td_string_list[0] = td_string_list[0].lstrip(' ')
                cell_data_list.append(td_string_list)
        # print '\033[31mheader_list:\t\033[0m', header_list
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list
        return Silver_Gold_BKC_string, self.date_string, header_list, cell_data_list

    def get_platform_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)

        if not tr_list:
            return Silver_Gold_BKC_string, self.date_string, [], [], []
        effective_url_list = []
        for tr in tr_list:
            temp = []
            url_list = []
            soup = BeautifulSoup(str(tr), 'html.parser')
            th_list = soup.find_all('th')
            td_list = soup.find_all('td')
            #获取header_list
            if th_list:
                th_strings = soup.strings
                for th in th_strings:
                    str_th = th.replace('\n', '')
                    if len(str_th) == 0:
                        continue
                    header_list.append(th)
            if td_list:
                # 获取url链接
                temp_url_list = soup.findAll(name='a', attrs={'href': re.compile(r'[https|http]:(.*?)')})
                if temp_url_list:
                    for ele in temp_url_list:
                        soup_href = BeautifulSoup(str(ele), 'html.parser')
                        string_url = soup_href.a['href']
                        url_list.append(string_url)
                    effective_url_list.append(url_list)
                else:
                    effective_url_list.append([])
                #需要循环处理，防止丢失无内容的项
                for td in td_list:
                    td_soup = BeautifulSoup(str(td), 'html.parser')
                    temp_list = list(td_soup.td.strings)
                    num = temp_list.count('\n')
                    for i in range(num):
                        temp_list.remove('\n')
                    if ' ' in temp_list:
                        temp_list.remove(' ')
                    if not temp_list:
                        temp.append('')
                        continue
                    for k in temp_list:
                        temp.append(k)
                cell_data_list.append(temp)
        header_list = header_list[1:]
        # print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        # print '\033[32meffective_url_list:\t\033[0m', effective_url_list, len(effective_url_list)
        return Silver_Gold_BKC_string, self.date_string, effective_url_list, header_list, cell_data_list

    def get_rework_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)

        if not tr_list:
            return Silver_Gold_BKC_string, [], self.date_string
        object_string_list = []
        for tr in tr_list:
            soup = BeautifulSoup(str(tr), 'html.parser')
            h4_list = soup.find_all('h4')
            for h4 in h4_list:
                h4_string = BeautifulSoup(str(h4), 'html.parser')
                temp = list(h4_string.strings)
                for strings in temp:
                    try:
                        if strings == '\n':
                            temp.remove(strings)
                        if strings.isspace():
                            temp.remove(strings)
                    except ValueError:
                        pass
                if h4 != h4_list[-1]:
                    if temp[0] == u'\xc2\xa0':
                        continue
                    object_string_list.append(' '.join(temp).strip(u'\xb7'))
                else:
                    for i in range(len(temp)):
                        if temp[i] == u'l' or temp[i] == u'\xa0' or temp[i] == u'\xb7' or temp[i] == u'\xc2\xa0':
                            continue
                        object_string_list.append(temp[i])
                if len(object_string_list) == 2:
                    object_string_list.append('')

            if len(object_string_list) < 10:
                # 得到tr节点下的所有文本
                text = list(soup.td.children)
                object_string_list = []
                for child in text:
                    soup_son = BeautifulSoup(str(child), 'html.parser')
                    child_tag_string_list = list(soup_son.strings)
                    if child_tag_string_list[0] == u'\n' or child_tag_string_list[0] == u'\xc2\xa0':
                        continue
                    temp_string = ''.join(child_tag_string_list).strip(' ')
                    if len(object_string_list) == 2:
                        object_string_list.append('')
                    object_string_list.append(temp_string)
        # 对提取的字符串列表进行清洗，统一组合格式：空格分隔  Mon Apr 10 14:00:27 2017
        object_string_list = self._remove_non_alphanumeric_characters(object_string_list)
        object_string_list[0] = re.sub('[()]', '', object_string_list[0]).strip()
        # print object_string_list, len(object_string_list)
        return Silver_Gold_BKC_string, object_string_list, self.date_string

    def get_existing_sighting_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)

        if not tr_list:
            return Silver_Gold_BKC_string, [], [], [], self.date_string
        #获取链接地址
        cell_last_data = []
        effective_url_list = []
        for tr in tr_list:
            soup = BeautifulSoup(str(tr), 'html.parser')
            if tr == tr_list[0]:
                header_list = list(soup.strings)[1::2]
                continue
            temp = []
            strings = []
            url_list = []

            td_list = soup.find_all('td')
            for td in td_list:
                # print td
                soup_td = BeautifulSoup(str(td), 'html.parser')
                # 获取url链接
                temp_url_list = soup_td.find_all('a')
                # if temp_url_list:
                #     print list(temp_url_list)
                if temp_url_list:
                    for ele in temp_url_list:
                        soup_href = BeautifulSoup(str(ele), 'html.parser')
                        # print list(soup_href.a.strings)
                        string_list = list(soup_href.a.strings)

                        if string_list:
                            strings.append(string_list[0])
                        #BKC第13周比较特殊，有个双层链接
                        try:
                            special_data = soup_href.findAll(name='a', attrs={'1358': re.compile(r'(.*)')})
                            if special_data:
                                url_list.append('[BKC][FPGA][s927')
                        except KeyError:
                            pass

                        string_url = soup_href.findAll(name='a', attrs={'href': re.compile(r'[https|http]:(.*?)')})
                        if string_url:
                            #排除无效的url
                            if re.search('.*id="></a>$', str(string_url[0])):
                                continue
                            #提取有效的url链接
                            for url in string_url:
                                url_soup = BeautifulSoup(str(url), 'html.parser')
                                url = url_soup.a['href']
                                url_list.append(url)

                td_son_list = list(soup_td.strings)
                for i in td_son_list:
                    if i == '\n':
                        td_son_list.remove(i)
                    if len(i) == 1 and i.isspace():
                        num = td_son_list.count(' ')
                        try:
                            for nu in range(num):
                                td_son_list.remove(i)
                        except:
                            pass
                if len(td_son_list) != 0:
                    if len(td_son_list) == 1:
                        temp.append(td_son_list[0])
                    else:
                        temp.append(td_son_list[-1])
                if len(td_son_list) == 0:
                    temp.append('')

            cell_last_data.append(strings)
            #去除多余的换行符和空格，防止插入excel表格时报 255 characters since it exceeds Excel's limit for URLS 长度超过限定长度
            for k in range(len(url_list)):
                url_list[k] = re.sub('[\s]', '', url_list[k])
            effective_url_list.append(url_list)
            #移除非字母数字字符
            temp = self._remove_non_alphanumeric_characters(temp)
            cell_data_list.append(temp)

        for k in range(len(cell_data_list)):
            cell_data_list[k].pop()
            if len(cell_last_data[k]) >= 2:
                for e in cell_last_data[k][1:]:
                    cell_data_list[k].append(e)
        # print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        # print '\033[32meffective_link_address_list:\t\033[0m', effective_url_list, len(effective_url_list)
        return Silver_Gold_BKC_string, effective_url_list, header_list, cell_data_list, self.date_string

    def get_new_sightings_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)

        if not tr_list:
            return Silver_Gold_BKC_string, self.date_string, [], [], []
        effective_url_list = []
        for tr in tr_list:
            url_list = []
            temp = []
            soup = BeautifulSoup(str(tr), 'html.parser')
            th_list = soup.find_all('th')
            td_list = soup.find_all('td')
            if th_list:
                header_list = list(soup.tr.strings)
                num = header_list.count('\n')
                for i in range(num):
                    header_list.remove('\n')
            if td_list:
                # 获取url链接
                temp_url_list = soup.findAll(name='a', attrs={'href': re.compile(r'[https|http]:(.*?)')})
                if temp_url_list:
                    # 提取有效的url链接
                    for url in temp_url_list:
                        url_soup = BeautifulSoup(str(url), 'html.parser')
                        url = url_soup.a['href']
                        # 排除无效的url 无效的url id后不跟数字
                        if not re.search('(\d+)$', str(url)):
                            continue
                        url_list.append(url)
                    effective_url_list.append(url_list)
                #需要循环处理，防止丢失无内容的项
                for td in td_list:
                    td_soup = BeautifulSoup(str(td), 'html.parser')
                    temp_list = list(td_soup.td.strings)
                    num = temp_list.count('\n')
                    for i in range(num):
                        temp_list.remove('\n')
                    if ' ' in temp_list:
                        temp_list.remove(' ')
                    if not temp_list:
                        temp.append('')
                        continue
                    temp.append(temp_list[0])
                temp = self._remove_non_alphanumeric_characters(temp)
                cell_data_list.append(temp)
        #第40周会出现很多换行符和多余的空格
        # deal_url_list = [ [ re.sub('[\n| ]', '', url) for url in url_first] for url_first in effective_url_list ]
        # if len(deal_url_list) <= len(effective_url_list):
        #     effective_url_list = copy.deepcopy(deal_url_list)
        # print '\033[31mheader_list:\t\033[0m', header_list
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list
        # print '\033[32meffective_url_list:\t\033[0m', effective_url_list
        return Silver_Gold_BKC_string, self.date_string, effective_url_list, header_list, cell_data_list

    def get_closed_sightings_data(self, data_type, bkc_flag=True):
        Silver_Gold_BKC_string, tr_list, header_list, cell_data_list = self.judge_silver_bkc_func(data_type, bkc_flag)

        if not tr_list:
            return Silver_Gold_BKC_string, self.date_string, [], [], []
        # print tr_list
        effective_url_list = []
        for tr in tr_list:
            url_list = []
            temp = []
            soup = BeautifulSoup(str(tr), 'html.parser')
            th_list = soup.find_all('th')
            td_list = soup.find_all('td')
            if th_list:
                header_list = list(soup.tr.strings)
                num = header_list.count('\n')
                for i in range(num):
                    header_list.remove('\n')
            if td_list:
                # 获取url链接
                temp_url_list = soup.findAll(name='a', attrs={'href': re.compile(r'[https|http]:(.*?)')})
                if temp_url_list:
                    for ele in temp_url_list:
                        soup_href = BeautifulSoup(str(ele), 'html.parser')
                        string_url = soup_href.a['href']
                        url_list.append(string_url)
                    effective_url_list.append(url_list)
                #需要循环处理，防止丢失无内容的项
                for td in td_list:
                    td_soup = BeautifulSoup(str(td), 'html.parser')
                    temp_list = list(td_soup.td.strings)
                    num = temp_list.count('\n')
                    for i in range(num):
                        temp_list.remove('\n')
                    if ' ' in temp_list:
                        temp_list.remove(' ')
                    if not temp_list:
                        temp.append('')
                        continue
                    temp.append(temp_list[0])
                temp = self._remove_non_alphanumeric_characters(temp)
                cell_data_list.append(temp)
        # print '\033[31mheader_list:\t\033[0m', header_list
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list
        # print '\033[32meffective_url_list:\t\033[0m', effective_url_list
        return Silver_Gold_BKC_string, self.date_string, effective_url_list, header_list, cell_data_list

    def get_caseresult_data(self, data_type, bkc_flag=True):
        key_num_string = self.save_file_name.split('_')[0]
        #默认是BKC caseresult
        bkc_case_result_url = 'https://dcg-oss.intel.com/test_report/report_case/' + key_num_string + '/BKC/' + self.grandfather_path
        if not bkc_flag:
            bkc_case_result_url = 'https://dcg-oss.intel.com/test_report/report_case/' + key_num_string + '/Silver/' + self.grandfather_path

        html = urllib2.urlopen(bkc_case_result_url).read()
        soup = BeautifulSoup(str(html), 'html.parser')
        #获取数据部分头部的标记字符串  Purley-FPGA WW12
        table_tip = soup.find_all('table')[0]
        tip_soup = BeautifulSoup(str(table_tip), 'html.parser')
        tip_string_list = list(tip_soup.strings)
        num = tip_string_list.count('\n')
        for i in range(num):
            tip_string_list.remove('\n')
        tip_string = tip_string_list[0].strip(' ')
        temp = tip_string.split(' ')
        tip_string = temp[0] + ' ' + temp[-1]
        #获取有效的待插入数据
        table = soup.find_all('table')[1]
        soup_tr = BeautifulSoup(str(table), 'html.parser')
        tr_list = soup_tr.find_all('tr')

        header_list = []
        cell_data_list = []
        if not tr_list:
            return self.date_string, [], [], []
        #获取表列名
        for tr in tr_list:
            temp = []
            soup = BeautifulSoup(str(tr), 'html.parser')
            th_list = soup.find_all('th')
            td_list = soup.find_all('td')
            if th_list:
                header_list = list(soup.tr.strings)
                num = header_list.count('\n')
                for i in range(num):
                    header_list.remove('\n')
            #获取表单元格数据
            if td_list:
                #需要循环处理，防止丢失无内容的项
                for td in td_list:
                    td_soup = BeautifulSoup(str(td), 'html.parser')
                    temp_list = list(td_soup.td.strings)
                    num = temp_list.count('\n')
                    for i in range(num):
                        temp_list.remove('\n')
                    if ' ' in temp_list:
                        temp_list.remove(' ')
                    if not temp_list:
                        temp.append('')
                        continue
                    temp.append(temp_list[0])
                cell_data_list.append(temp)
        # print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        # print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        # print '\033[32mtip_string:\t\033[0m', tip_string, len(tip_string)
        # for cell in cell_data_list:
        #     print cell
        return self.date_string, tip_string, header_list, cell_data_list


if __name__ == '__main__':
    import time
    start = time.time()

    key_url_list = []
    f = open(os.getcwd() + os.sep + 'report_html' + os.sep + 'url_info.txt')
    for line in f:
        if 'Purley-FPGA' in line and 'Silver' in line:
            key_url_list.append(line.strip('\n'))
    cache = DiskCache()
    # key_url_list = ['https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/Silver/2017%20WW16/5885_Silver.html',
    #                 'https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/BKC/2016%20WW43/4146_BKC.html',
    #                 'https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/BKC/2016%20WW42/4112_BKC.html',
    #                 'https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/BKC/2016%20WW40/4064_BKC.html']
    key_url_list = ['https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/Silver/2017%20WW17/5917_Silver.html']
    for url in key_url_list:
        verify = verify_validity_url(url)
        if not verify:
            continue
        print '\033[31m开始提取 %s 数据\033[0m' % url
    # url = 'https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/Silver/2017%20WW11/5691_Silver.html'
        obj = GetAnalysisData(url, get_info_not_save_flag=True, cache=cache)
        # obj.save_html()
        # obj.get_caseresult_data('Platform Integration Validation Result')
        # obj.get_platform_data('Platform Integration Validation Result', True)
        # obj.get_hw_data('HW Configuration')
        # obj.get_rework_data('HW Rework', True)
        # Silver_Gold_BKC_string, header_length, date_string, url_list, header_list, cell_data_list, left_col_list_1, left_col_list_2 = \
        # obj.get_sw_data('SW Configuration', True)
    # obj.get_hw_data('HW Configuration')
    # obj.get_ifwi_data('IFWI Configuration')
    #     obj.get_platform_data('Platform Integration Validation Result')
    # obj.get_rework_data('HW Rework', True)
    # obj.get_existing_sighting_data('Existing Sightings')
    # obj.get_new_sightings_data('New Sightings')
    # obj.get_closed_sightings_data('Closed Sightings')
        obj.get_caseresult_data('Platform Integration Validation Result')
    # f = unichr(213)
    # start = time.time()
    # unicode_set = set()
    # replace_list = []
    # for i in xrange(65536):
    #     unicode_set.add(unichr(i))
    #     if unichr(i) in [u'\u2018', u'\xb7', u'\xa0']:
    #         print i, unichr(i)
    #         replace_list.append(unichr(i))
    # print unicode_set
    # print replace_list
    print time.time() - start

#\xe2\u20ac\u2122t '
#\u2019 '右单引号
#\u2018 '左单引号
#OrderedDict([(u'\xe2', 226), (u'\u2018', 8216), (u'\u2019', 8217), (u'\u20ac', 8364), (u'\u2122', 8482)])
    # unicode_character_list = set()
    # import collections
    # temp_list = collections.OrderedDict()
    # for num in range(sys.maxunicode+1):
    #     unicode_character_list.add(unichr(num))
    #     if unichr(num) in [u'\u2018', u'\u2019', u'\u20ac', u'\u2122', u'\xe2']:
    #         print unichr(num)
    #         temp_list[unichr(num)] = num
    # print unicode_character_list
    # print temp_list

