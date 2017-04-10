#!/usr/bin/env python
# -*- coding: utf-8 -*-

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
from bs4 import BeautifulSoup

# reload(sys)
# sys.setdefaultencoding('utf-8')

# socket.setdefaulttimeout(5)

#https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/Silver/2017%20WW10/5658_Silver.html
class GetAnalysisData():
    def __init__(self, data_url, get_info_not_save_flag=False):
        self.data_url = data_url
        self.get_info_not_save_flag = get_info_not_save_flag
        try:
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
        self.save_html()

    #获取url详细信息
    def get_info_detail(self):
        return self.data_url

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

    #提取通用的部分
    def __common_regex(self, data_type, bkc_flag=False, header_list=None, cell_data_list=None):
        if not header_list:
            header_list = []
        if not cell_data_list:
            cell_data_list = []
        # print data_type, bkc_flag
        if not os.path.exists(self.save_file_path):
            raise UserWarning('文件不存在!!!')
        #BKC标记开启，自动更换为文件BKC地址
        if bkc_flag:
            # print self.save_file_path
            self.save_file_path = self.save_file_path.replace('Silver', 'BKC')
            # print self.save_file_path
            # if os.path.exists(self.save_file_path):
            search_file_list = os.listdir(self.save_file_path)
            # print search_file_list
            bkc_file_name = ''
            for file in search_file_list:
                match = re.match(r'\d+', file)
                if match:
                    bkc_file_name = file
            if len(bkc_file_name) == 0:
                print 'BKC文件不存在'
                return []
            self.save_file_name = bkc_file_name

        # 读取文本中的html代码
        with codecs.open(self.save_file_path + os.sep + self.save_file_name, 'r', encoding='utf-8') as f:
            data = f.readlines()
        data = ''.join(data)
        # 提取HW Configuration部分的代码
        regex = re.compile(r'<span class="sh2">&nbsp; %s </span>(.*?)<div class="panel-heading">' % data_type,
                           re.S | re.M)
        header = re.findall(regex, data)
        string_data = ''.join(header)
        # 提取所有的tr部分
        soup_tr = BeautifulSoup(string_data, 'html.parser')
        tr_list = soup_tr.find_all(re.compile('tr'))

        #增加保险措施，通过关键字提取文本html实现
        if not tr_list:
            print '提取的tr_list为空!开始通过文本提取匹配内容......'
            file_path = self.save_file_path + os.sep + self.save_file_name
            print file_path
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
            soup = BeautifulSoup(str(lines), 'html.parser')
            tr_list = soup.find_all('tr')
            print '文本提取匹配内容结束......'

        return tr_list, header_list, cell_data_list

    #获取hw有效的列名列表
    def __get_hw_effective_header_list(self, effective_header_list):
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
    def __insert_numers_to_cell_data_list(self, cell_data_list):
        try:
            effective_header_list, effective_num_list = self.__get_hw_effective_header_list(cell_data_list[0])
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

    def get_hw_data(self, data_type):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type)
        # print tr_list, len(tr_list)
        if not tr_list:
            return None, None, None, None
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
                        effective_header_list, effective_num_list = self.__get_hw_effective_header_list(effective_header_list)
            if temp:
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
                effective_header_list, effective_num_list = self.__get_hw_effective_header_list(effective_header_list)
                # if not header_list:
                #     th_soup = BeautifulSoup(str(tr_list[0]), 'html.parser')
                #     th_list = th_soup.find_all('th')
                #     header_list.append('')

        # print 'effective_header_list:\t', effective_header_list
        # cell_data_list.insert(0, effective_header_list)
        # print cell_data_list
        cell_data_list = self.__insert_numers_to_cell_data_list(cell_data_list)
        # print cell_data_list
        if cell_data_list:
            cell_data_list[0] = effective_header_list
        # print cell_data_list, len(cell_data_list)
        header_cell_data_info = collections.OrderedDict()
        [header_cell_data_info.update({header_list[key]: cell_data_list[key]}) for key in range(len(header_list))]
        header_list = header_list[ -(len(cell_data_list) -1): ]
        if header_list[0].startswith('Ingredient'):
            header_list = header_list[1:]
        # 行坐标列表
        # print effective_header_list, len(effective_header_list)
        # 列坐标列表
        # print header_list
        return self.date_string, effective_header_list, header_list, cell_data_list

    def get_sw_data(self, data_type):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type)
        soup_header = BeautifulSoup(str(tr_list[0]), 'html.parser')
        header_list = list(soup_header.strings)
        num = header_list.count('\n')
        for j in range(num):
            header_list.remove('\n')
        for i in range(len(header_list)):
            header_list[i] = header_list[i].replace('\n', '')
        #排除部分周会有更多的列
        header_list = header_list[:4]
        difference_length = len(header_list) - len(header_list[:4])
        header_length = len(header_list)
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
            for td in actual_td_list:
                soup = BeautifulSoup(str(td), 'html.parser')
                string_list = list(soup.strings)
                #正则取匹配url链接
                obj_list = re.findall('<a href="(.*?)">', str(td), re.M|re.S)
                if obj_list:
                    #逐个添加
                    for i in obj_list:
                        temp_url_list.append(i)
                for i in string_list:
                    if i == '\n':
                        string_list.remove('\n')
                string_list[0] = string_list[0].replace('\n', '')
                #可能会含有多个的url链接，需要添加整个列表string_list，防止丢失数据
                for ele in string_list:
                    ele = ele.strip(' ')
                    if '\xa0' in ele.encode('utf-8'):
                        ele = ele.encode('utf-8').replace('\xa0', '')
                        qstring += ele.decode(chardet.detect(ele)['encoding'])
                        continue
                    temp.append(ele)
                temp.append(qstring)
            #添加每一行的url列表
            url_list.append(temp_url_list)
            left_col_list_1.append(head_constant)
            left_col_list_2.append(head_constant_1)
            num = temp.count('')
            for ele in range(num):
                temp.remove('')
            if temp[1] == temp[-2]:
                temp.remove(temp[-2])
            cell_data_list.append(temp)
        #去除多余的空格和换行
        for i in range(len(left_col_list_1)):
            left_col_list_1[i] = left_col_list_1[i].replace(' ', '')
            left_col_list_1[i] = left_col_list_1[i].replace('\n', '')
        for i in range(len(left_col_list_2)):
            left_col_list_2[i] = left_col_list_2[i].replace(' ', '')
            left_col_list_2[i] = left_col_list_2[i].replace('\n', '')
        # print cell_data_list
        # print header_list
        # print left_col_list_1
        # print left_col_list_2
        return self.date_string, url_list, header_list, cell_data_list, left_col_list_1, left_col_list_2
        #返回表头列表，行单元格信息列表，左边两列信息列表

    def get_ifwi_data(self, data_type):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type)
        for tr in tr_list:
            soup = BeautifulSoup(str(tr), 'html.parser')
            th_list = soup.find_all('th')
            #提取表头
            if th_list:
                for th in th_list:
                    b = re.split('[><]', th.encode('utf-8'))
                    header_list.append(b[2])
            if not th_list:
                td_list = list(soup.strings)
                num = td_list.count('\n')
                for j in range(num):
                    td_list.remove('\n')
                cell_data_list.append(td_list)
        return self.date_string, header_list, cell_data_list

    def get_platform_data(self, data_type, bkc_flag=True):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type, bkc_flag=True)
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
                    # print temp_list
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
                # print temp
                cell_data_list.append(temp)
        header_list = header_list[1:]
        print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        print '\033[32meffective_url_list:\t\033[0m', effective_url_list, len(effective_url_list)
        return self.date_string, effective_url_list, header_list, cell_data_list

    def get_rework_data(self, data_type):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type)
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
                    object_string_list.append(''.join(temp).strip(u'\xb7'))
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
                    if u'\xb7' in child_tag_string_list:
                        child_tag_string_list.remove(u'\xb7')
                    if u'\xa0' in child_tag_string_list:
                        child_tag_string_list.remove(u'\xa0')
                    for i in range(len(child_tag_string_list)):
                        child_tag_string_list[i] = re.sub('[\xa0|\n]', '', child_tag_string_list[i])
                    temp_string = ''.join(child_tag_string_list).strip(' ')
                    if len(object_string_list) == 2:
                        object_string_list.append('')
                    object_string_list.append(temp_string)

        # print object_string_list, len(object_string_list)
        return object_string_list, self.date_string

    def get_existing_sighting_data(self, data_type, bkc_flag=True):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type, bkc_flag)
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
            effective_url_list.append(url_list)
            cell_data_list.append(temp)
            # print cell_last_data
            # print url_list, len(url_list)
        for k in range(len(cell_data_list)):
            cell_data_list[k].pop()
            if len(cell_last_data[k]) >= 2:
                for e in cell_last_data[k][1:]:
                    cell_data_list[k].append(e)
        print '\033[31mheader_list:\t\033[0m', header_list, len(header_list)
        print '\033[36mcell_data_list:\t\033[0m', cell_data_list, len(cell_data_list)
        print '\033[32meffective_link_address_list:\t\033[0m', effective_url_list, len(effective_url_list)
        return effective_url_list, header_list, cell_data_list, self.date_string

    def get_new_sightings_data(self, data_type, bkc_flag=True):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type, bkc_flag)
        # print tr_list
        effective_url_list = []
        for tr in tr_list:
            # print tr
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
                    # print temp_list
                    # print temp_list[0]
                    temp.append(temp_list[0])
                cell_data_list.append(temp)
        print '\033[31mheader_list:\t\033[0m', header_list
        print '\033[36mcell_data_list:\t\033[0m', cell_data_list
        print '\033[32meffective_url_list:\t\033[0m', effective_url_list
        return self.date_string, effective_url_list, header_list, cell_data_list

    def get_closed_sightings_data(self, data_type, bkc_flag=True):
        tr_list, header_list, cell_data_list = self.__common_regex(data_type, bkc_flag)
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
                cell_data_list.append(temp)
        # print header_list
        # print cell_data_list
        # print effective_url_list
        return self.date_string, effective_url_list, header_list, cell_data_list

    def get_caseresult_data(self, data_type, bkc_flag=True):
        #BKC   http://dcg-oss.intel.com/test_report/report_case/5736/BKC/Purley_FPGA/
        key_num_string = self.save_file_name.split('_')[0]
        bkc_case_result_url = 'http://dcg-oss.intel.com/test_report/report_case/' + key_num_string + '/BKC/' + self.grandfather_path
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
        return tip_string, header_list, cell_data_list

if __name__ == '__main__':
    url = 'https://dcg-oss.intel.com/ossreport/auto/Purley-FPGA/Silver/2017%20WW13/5774_Silver.html'
    obj = GetAnalysisData(url)
    obj.save_html()
    # obj.get_sw_data('SW Configuration')
    # obj.get_hw_data('HW Configuration')
    # obj.get_ifwi_data('IFWI Configuration')
    # obj.get_platform_data('Platform Integration Validation Result')
    # obj.get_rework_data('HW Rework')
    obj.get_existing_sighting_data('Existing Sightings')
    # obj.get_new_sightings_data('New Sightings')
    # obj.get_closed_sightings_data('Closed Sightings')
    # obj.get_caseresult_data('Platform Integration Validation Result')



