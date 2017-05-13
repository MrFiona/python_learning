#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-04-25 14:55
# Author  : MrFiona
# File    : public_use_function.py
# Software: PyCharm Community Edition

import re
import os
import copy
import glob
import time
import shutil
import pstats
import urllib2
import cProfile
import HTMLParser
import win32com.client
from machine_config import MachineConfig
from setting_gloab_variable import DO_PROF, SRC_EXCEL_DIR, BACKUP_EXCEL_DIR, PURL_BAK_STRING, CONFIG_FILE_PATH


def get_week_num_config():
    # 读取配置文件周数配置
    conf = MachineConfig(CONFIG_FILE_PATH)
    week_num = conf.get_node_info('week_config', 'week_num')
    print week_num
    return week_num


def get_url_list_by_keyword(pre_keyword, back_keyword, key_url_list=None, reserve_url_num=50):
    if not key_url_list:
        key_url_list = []

    f = open(os.getcwd() + os.sep + 'report_html' + os.sep + 'url_info.txt')
    for line in f:
        if pre_keyword in line and back_keyword in line:
            key_url_list.append(line.strip('\n'))
    key_url_list = key_url_list[:reserve_url_num]
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


#验证url的有效性
def verify_validity_url(url):
    print '\033[36m验证\033[0m \033[36m[ \033[43m%s\033[0m \033[36m]开始\033[0m' %(url)
    try:
        response = urllib2.urlopen(url)
    except urllib2.URLError, e:
        if hasattr(e, 'reason'):  # stands for URLError
            print "can not reach a server,writing..."
            print "write url success!"
            return False
        elif hasattr(e, 'code'):  # stands for HTTPError
            print "find http error, writing..."
            print "write url success!"
            return False
        else:  # stands for unknown error
            print "unknown error, writing..."
            print "write url success!"
            return False
    else:
        # print "url is reachable!"
        # else 中不用再判断 response.code 是否等于200,若没有抛出异常，肯定返回200,直接关闭即可
            response.close()
            return True
    finally:
        print '\033[36m验证\033[0m \033[36m[ \033[43m%s\033[0m \033[36m]结束\033[0m' % (url)


def return_label_num_list(self):
    # 写入多个周的数据
    for i in range(24):
        # i为偶数
        if i & 1 == 0:
            if i < 2:
                x = ['D', 'Q']
            elif i >= 2:
                prefix = chr(ord('A') + (i - 2) / 2)
                x = [prefix + 'D', prefix + 'Q']
        elif i & 1 == 1:
            if i < 3:
                x = ['Q', 'AD']
            elif i > 3:
                prefix_font = chr(ord('A') + (i - 3) / 2)
                prefix_behind = chr(ord('A') + (i - 3) / 2 + 1)
                x = [prefix_font + 'Q', prefix_behind + 'D']
            elif i == 3:
                x = ['AQ', 'BD']
        self.alp_label_list.append(x)


def generate_column_alp_list(self):
    big_list = []
    alp_label_list = []
    effective_alp_list = []
    # 将首字母对转化为字母对列表,此处要用深拷贝，不然会导致两个同时变化
    first_alp = copy.deepcopy(alp_label_list)
    alp_pair_list = copy.deepcopy(alp_label_list)
    big_list.append(first_alp)
    # 循环执行10次，生成10个列对应的字母标记
    for nu in range(9):
        for label in range(len(alp_pair_list)):
            for i in range(len(alp_pair_list[label])):
                ele = alp_pair_list[label][i]
                if len(ele) == 1:
                    alp_pair_list[label][i] = chr(ord(alp_pair_list[label][i]) + 1)
                elif len(ele) == 2:
                    alp_pair_list[label][i] = alp_pair_list[label][i][0] + chr(ord(alp_pair_list[label][i][1]) + 1)
        # 这里需要深拷贝
        k = copy.deepcopy(alp_pair_list)
        big_list.append(k)

    # 每次新增空列表从后插值
    for i in range(24):
        effective_alp_list.append([])
        for alp in big_list:
            effective_alp_list[-1].append(alp[i])


def translate_test(self, num_total):
    # print 'num_total: ', num_total
    integer_num_alp = num_total - 65
    if integer_num_alp >= 0:
        # print 'integer_num_alp: ', integer_num_alp
        thousand_bit_num = integer_num_alp / (26 ** 4)
        remainder_thousand_bit_num = integer_num_alp % (26 ** 4)
        hundred_bit_num = remainder_thousand_bit_num / (26 ** 3)
        remainder_hundred_bit_num = remainder_thousand_bit_num % (26 ** 3)
        ten_bit_num = remainder_hundred_bit_num / (26 ** 2)
        remainder_ten_bit_num = remainder_hundred_bit_num % (26 ** 2)
        bit_num = remainder_ten_bit_num / 26
        remainder_bit_num = remainder_ten_bit_num % 26
        if thousand_bit_num > 26:
            raise UserWarning('只支持在456976列范围内!!!')

        object_alp = (chr(ord('A') + thousand_bit_num - 1) if thousand_bit_num else '') + (
        chr(ord('A') + hundred_bit_num - 1) if hundred_bit_num else '') + \
                     (chr(ord('A') + ten_bit_num - 1) if ten_bit_num else '') + (
                     chr(ord('A') + bit_num - 1) if bit_num else '') + \
                     (chr(ord('A') + remainder_bit_num))
        return object_alp


def canculate_head_num(self, multiple, num, add_num=0):
    url_length = len(self.FPGA_Silver_url_list)
    # return multiple*100 - multiple*(num + 1) + add_num
    return multiple*(100 - url_length + num) + add_num


def hidden_data_by_column(sheet_name, url_list, multiple):
    hidden_length = get_week_num_config() - len(url_list)
    # hidden_length = 50 - len(url_list)
    sheet_name.set_column(0, multiple * hidden_length - 1, options={'hidden': True})


# 性能分析装饰器定义
def do_cprofile(filename):
    """
    Decorator for function profiling.
    """
    def wrapper(func):
        def profiled_func(*args, **kwargs):
            # Flag for do profiling or not.
            if DO_PROF:
                profile = cProfile.Profile()
                profile.enable()
                result = func(*args, **kwargs)
                profile.disable()
                # Sort stat by internal time.
                sortby = "tottime"
                ps = pstats.Stats(profile).sort_stats(sortby)
                ps.dump_stats(filename)
            else:
                result = func(*args, **kwargs)
            return result
        return profiled_func
    return wrapper


#默认保留20个文件
def backup_excel_file(src_dir=SRC_EXCEL_DIR, backup_dir=BACKUP_EXCEL_DIR, reserve_file_max_num=20):
    backup_name = backup_dir + os.sep + 'backup_excel_' + time.strftime('%Y_%m_%d_%H_%M_%S', time.localtime(time.time()))
    #check excel file num default max=20
    orignal_file_list = glob.glob(src_dir + os.sep + '*.xlsx')

    #原始备份目录不存在则跳过备份
    if not os.path.exists(src_dir):
        return

    if os.path.exists(src_dir):
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)

        backup_file_list, file_num = os.listdir(backup_dir), len(os.listdir(backup_dir))
        if file_num >= reserve_file_max_num:
            backup_file_list.sort()
            #删除时间最久的目录以及文件
            for file in os.listdir(backup_dir + os.sep + backup_file_list[0]):
                os.remove(backup_dir + os.sep + backup_file_list[0] + os.sep + file)
            os.rmdir(backup_dir + os.sep + backup_file_list[0])

        if not os.path.exists(backup_name):
            os.makedirs(backup_name)

        for file in orignal_file_list:
            shutil.copy2(file, backup_name)


class FilterTag():
    def __init__(self):
        pass

    def filterHtmlTag(self, htmlStr):
        '''
        过滤html中的标签
        :param htmlStr:html字符串 或是网页源码
        '''
        self.htmlStr = htmlStr
        # print self.htmlStr
        # print len(self.htmlStr)
        # 先过滤CDATA
        re_cdata = re.compile('//<!--\[CDATA\[[^-->]*//\]\]>', re.I)  # 匹配CDATA
        re_script = re.compile('<\s*script[^>]*>[^<]*<\s*/\s*script\s*>', re.I)  # Script
        re_style = re.compile('<\s*style[^>]*>[^<]*<\s*/\s*style\s*>', re.I)  # style
        re_br = re.compile('<br\s*? ?="">')  # 处理换行
        re_h = re.compile('<!--?\w+[^-->]*>')  # HTML标签
        re_comment = re.compile('<!--[^>]*-->')  # HTML注释
        s = re_cdata.sub('', htmlStr)  # 去掉CDATA
        s = re_script.sub('', s)  # 去掉SCRIPT
        s = re_style.sub('', s)  # 去掉style
        s = re_br.sub('\n', s)  # 将br转换为换行
        blank_line = re.compile('\n+')  # 去掉多余的空行
        s = blank_line.sub('\n', s)
        s = re_h.sub('', s)  # 去掉HTML 标签
        s = re_comment.sub('', s)  # 去掉HTML注释
        # 去掉多余的空行
        blank_line = re.compile('\n+')
        s = blank_line.sub('\n', s)
        filterTag = FilterTag()
        s = filterTag.replaceCharEntity(s)  # 替换实体
        # print len(s)
        s = self.replaceCharEntity(s)
        # print len(s)
        # print  s
        return s


    def replaceCharEntity(self, htmlStr):
        '''
        替换html中常用的字符实体
        使用正常的字符替换html中特殊的字符实体
        可以添加新的字符实体到CHAR_ENTITIES 中
    CHAR_ENTITIES是一个字典前面是特殊字符实体  后面是其对应的正常字符
        :param htmlStr:
        '''
        self.htmlStr = htmlStr
        CHAR_ENTITIES = {'nbsp': ' ', '160': ' ',
                         'lt': '<', '60': '<',
                         'gt': '>', '62': '>',
                         'amp': '&', '38': '&',
                         'quot': '"', '34': '"', }
        re_charEntity = re.compile(r'&#?(?P<name>\w+);')
        sz = re_charEntity.search(htmlStr)
        while sz:
            entity = sz.group()  # entity全称，如>
            key = sz.group('name')  # 去除&;后的字符如（" "--->key = "nbsp"）    去除&;后entity,如>为gt
            try:
                htmlStr = re_charEntity.sub(CHAR_ENTITIES[key], htmlStr, 1)
                sz = re_charEntity.search(htmlStr)
            except KeyError:
                # 以空串代替
                htmlStr = re_charEntity.sub('', htmlStr, 1)
                sz = re_charEntity.search(htmlStr)
        # print htmlStr
        return htmlStr

    def replace(self, s, re_exp, repl_string):
        return re_exp.sub(repl_string)

    def strip_tags(self, htmlStr):
        '''
        使用HTMLParser进行html标签过滤
        :param htmlStr:
        '''

        self.htmlStr = htmlStr
        htmlStr = htmlStr.strip()
        htmlStr = htmlStr.strip("\n")
        result = []
        parser = HTMLParser.HTMLParser()
        parser.handle_data = result.append
        parser.feed(htmlStr)
        parser.close()
        return ''.join(result)

    def stripTagSimple(self, htmlStr):
        '''
        最简单的过滤html <>标签的方法    注意必须是<任意字符>  而不能单纯是<>
        :param htmlStr:
        '''
        print len(htmlStr)
        self.htmlStr = htmlStr
        #         dr =re.compile(r'<[^>]+>',re.S)
        dr = re.compile(r'<!--?\w+[^-->]*>', re.S)
        htmlStr = re.sub(dr, '', htmlStr)
        print len(htmlStr)
        return htmlStr


class easyExcel:
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):  # 打开文件或者新建文件（如果不存在的话）
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''


    def save(self, newfilename=None):  # 保存文件
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):  # 关闭文件
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):  # 获取单元格的数据
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):  # 设置单元格的数据
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):  # 获得一块区域的数据，返回为一个二维元组
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  # 插入图片
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self, before):  # 复制工作表
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None, shts(1))


def get_report_data(sheet_name, purl_bak_string='Purley-FPGA', cell_data_list=None):
    WEEK_NUM = get_week_num_config()
    Silver_url_list = get_url_list_by_keyword(PURL_BAK_STRING, 'Silver')
    if cell_data_list is None:
        cell_data_list = []

    win_book = easyExcel(os.getcwd() + os.sep + 'excel_dir' + os.sep + purl_bak_string + '_report_result.xlsx')

    if sheet_name == 'Save-Miss':
        for i in range(1, 8):
            temp_cell_list = []
            for j in range(1, 3):
                data = win_book.getCell(sheet=sheet_name, row=i, col=j)

                if i in (4, 5) and isinstance(data, float):
                    data = '%.2f%%' %(data*100)
                elif i in (6, 7) and isinstance(data, float):
                    data = '%.2f' %(data)
                elif i == 2 and isinstance(data, float):
                    data = int(data)
                elif data is None:
                    data = ''

                temp_cell_list.append(data)

            for k in range(WEEK_NUM + 3 - len(Silver_url_list), WEEK_NUM + 4):
            # for k in range(53 - len(Silver_url_list), 54):
                data = win_book.getCell(sheet=sheet_name, row=i, col=k)
                if i in (1, 2, 6, 7) and isinstance(data, float):
                    data = int(data)
                elif i in (4, 5) and isinstance(data, float):
                    data = '%.f%%' %(data*100)
                elif data is None:
                    data = ''

                temp_cell_list.append(data)

            cell_data_list.append(temp_cell_list)

    elif sheet_name in ('NewSi', 'ExistingSi'):
        fstop_flag = False
        for i in range(3, 200):
            temp_cell_list = []
            #NewSi和existingSi用上一周的
            for j in range((WEEK_NUM - len(Silver_url_list) + 1)*13 + 3, (WEEK_NUM - len(Silver_url_list) + 1)*13 + 1 + 13):
            # for j in range((50 - len(Silver_url_list) + 1)*13 + 3, (50 - len(Silver_url_list) + 1)*13 + 1 + 13):
                data = win_book.getCell(sheet=sheet_name, row=i, col=j)
                if j == ((WEEK_NUM - len(Silver_url_list) + 1)*13 + 1 + 1) and not data:
                # if j == ((50 - len(Silver_url_list) + 1)*13 + 1 + 1) and not data:
                    fstop_flag = True

                if data is None:
                    data = ''
                elif isinstance(data, float):
                    data = int(data)

                temp_cell_list.append(data)

            if fstop_flag:
                break

            #排除True和False
            # print temp_cell_list
            temp = [ cell for cell in temp_cell_list[2:] if isinstance(cell, (unicode, str)) and len(cell) != 0 ]
            if not temp:
                continue
            cell_data_list.append(temp_cell_list)

    else:
        fstop_flag = False
        m = (WEEK_NUM - len(Silver_url_list)) * 41 + 35 + 10 + 2
        # m = (50 - len(Silver_url_list)) * 41 + 35 + 10 + 2
        for i in range(7, 400):
            temp_cell_list = []
            for j in range(m, m + 4):
                data = win_book.getCell(sheet=sheet_name, row=i, col=j)
                if data is None:
                    data = ''
                if (j == m + 1) and data == '':
                    fstop_flag = True

                temp_cell_list.append(data)

            if fstop_flag:
                break
            # print temp_cell_list
            cell_data_list.append(temp_cell_list)

    win_book.save()
    win_book.close()
    # print 'sheet_name:\t', sheet_name
    # print cell_data_list
    return cell_data_list


#进一步处理html文件，增加表格样式
def deal_html_data():
    object_path = os.getcwd() + os.sep + 'html_result'
    read_file_list = glob.glob(object_path + os.sep + '*.html')

    for file in read_file_list:
        read_file = open(file, 'r')
        write_file = open(object_path + os.sep + 'temp.html', 'w')

        line = read_file.readline()
        write_file.write(line)
        while len(line) != 0:
            line = read_file.readline()
            if '</head>' == line.strip('\n'):
                if 'Save-Miss' in file:
                    write_file.write('<style type="text/css">\n\ttd{word-break:break-all;word-wrap:break-word;max-width:200px;font-family:Calibri;font-size:14px}\n</style>\n</head>\n')
                else:
                    write_file.write('<style type="text/css">\n\ttd{word-break:break-all;word-wrap:break-word;max-width:400px;font-family:Calibri;font-size:14px}\n</style>\n</head>\n')
                continue
            write_file.write(line)
        read_file.close()
        write_file.close()
        os.remove(file)
        shutil.copy(object_path + os.sep + 'temp.html', file)
        os.remove(object_path + os.sep + 'temp.html')

if __name__ == '__main__':
    start = time.time()
    # backup_excel_file()
    # cell_data_list = get_report_data('CaseResult', purl_bak_string='Purley-FPGA')
    # cell_data_list = get_report_data('ExistingSi', purl_bak_string='Purley-FPGA')
    # cell_data_list = get_report_data('Save-Miss', purl_bak_string='Purley-FPGA')
    get_week_num_config()
    print time.time() - start