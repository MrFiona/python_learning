#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Time    : 2017-03-21 09:40
# Author  : MrFiona
# File    : insert_excel.py
# Software: PyCharm Community Edition

import os
import re
import sys
import copy
import time
import xlrd
import Queue
import extract_data
import openpyxl
import threading
import xlsxwriter
reload(sys)
sys.setdefaultencoding('utf-8')
from cache_mechanism import DiskCache

class InsertDataIntoExcel(object):
    def __init__(self, verificate_flag=False, cache=None):
        self.cache = cache
        self.alp_label_list = []
        self.effective_alp_list = []
        #验证标志,默认不开启,即预测
        self.verificate_flag = verificate_flag
        #url列表初始化
        self.return_label_num_list()
        self.generate_column_alp_list()
        self.FPGA_Silver_url_list = self.get_url_list_by_keyword('Purley-FPGA', 'Silver')
        self.Bakerville_url_list = self.get_url_list_by_keyword('Bakerville', 'Silver')
        # print self.FPGA_Silver_url_list
        print self.Bakerville_url_list

        # 创建一个新的excel工作簿  文件对象
        self.workbook = xlsxwriter.Workbook('report_result_3.xlsx', options={'strings_to_urls': False})
        #获取原始公式数据
        self.rb = openpyxl.load_workbook('C:\Users\pengzh5x\Desktop\ITF_Skylake_FPGA_BKC_TestCase_WW13_v1.5.8.xlsx', data_only=False)

        # 在excel工作簿中增加工作表单
        self.worksheet_change = self.workbook.add_worksheet('Change History')
        self.worksheet_newsi = self.workbook.add_worksheet('NewSi')
        self.worksheet_existing = self.workbook.add_worksheet('ExistingSi')
        self.worksheet_closesi = self.workbook.add_worksheet('ClosedSi')
        self.worksheet_rework = self.workbook.add_worksheet('Rework')
        self.worksheet_hw = self.workbook.add_worksheet('HW')
        self.worksheet_sw_orignal = self.workbook.add_worksheet('SW_Original')
        self.worksheet_sw = self.workbook.add_worksheet('SW')
        self.worksheet_ifwi_orignal = self.workbook.add_worksheet('IFWI_Original')
        self.worksheet_ifwi = self.workbook.add_worksheet('IFWI')
        self.worksheet_platform = self.workbook.add_worksheet('ValidationResult')
        self.worksheet_mapping = self.workbook.add_worksheet('Mapping')
        self.worksheet_caseresult = self.workbook.add_worksheet('CaseResult')
        self.worksheet_save_miss = self.workbook.add_worksheet('Save-Miss')
        self.worksheet_test_time = self.workbook.add_worksheet('TestTime')

        #读取样板excel
        self.read_workbook = xlrd.open_workbook('copy.xlsx')

        # 为excel对象设定显示格式
        self.bold = self.workbook.add_format({'bold': 1, 'align': 'center', 'font_size': 12})
        self.bold_merge = self.workbook.add_format({'bold': 2, 'align': 'justify', 'valign': 'vcenter', 'bg_color': '#C6EFCE'})
        # 为当前workbook添加一个样式名为titleformat
        self.titleformat = self.workbook.add_format()
        self.titleformat_header = self.workbook.add_format()
        self.titleformat_header.set_bold()  # 设置粗体字
        self.titleformat.set_font_size(12)  # 设置字体大小为10
        self.titleformat_header.set_font_size(12)  # 设置字体大小为10
        self.titleformat.set_font_name('Microsoft yahei')  # 设置字体样式为雅黑
        self.titleformat.set_align('center')  # 设置水平居中对齐
        self.titleformat_header.set_align('center')  # 设置水平居中对齐
        self.titleformat.set_align('vcenter')  # 设置垂直居中对齐
        self.titleformat_header.set_align('vcenter')  # 设置垂直居中对齐
        self.format1 = self.workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        self.mapping_format = self.workbook.add_format({'bg_color': '#EAC100'})
        self.mapping_blue_format = self.workbook.add_format({'bg_color': '#ECF5FF'})
        self.url_format = self.workbook.add_format({'font_color': 'blue', 'underline': 1, 'align':'center', 'valign':'vcenter'})

        self.titleformat_merage = self.workbook.add_format()
        self.titleformat_merage.set_font_size(12)  # 设置字体大小为10
        self.titleformat_merage.set_align('center')  # 设置水平居中对齐
        self.titleformat_merage.set_align('vcenter')  # 设置垂直居中对齐
        # Add a format for the header cells.
        self.header_format = self.workbook.add_format({
            'border': 1,
            'bg_color': '#C6EFCE',
            'bold': True,
            # 'text_wrap': True,
            # 'valign': 'vcenter',
            # 'indent': 1,
        })

    def get_url_list(self):
        return self.FPGA_Silver_url_list

    def get_url_list_by_keyword(self, pre_keyword, back_keyword, key_url_list=None):
        if not key_url_list:
            key_url_list = []
            
        f = open(os.getcwd() + os.sep + 'report_html' + os.sep + 'url_info.txt')
        for line in f:
            if pre_keyword in line and back_keyword in line:
                key_url_list.append(line.strip('\n'))
        # print key_url_list, len(key_url_list)
        return key_url_list

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
        # 将首字母对转化为字母对列表,此处要用深拷贝，不然会导致两个同时变化
        first_alp = copy.deepcopy(self.alp_label_list)
        alp_pair_list = copy.deepcopy(self.alp_label_list)
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
            self.effective_alp_list.append([])
            for alp in big_list:
                self.effective_alp_list[-1].append(alp[i])
        # for j in self.effective_alp_list:
        #     print 'ggg', j

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
        return multiple*num + add_num

    def get_formula_data(self, data_type, wb_sheet_name):
        rb_sheet = self.rb.get_sheet_by_name(data_type)
        row, col = 1, 1
        for rs in rb_sheet.rows:
            col = 1
            if data_type == 'Mapping':
                if row == 5:
                    row += 1
                    continue
            for cell in rs:
                data = rb_sheet.cell(row=row, column=col).value
                if isinstance(data, (str, long, unicode, bool, float)):
                    if data_type == 'Mapping':
                        if isinstance(data, unicode) and (data.startswith('WW') or data.startswith('EXX')):
                            col += 1
                            continue
                    wb_sheet_name.write(row - 1, col - 1, data)
                col += 1
            row += 1
        #清除最新一周数据，防止拷贝多余的数据
        #NewSi、ExistingSi前四列是公式
        if data_type == 'NewSi' or data_type == 'ExistingSi':
            for row in range(2, 50):
                for i in range(len(self.FPGA_Silver_url_list)):
                    for col in range(self.canculate_head_num(13, i, 4), self.canculate_head_num(13, i, 13)):
                        wb_sheet_name.write(row, col, ' ')
                    wb_sheet_name.write(1, self.canculate_head_num(13, i), ' ')
        #ClosedSi前四列是公式
        elif data_type == 'ClosedSi':
            for row in range(2, 50):
                for i in range(len(self.FPGA_Silver_url_list)):
                    for col in range(self.canculate_head_num(13, i, 4), self.canculate_head_num(13, i, 14)):
                        wb_sheet_name.write(row, col, ' ')
        #Rework第二列是公式
        elif data_type == 'Rework':
            for row in range(1, 50):
                for i in range(len(self.FPGA_Silver_url_list)):
                    for col in range(self.canculate_head_num(3, i), self.canculate_head_num(3, i, 3)):
                        if col == self.canculate_head_num(3, i, 1):
                            continue
                        wb_sheet_name.write(row, col, ' ')
        # HW第二行和第3列是公式
        elif data_type == 'HW':
            for row in range(1, 50):
                if row == 1:
                    continue
                for i in range(len(self.FPGA_Silver_url_list)):
                    for col in range(self.canculate_head_num(13, i), self.canculate_head_num(13, i, 14)):
                        if col == self.canculate_head_num(13, i, 2):
                            continue
                        wb_sheet_name.write(row, col, ' ')
        # SW_Original第1、2列是公式
        elif data_type == 'SW_Original':
            for row in range(1, 50):
                for i in range(len(self.FPGA_Silver_url_list)):
                    for col in range(self.canculate_head_num(9, i, 2), self.canculate_head_num(9, i, 10)):
                        wb_sheet_name.write(row, col, ' ')
        # SW第1、2列是公式
        elif data_type == 'SW':
            for row in range(0, 50):
                if row == 14:
                    continue
                for i in range(len(self.FPGA_Silver_url_list)):
                    for col in range(self.canculate_head_num(9, i, 2), self.canculate_head_num(9, i, 10)):
                        wb_sheet_name.write(row, col, ' ')
        # IFWI_Original和IFWI第1、2列是公式
        elif data_type == 'IFWI_Original' or data_type == 'IFWI':
            for row in range(3, 50):
                for i in range(len(self.FPGA_Silver_url_list)):
                    for col in range(self.canculate_head_num(6, i, 2), self.canculate_head_num(6, i, 5)):
                        wb_sheet_name.write(row, col, ' ')

    def insert_Newsi_data(self, k):
        print '开始获取 New Sightings 数据........\n'
        # 获取公式并插入指定位置
        self.get_formula_data(u'NewSi', self.worksheet_newsi)
        header_fix_list = [u'ID', u'Title', u'Priority', u'Severity', u'Owner', u'Status', u'Subsystem', u'Promoted ID']
        for j in range(len(self.FPGA_Silver_url_list)):
            print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
            obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
            if j == 0:
                date_string, effective_url_list, header_list, cell_data_list = obj.get_new_sightings_data('New Sightings', self.verificate_flag)
            else:
                date_string, effective_url_list, header_list, cell_data_list = obj.get_new_sightings_data('New Sightings', True)

            self.worksheet_newsi.conditional_format(1, self.canculate_head_num(13, j, 1), 12, self.canculate_head_num(13, j, 3),
                                                         {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
            insert_data_list_1 = ['NOT Reported in next release?', 'NOT Reported in previous release?',
                                'Purley-FPGA No corresponding test case in test result?']
            insert_data_stirng = 'IF(ISBLANK(E3),"",IF(D3,"",IF(ISERROR(MATCH(E3,CaseResult!Q:Q,0)),"",INDEX(CaseResult!N:N,MATCH(E3,CaseResult!Q:Q,0)))))'
            self.worksheet_newsi.write(0, self.canculate_head_num(13, j), insert_data_stirng)
            self.worksheet_newsi.write(0, self.canculate_head_num(13, j, 2), 'BKC', self.format1)
            self.worksheet_newsi.write(1, self.canculate_head_num(13, j), date_string, self.format1)
            self.worksheet_newsi.write(0, self.canculate_head_num(13, j, 3), 'ISERROR(MATCH(E3,CaseResult!$P:$P,0)),')
            self.worksheet_newsi.write_row(0, self.canculate_head_num(13, j, 4), header_list, self.header_format)
            #数据为空时插入默认的表列名
            if not header_list:
                self.worksheet_newsi.write_row(0, self.canculate_head_num(13, j, 4), header_fix_list, self.header_format)
            self.worksheet_newsi.write_row(1, self.canculate_head_num(13, j, 1), insert_data_list_1)
            self.worksheet_newsi.write(0, self.canculate_head_num(13, j, 12), 'Comments')

            if not effective_url_list and not header_list and not cell_data_list:
                continue

            num_url_list = []
            for ele in effective_url_list:
                num_url_list.append(len(ele))

            for i in range(len(cell_data_list)):
                #插入非url部分数据
                self.worksheet_newsi.write_row(2 + i, self.canculate_head_num(13, j, 5), cell_data_list[i][1:-1])
                #插入url数据部分
                self.worksheet_newsi.write_url(2 + i, self.canculate_head_num(13, j, 4), effective_url_list[i][0], self.url_format, cell_data_list[i][0])
                if num_url_list[i] > 1:
                    self.worksheet_newsi.write_url(2 + i, self.canculate_head_num(13, j, 11), effective_url_list[i][1], self.url_format, cell_data_list[i][-1])
                elif num_url_list[i] == 1:
                    self.worksheet_newsi.write(2 + i, self.canculate_head_num(13, j, 11), cell_data_list[i][-1])

    def insert_ExistingSi_data(self, k):
        prompt_statement_list = ['NOT Reported in next release?', 'NOT Reported in previous release?',
                                 'Purley-FPGA No corresponding test case in test result?']
        # Add the standard url link format.
        fail_url_list = []
        print '开始获取  Existing Sightings 数据........\n'
        # 获取公式并插入指定位置
        self.get_formula_data('ExistingSi', self.worksheet_existing)
        for j in range(len(self.FPGA_Silver_url_list)):
            print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
            obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
            if j == 0:
                url_list, header_list, cell_data_list, date_string = obj.get_existing_sighting_data('Existing Sightings', self.verificate_flag)
            else:
                url_list, header_list, cell_data_list, date_string = obj.get_existing_sighting_data('Existing Sightings', True)

            #增加一列comments
            header_list.append('comments')
            try:
                #标记True为红色
                self.worksheet_existing.conditional_format(2, self.canculate_head_num(13, j, 1), 20, self.canculate_head_num(13, j, 3),
                                                     {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
                red = self.workbook.add_format({'color': 'red'})
                self.worksheet_existing.write_rich_string(1, self.canculate_head_num(13, j), red, date_string, self.titleformat)
                self.worksheet_existing.write(0, self.canculate_head_num(13, j, 2), 'BKC', self.header_format)
                self.worksheet_newsi.write(1, self.canculate_head_num(13, j), date_string, self.format1)
                # 写入工作表的第二行,提示语
                self.worksheet_existing.write_row(1, self.canculate_head_num(13, j, 1), prompt_statement_list)
                if not url_list and not header_list and not cell_data_list:
                    continue
                #写入表头行
                self.worksheet_existing.write_row(0, self.canculate_head_num(13, j, 4), header_list, self.titleformat_header)

                # 最后一列可能会出现多值多行的情况，计算每行数据占有的行数
                line_num_list = []
                for ele in cell_data_list:
                    num = len(ele[7:])
                    line_num_list.append(num)
                # print line_num_list

                nu = 2
                # 插入数据第一列到第三列
                for line in range(len(cell_data_list)):
                    if line_num_list[line] <= 1:
                        self.worksheet_existing.write_row(nu, self.canculate_head_num(13, j, 5), cell_data_list[line][1:7])
                        if line_num_list[line] == 0:
                            self.worksheet_existing.write_url(nu, self.canculate_head_num(13, j, 4), url_list[line][0], self.url_format, str(cell_data_list[line][0]))
                            self.worksheet_existing.write(nu, self.canculate_head_num(13, j, 11), ' ')
                        elif line_num_list[line] == 1:
                            self.worksheet_existing.write_url(nu, self.canculate_head_num(13, j, 4), url_list[line][0], self.url_format, str(cell_data_list[line][0]))
                            self.worksheet_existing.write_url(nu, self.canculate_head_num(13, j, 11), url_list[line][1], self.url_format, str(cell_data_list[line][7]))
                        nu += 1

                    elif line_num_list[line] >= 2:
                        length_merge = line_num_list[line]
                        self.worksheet_existing.write_url(nu, self.canculate_head_num(13, j, 4), url_list[line][0], self.url_format, str(cell_data_list[line][0]))
                        self.worksheet_existing.write_row(nu, self.canculate_head_num(13, j, 5), cell_data_list[line][1:7],self.titleformat_merage)
                        for i in range(4, 11):
                            self.worksheet_existing.merge_range(nu, self.canculate_head_num(13, j, i),nu + line_num_list[line] - 1, self.canculate_head_num(13, j, i), '')
                        for m in range(length_merge):
                            self.worksheet_existing.write_url(nu + m, self.canculate_head_num(13, j, 11), url_list[line][m+1], self.url_format,str(cell_data_list[line][7 + m]))
                        nu += line_num_list[line]
            except:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取  Existing Sightings 数据........\n'

    def insert_ClosedSi_data(self, k):
        print '开始获取 New Sightings 数据........\n'
        # 获取公式并插入指定位置
        self.get_formula_data('ClosedSi', self.worksheet_closesi)
        for j in range(len(self.FPGA_Silver_url_list)):
            print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
            obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
            if j == 0:
                date_string, effective_url_list, header_list, cell_data_list = obj.get_closed_sightings_data('Closed Sightings', self.verificate_flag)
            else:
                date_string, effective_url_list, header_list, cell_data_list = obj.get_closed_sightings_data('Closed Sightings', True)

            self.worksheet_closesi.conditional_format(1, self.canculate_head_num(13, j, 1), 12, self.canculate_head_num(13, j, 3),
                                                         {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
            insert_data_list_1 = ['NOT Reported in next release?', 'NOT Reported in previous release?',
                                'Purley-FPGA No corresponding test case in test result?']
            insert_data_stirng = 'IF(ISBLANK(E3),"",IF(D3,"",IF(ISERROR(MATCH(E3,CaseResult!Q:Q,0)),"",INDEX(CaseResult!N:N,MATCH(E3,CaseResult!Q:Q,0)))))'
            self.worksheet_closesi.write(0, self.canculate_head_num(13, j), insert_data_stirng)
            self.worksheet_closesi.write(0, self.canculate_head_num(13, j, 2), 'BKC', self.format1)
            self.worksheet_closesi.write(0, self.canculate_head_num(13, j, 3), 'ISERROR(MATCH(E3,CaseResult!$P:$P,0)),')
            self.worksheet_closesi.write_row(1, self.canculate_head_num(13, j, 1), insert_data_list_1)
            # Set up some formats to use.
            red = self.workbook.add_format({'color': 'red'})
            self.worksheet_closesi.write_rich_string(1, self.canculate_head_num(13, j), red, date_string, self.titleformat)
            if not effective_url_list and not header_list and not cell_data_list:
                continue
            self.worksheet_closesi.write_row(0, self.canculate_head_num(13, j, 4), header_list, self.header_format)

            num_url_list = []
            for ele in effective_url_list:
                num_url_list.append(len(ele))

            for i in range(len(cell_data_list)):
                #插入非url部分数据
                self.worksheet_closesi.write_row(2 + i, self.canculate_head_num(13, j, 5), cell_data_list[i][1:-1])
                #插入url数据部分
                self.worksheet_closesi.write_url(2 + i, self.canculate_head_num(13, j, 4), effective_url_list[i][0], self.url_format, cell_data_list[i][0])
                if num_url_list[i] > 1:
                    self.worksheet_closesi.write_url(2 + i, self.canculate_head_num(13, j, 11), effective_url_list[i][1], self.url_format, cell_data_list[i][-1])
                elif num_url_list[i] == 1:
                    self.worksheet_closesi.write(2 + i, self.canculate_head_num(13, j, 11), cell_data_list[i][-1])

    def insert_Rework_data(self, k):
        fail_url_list = []
        print '开始获取 HW Rework 数据........\n'
        # 获取公式并插入指定位置
        self.get_formula_data('Rework', self.worksheet_rework)
        for j in range(len(self.FPGA_Silver_url_list)):
            try:
                print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
                obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
                #最新的一周在验证标志未开启的情况下取Silver数据,否则取BKC数据
                if j == 0:
                    object_string_list, date_string = obj.get_rework_data('HW Rework', self.verificate_flag)
                else:
                    object_string_list, date_string = obj.get_rework_data('HW Rework', True)

                # 标记True为红色
                self.worksheet_rework.conditional_format(1, self.canculate_head_num(3, j, 1), 100, self.canculate_head_num(3, j, 1),
                                                           {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
                self.worksheet_rework.write_row(0, self.canculate_head_num(3, j), ['HW Rework', 'Changed from previous configuration?'])
                self.worksheet_rework.merge_range(1, self.canculate_head_num(3, j), 11, self.canculate_head_num(3, j), "")
                # Set up some formats to use.
                red = self.workbook.add_format({'color': 'red'})
                self.worksheet_rework.write_rich_string(1,self.canculate_head_num(3, j), red, date_string, self.titleformat)
                if not object_string_list:
                    continue
                self.worksheet_rework.write_column(1, self.canculate_head_num(3, j, 2), object_string_list)
            except:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 HW Rework 数据........\n'

    def insert_HW_data(self, k):
        insert_row_formula_list = []
        fail_url_list = []
        print '开始获取 HW Configuration 数据........\n'
        # 获取公式并插入指定位置
        self.get_formula_data('HW', self.worksheet_hw)
        for j in range(len(self.FPGA_Silver_url_list)):
            print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
            # 获取提取的待插入excel表的数据
            obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
            if j == 0:
                date_string, row_coordinates_list, column_coordinates_list, cell_data_list = obj.get_hw_data('HW Configuration', self.verificate_flag)
            else:
                date_string, row_coordinates_list, column_coordinates_list, cell_data_list = obj.get_hw_data('HW Configuration', True)
            # self.worksheet_hw.write(2, self.canculate_head_num(13, j), date_string, self.titleformat)
            # Set up some formats to use.
            # 合并单元格
            self.worksheet_hw.merge_range('{0}3:{0}9'.format(self.translate_test(ord('A') + 13 * j)), "", self.titleformat_merage)
            red = self.workbook.add_format({'color': 'red'})
            self.worksheet_hw.write_rich_string(2, self.canculate_head_num(13, j), red, date_string, self.titleformat)
            # 标记True为红色
            self.worksheet_hw.conditional_format(0, self.canculate_head_num(13, j, 0), 100, self.canculate_head_num(13, j, 10000), {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
            #判断这周有无数据， 例如：第40周的数据为空
            if not row_coordinates_list and not column_coordinates_list and not cell_data_list:
                print '本周数据为空!!!'
                fail_url_list.append(self.FPGA_Silver_url_list[j])
                continue
            try:
                self.worksheet_hw.write('{0}1'.format( self.translate_test(ord('A') + 13*j) ), 'HW')
                # 写入工作表的第一行,现在只支持system 1 到HWsystem 10，截取10行数据
                self.worksheet_hw.write_row('{0}1' .format( self.translate_test(ord('A') + 3 + 13 * j) ), row_coordinates_list[:10])
                #写入工作表的第二行
                common_formula_lable_list = self.effective_alp_list[j]
                for i in range(len(common_formula_lable_list)):
                    row_formula = '=OR(IF(OR(ISBLANK({0}3),ISBLANK({1}3)),FALSE,NOT(EXACT({0}3,{1}3))),IF(OR(ISBLANK({0}4),ISBLANK({1}4)),FALSE,NOT(EXACT({0}4,{1}4))),IF(OR(ISBLANK({0}5),ISBLANK({1}5)),FALSE,NOT(EXACT({0}5,{1}5))),IF(OR(ISBLANK({0}6),ISBLANK({1}6)),FALSE,NOT(EXACT({0}6,{1}6))),IF(OR(ISBLANK({0}7),ISBLANK({1}7)),FALSE,NOT(EXACT({0}7,{1}7))),IF(OR(ISBLANK({0}8),ISBLANK({1}8)),FALSE,NOT(EXACT({0}8,{1}8))),IF(OR(ISBLANK({0}9),ISBLANK({1}9)),FALSE,NOT(EXACT({0}9,{1}9))))' .format(common_formula_lable_list[i][0], common_formula_lable_list[i][1])
                    insert_row_formula_list.append(row_formula)
                self.worksheet_hw.write_row('{0}2' .format( self.translate_test(ord('A') + 2 + 13 * j) ), ['Changed from previous configuration?'] + insert_row_formula_list)
                #按列插入数据
                insert_col_formula_list = []
                self.worksheet_hw.write_column('{0}3' .format( self.translate_test(ord('A') + 1 + 13 * j) ), column_coordinates_list)
                for col in range(3, 10):
                    part_args = self.effective_alp_list[j][0][0], self.effective_alp_list[j][0][1], \
                                self.effective_alp_list[j][1][0], self.effective_alp_list[j][1][1], \
                                self.effective_alp_list[j][2][0], self.effective_alp_list[j][2][1], \
                                self.effective_alp_list[j][3][0], self.effective_alp_list[j][3][1], \
                                self.effective_alp_list[j][4][0], self.effective_alp_list[j][4][1], \
                                self.effective_alp_list[j][4][0], self.effective_alp_list[j][4][1], \
                                self.effective_alp_list[j][5][0], self.effective_alp_list[j][5][1], \
                                self.effective_alp_list[j][6][0], self.effective_alp_list[j][6][1], \
                                self.effective_alp_list[j][7][0], self.effective_alp_list[j][7][1], \
                                self.effective_alp_list[j][8][0], self.effective_alp_list[j][8][1], \
                                self.effective_alp_list[j][9][0], self.effective_alp_list[j][9][1], col
                    column_formula = '=OR(IF(AND(ISBLANK({0}{22}),ISBLANK({1}{22})),FALSE,NOT(EXACT({0}{22},{1}{22}))),IF(AND(ISBLANK({2}{22}),ISBLANK({3}{22})),FALSE,NOT(EXACT({2}{22},{3}{22}))),IF(AND(ISBLANK({4}{22}),ISBLANK({5}{22})),FALSE,NOT(EXACT({4}{22},{5}{22}))),IF(AND(ISBLANK({6}{22}),ISBLANK({7}{22})),FALSE,NOT(EXACT({6}{22},{7}{22}))),IF(AND(ISBLANK({8}{22}),ISBLANK({9}{22})),FALSE,NOT(EXACT({8}{22},{9}{22}))),IF(AND(ISBLANK({10}{22}),ISBLANK({11}{22})),FALSE,NOT(EXACT({10}{22},{11}{22}))),IF(AND(ISBLANK({12}{22}),ISBLANK({13}{22})),FALSE,NOT(EXACT({12}{22},{13}{22}))),IF(AND(ISBLANK({14}{22}),ISBLANK({15}{22})),FALSE,NOT(EXACT({14}{22},{15}{22}))),IF(AND(ISBLANK({16}{22}),ISBLANK({17}{22})),FALSE,NOT(EXACT({16}{22},{17}{22}))),IF(AND(ISBLANK({18}{22}),ISBLANK({19}{22})),FALSE,NOT(EXACT({18}{22},{19}{22}))),IF(AND(ISBLANK({20}{22}),ISBLANK({21}{22})),FALSE,NOT(EXACT({20}{22},{21}{22}))))'.format( *part_args )
                    insert_col_formula_list.append(column_formula)
                self.worksheet_hw.write_column('{0}3' .format( self.translate_test(ord('A') + 2 + 13 * j) ), insert_col_formula_list)
                #按行插入数据
                for row_num in range(3,len(cell_data_list)+1):
                    #每个元素限制为10列 [:10]
                    self.worksheet_hw.write_row('{1}{0}'.format( row_num, self.translate_test(ord('A') + 3 + 13 * j) ), cell_data_list[row_num-1][:10])

            except:
                clear_row_data_list = []
                for k in range(13):
                    clear_row_data_list.append(' ')
                for row in range(1,10):
                    self.worksheet_hw.write_row('{0}{2}:{1}{2}' .format( self.translate_test(ord('A') + 13*j), self.translate_test(ord('A') + 12 + 13*j), row), clear_row_data_list)
                self.worksheet_hw.write(9, 1 , '=' + 'OR(IF(OR(ISBLANK(E3),ISBLANK(R3)),FALSE,NOT(EXACT(E3,R3))),IF(OR(ISBLANK(E4),ISBLANK(R4)),FALSE,NOT(EXACT(E4,R4))),IF(OR(ISBLANK(E5),ISBLANK(R5)),FALSE,NOT(EXACT(E5,R5))),IF(OR(ISBLANK(E6),ISBLANK(R6)),FALSE,NOT(EXACT(E6,R6))),IF(OR(ISBLANK(E7),ISBLANK(R7)),FALSE,NOT(EXACT(E7,R7))),IF(OR(ISBLANK(E8),ISBLANK(R8)),FALSE,NOT(EXACT(E8,R8))),IF(OR(ISBLANK(E9),ISBLANK(R9)),FALSE,NOT(EXACT(E9,R9))))')
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 HW Configuration 数据........\n'

    def insert_SW_Original_data(self, k):
        fail_url_list = []
        print '开始获取 SW_Original Configuration 数据........\n'
        # 获取公式并插入指定位置
        self.get_formula_data('SW_Original', self.worksheet_sw_orignal)
        for j in range(len(self.FPGA_Silver_url_list)):
            try:
                print '开始插入第[ %d ]个[ %s ]对应的数据' %(j+1, self.FPGA_Silver_url_list[j])
                obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
                if j == 0:
                    date_string, url_list, header_list, cell_data_list, left_col_list_1, left_col_list_2 = obj.get_sw_data('SW Configuration', self.verificate_flag)
                else:
                    date_string, url_list, header_list, cell_data_list, left_col_list_1, left_col_list_2 = obj.get_sw_data('SW Configuration', True)

                red = self.workbook.add_format({'color': 'red'})
                blue = self.workbook.add_format({'color': 'blue'})
                self.worksheet_sw_orignal.write_rich_string(0, self.canculate_head_num(9, j), red, date_string,
                                                          blue, ' Ingredient list Changed comparing to Next Release?',  self.header_format)
                self.worksheet_sw_orignal.write_rich_string(0, self.canculate_head_num(9, j, 1), blue,
                                                            'Ingredient list Changed comparing to Previous?', self.header_format)
                if not url_list and not header_list and not cell_data_list and not left_col_list_1 and not left_col_list_2:
                    continue
                # Set up some formats to use.
                # 标记True为红色
                self.worksheet_sw_orignal.conditional_format(1, self.canculate_head_num(9, j), 40, self.canculate_head_num(9, j, 1),
                                                         {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})

                self.worksheet_sw_orignal.write_row(0, self.canculate_head_num(9, j, 5), header_list, self.header_format)
                #插入数据
                #最后一列可能会出现多值多行的情况，计算每行数据占有的行数
                line_num_list = []
                for ele in cell_data_list:
                    num = len(ele[3:])
                    line_num_list.append(num)
                nu = 1
                first_insert_data = [cell_data_list[k][0] for k in range(len(cell_data_list))]
                second_insert_data = [cell_data_list[k][2] for k in range(len(cell_data_list))]
                #插入数据第一列到第三列
                for line in range(len(cell_data_list)):
                    if line_num_list[line] == 1:
                        self.worksheet_sw_orignal.write(nu, self.canculate_head_num(9, j, 5), first_insert_data[line])
                        self.worksheet_sw_orignal.write_url(nu, self.canculate_head_num(9, j, 6), url_list[line][0], self.url_format, cell_data_list[line][1])
                        self.worksheet_sw_orignal.write(nu, self.canculate_head_num(9, j, 7), second_insert_data[line])

                        if (len(url_list[line]) >= 2) and cell_data_list[line][3]:
                            self.worksheet_sw_orignal.write_url(nu, self.canculate_head_num(9, j, 8), url_list[line][1], self.url_format, cell_data_list[line][3])
                        else:
                            self.worksheet_sw_orignal.write(nu, self.canculate_head_num(9, j, 8), cell_data_list[line][3])
                        nu += 1

                    elif line_num_list[line] > 1:
                        length_merge = len(url_list[line][1:])
                        self.worksheet_sw_orignal.merge_range(nu, self.canculate_head_num(9, j, 5), nu + line_num_list[line] -1, self.canculate_head_num(9, j, 5), first_insert_data[line], self.titleformat_merage)
                        self.worksheet_sw_orignal.merge_range(nu, self.canculate_head_num(9, j, 6), nu + line_num_list[line] - 1, self.canculate_head_num(9, j, 6), '')
                        self.worksheet_sw_orignal.write_url(nu, self.canculate_head_num(9, j, 6), url_list[line][0], self.url_format, cell_data_list[line][1])
                        self.worksheet_sw_orignal.merge_range(nu, self.canculate_head_num(9, j, 7), nu + line_num_list[line] -1, self.canculate_head_num(9, j, 7), second_insert_data[line], self.titleformat_merage)

                        for i in range(nu, length_merge + nu):
                            self.worksheet_sw_orignal.write_url(i, self.canculate_head_num(9, j, 8), url_list[line][1 + i - nu], self.url_format, cell_data_list[line][3:][i - nu])
                        self.worksheet_sw_orignal.write(line_num_list[line] + nu - 1, self.canculate_head_num(9, j, 8), cell_data_list[line][3:][-1])
                        nu += line_num_list[line]
            except:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 SW_Original Configuration 数据........\n'

    def insert_IFWI_Orignal_data(self, k):
        print '开始获取 IFWI_Orignal Configuration 数据........\n'
        fail_url_list = []
        # 获取公式并插入指定位置
        self.get_formula_data('IFWI_Original', self.worksheet_ifwi_orignal)
        for j in range(len(self.FPGA_Silver_url_list)):
            try:
                print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
                obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
                if j == 0:
                    date_string, header_list, cell_data_list = obj.get_ifwi_data('IFWI Configuration', self.verificate_flag)
                else:
                    date_string, header_list, cell_data_list = obj.get_ifwi_data('IFWI Configuration', True)

                insert_rich_string_list = ['Ingredient list Changed comparing to Next Release?',
                                           'Ingredient list Changed comparing to Previous?',
                                           'Ingredient is not in Mapping table?']
                red = self.workbook.add_format({'color': 'red'})
                blue = self.workbook.add_format({'color': 'blue'})
                for i in range(3):
                    self.worksheet_ifwi_orignal.write_rich_string(2, self.canculate_head_num(6, j, i), blue, insert_rich_string_list[i], self.header_format)
                # Set up some formats to use.
                self.worksheet_ifwi_orignal.conditional_format(3, self.canculate_head_num(6, j), 50, self.canculate_head_num(6, j, 1),
                                                             {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
                self.worksheet_ifwi_orignal.write_rich_string(0, self.canculate_head_num(6, j, 1), blue, 'Ingredient list Changed?', self.header_format)
                self.worksheet_ifwi_orignal.write_rich_string(0, self.canculate_head_num(6, j), red, date_string, self.header_format)
                self.worksheet_ifwi_orignal.write_rich_string(1, self.canculate_head_num(6, j, 3), blue, 'IFWI Configuration ', self.header_format)

                if not header_list and not cell_data_list:
                    continue

                self.worksheet_ifwi_orignal.write_row(2, self.canculate_head_num(6, j, 3), header_list, self.header_format)
                #插入数据,需要考虑有数字的情况，前面加Nic字母
                for ele in range(len(cell_data_list)):
                    #以数字开头的元素前面加Nic
                    match_obj = re.match('\s+\d+', str(cell_data_list[ele][0]))
                    if match_obj:
                        cell_data_list[ele][0] = 'Nic' + cell_data_list[ele][0]
                    cell_data_list[ele][0] = cell_data_list[ele][0].lstrip(' ')

                for begain in range(len(cell_data_list)):
                    self.worksheet_ifwi_orignal.write_row(begain + 3,  self.canculate_head_num(6, j, 3), cell_data_list[begain])
            except ValueError:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 IFWI_Orignal Configuration 数据........\n'

    def insert_SW_data(self, k):
        print '开始获取 SW Configuration 数据........\n'
        fail_url_list = []
        #获取公式并插入指定位置
        self.get_formula_data('SW', self.worksheet_sw)
        for j in range(len(self.FPGA_Silver_url_list)):
            try:
                print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
                obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
                if j == 0:
                    date_string, url_list, header_list, cell_data_list, left_col_list_1, left_col_list_2 = obj.get_sw_data('SW Configuration', self.verificate_flag)
                else:
                    date_string, url_list, header_list, cell_data_list, left_col_list_1, left_col_list_2 = obj.get_sw_data('SW Configuration', True)

                red = self.workbook.add_format({'color': 'red'})
                blue = self.workbook.add_format({'color': 'blue'})
                self.worksheet_sw.write_rich_string(14, self.canculate_head_num(9, j), red, date_string,
                                                    blue, ' Ingredient list Changed comparing to Next Release?', self.header_format)
                self.worksheet_sw.write_rich_string(14, self.canculate_head_num(9, j, 1), blue,
                                                    'Ingredient list Changed comparing to Previous?', self.header_format)
                if not url_list and not header_list and not cell_data_list and not left_col_list_1 and not left_col_list_2:
                    continue
                # Set up some formats to use.
                self.worksheet_sw.conditional_format(1, self.canculate_head_num(9, j), 40, self.canculate_head_num(9, j, 1),
                                                             {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
                self.worksheet_sw.write_row(14, self.canculate_head_num(9, j, 5), header_list, self.header_format)
                #处理并插入数据
                insert_data = zip( cell_data_list, url_list )
                insert_data.sort(key=lambda x: x[0][0].upper())
                # insert_data.sort(cmp=lambda x,y : cmp(x[0][0].upper(), y[0][0].upper()))
                # 插入数据
                first_insert_data = [insert_data[k][0][0] for k in range(len(insert_data))]
                second_insert_data = [insert_data[k][0][2] for k in range(len(insert_data))]
                #注意中间插入了一行标头分隔了数据
                self.worksheet_sw.write_column(0, self.canculate_head_num(9, j, 5), first_insert_data[:14])
                self.worksheet_sw.write_column(0, self.canculate_head_num(9, j, 7), second_insert_data[:14])
                self.worksheet_sw.write_column(15, self.canculate_head_num(9, j, 5), first_insert_data[14:])
                self.worksheet_sw.write_column(15, self.canculate_head_num(9, j, 7), second_insert_data[14:])
                # 插入url部分
                for nu in range(len(insert_data)):
                    if nu == 14:
                        continue
                    self.worksheet_sw.write_url(nu, self.canculate_head_num(9, j, 6), insert_data[nu][1][0], self.url_format, insert_data[nu][0][1])
                    if insert_data[nu][0][3] == 'Release Notes':
                        self.worksheet_sw.write_url(nu  , self.canculate_head_num(9, j, 8), insert_data[nu][1][1], self.url_format, insert_data[nu][0][3])
                    else:
                        self.worksheet_sw.write(nu, self.canculate_head_num(9, j, 8), insert_data[nu][0][3])
            except:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 SW Configuration 数据........\n'

    def insert_IFWI_data(self, k):
        print '开始获取 IFWI Configuration 数据........\n'
        fail_url_list = []
        # 获取公式并插入指定位置
        self.get_formula_data('IFWI', self.worksheet_ifwi)
        for j in range(len(self.FPGA_Silver_url_list)):
            try:
                print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
                obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
                if j == 0:
                    date_string, header_list, cell_data_list = obj.get_ifwi_data('IFWI Configuration',self.verificate_flag)
                else:
                    date_string, header_list, cell_data_list = obj.get_ifwi_data('IFWI Configuration', True)

                red = self.workbook.add_format({'color': 'red'})
                blue = self.workbook.add_format({'color': 'blue'})
                self.worksheet_ifwi.write_rich_string(0, self.canculate_head_num(6, j, 1), blue, 'Ingredient list Changed?', self.header_format)
                self.worksheet_ifwi.write_rich_string(0, self.canculate_head_num(6, j), red, date_string, self.header_format)
                self.worksheet_ifwi.write_rich_string(1, self.canculate_head_num(6, j, 3), blue, 'IFWI Configuration ', self.header_format)
                insert_rich_string_list = ['Ingredient list Changed comparing to Next Release?',
                                           'Ingredient list Changed comparing to Previous?',
                                           'Ingredient is not in Mapping table?']
                for i in range(3):
                    self.worksheet_ifwi.write_rich_string(2, self.canculate_head_num(6, j, i), blue, insert_rich_string_list[i], self.header_format)
                if not header_list and not cell_data_list:
                    continue
                # Set up some formats to use.
                self.worksheet_ifwi.conditional_format(3, self.canculate_head_num(6, j), 50, self.canculate_head_num(6, j, 1),
                                                             {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
                self.worksheet_ifwi.write_row(2, self.canculate_head_num(6, j, 3), header_list, self.header_format)
                #插入数据,需要考虑有数字的情况，前面加Nic字母
                for ele in range(len(cell_data_list)):
                    #以数字开头的元素前面加Nic
                    match_obj = re.match('^\d+', str(cell_data_list[ele][0]))
                    if match_obj:
                        cell_data_list[ele][0] = 'Nic' + cell_data_list[ele][0]
                    cell_data_list[ele][0] = cell_data_list[ele][0].lstrip(' ')

                cell_data_list.sort(key=lambda x: x[0].upper())
                for begain in range(len(cell_data_list)):
                    self.worksheet_ifwi.write_row(begain + 3,  self.canculate_head_num(6, j, 3), cell_data_list[begain])
            except ValueError:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 IFWI Configuration 数据........\n'

    def insert_Mapping(self, k):
        width = len(self.FPGA_Silver_url_list)
        # 获取公式并插入指定位置
        self.get_formula_data('Mapping', self.worksheet_mapping)
        first_value_list = ['Considering SW Ingredient adding as change?'] + ['TRUE']*len(self.FPGA_Silver_url_list)
        second_value_list = first_value_list
        third_value_list = ['Considering SW Ingredient adding as change?'] + ['FALSE']*len(self.FPGA_Silver_url_list)
        fourth_value_list = third_value_list
        # 标记True为红色
        self.worksheet_mapping.conditional_format(5+width, 5, 153, 30, {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
        self.worksheet_mapping.conditional_format(5+width, 31, 153, 64, {'type': 'cell', 'criteria': '=', 'value': 1, 'format': self.format1})
        self.worksheet_mapping.write_row(0, 4, first_value_list, self.mapping_format)
        self.worksheet_mapping.write_row(1, 4, second_value_list, self.mapping_format)
        self.worksheet_mapping.write_row(2, 4, third_value_list, self.mapping_format)
        self.worksheet_mapping.write_row(3, 4, fourth_value_list, self.mapping_format)
        # for i in range(4, width+8):
        #     self.worksheet_mapping.write_row(i, 0, [' ']*30, self.mapping_blue_format)
        insert_date_list = []
        for url in self.FPGA_Silver_url_list:
            split_string_list = str(url).split('/')
            object_string = split_string_list[-2][-4:]
            insert_date_list.append(object_string + ' Mapped?')
        # insert_date_list_column = insert_date_list + ['Total case', 'basic case', 'optionnal case', 'Selected case', 'Covered Issue',
        #                          'Covered Issue Delta', 'Missed Issue', 'Missed Issue Delta']
        # self.worksheet_mapping.write_column(5, 30, insert_date_list_column, self.header_format)
        self.worksheet_mapping.write_row(5+width-2, 5, insert_date_list + ['EXX_WW49', 'changed?\Test?'], self.header_format)
        self.worksheet_mapping.write_column(5, 5+width+1, insert_date_list, self.header_format)
        red = self.workbook.add_format({'color': 'red'})
        self.worksheet_mapping.write_rich_string(5+width-2, 5+width-1, red, 'EXX_WW49', self.header_format)

        row_header_list = []
        row_value_list = ['Power Management', 'Networking', 'USB', 'FPGA', 'Video', 'Storage', 'PCIe', 'Manageability',
                           'Virtualization', 'Memory', 'Stress & Stability', 'Security', 'Processor', 'DPDK', 'DPDK_FVL_10G',
                           'QAT', 'RAS', 'Summary']
        for head in row_value_list:
            row_header_list.append(head)
            row_header_list.append(' ')
        self.worksheet_mapping.write_row(4, 5+width+3, row_header_list, self.header_format)

    def insert_CaseResult(self, k):
        print '开始获取 CaseResult 数据........\n'
        fail_url_list = []
        # 获取公式并插入指定位置
        self.get_formula_data('CaseResult', self.worksheet_caseresult)
        for j in range(1):
            try:
                print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
                obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
                if j == 0:
                    date_string, tip_string, header_list, cell_data_list = obj.get_caseresult_data('Platform Integration Validation Result', self.verificate_flag)
                else:
                    date_string, tip_string, header_list, cell_data_list = obj.get_caseresult_data('Platform Integration Validation Result', True)

                if not tip_string and not header_list and not cell_data_list:
                    continue
                # 标记True为红色
                self.worksheet_caseresult.conditional_format(0, 0, 5000, 1000, {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})
            except:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 CaseResult 数据........\n'

    def insert_Platform_data(self, k):
        print '开始获取 Platform Integration Validation Result 数据........\n'
        fail_url_list = []
        for j in range(len(self.FPGA_Silver_url_list)):
            try:
                print '开始插入第[ %d ]个[ %s ]对应的数据' % (j + 1, self.FPGA_Silver_url_list[j])
                obj = extract_data.GetAnalysisData(self.FPGA_Silver_url_list[j], get_info_not_save_flag=False, cache=self.cache)
                if j == 0:
                    date_string, url_list, header_list, cell_data_list = obj.get_platform_data('Platform Integration Validation Result', self.verificate_flag)
                else:
                    date_string, url_list, header_list, cell_data_list = obj.get_platform_data('Platform Integration Validation Result', True)
                # Set up some formats to use.
                red = self.workbook.add_format({'color': 'red'})
                self.worksheet_platform.write_rich_string(0, self.canculate_head_num(12, j), red, date_string, self.titleformat)
                if not url_list and not header_list and not cell_data_list:
                    continue
                # Set up some formats to use.
                # 标记True为红色
                self.worksheet_platform.conditional_format(2, self.canculate_head_num(12, j, 5), 20, self.canculate_head_num(12, j, 5),
                                                         {'type': 'cell', 'criteria': '>', 'value': 0, 'format': self.format1})
                # 最后一列可能会出现多值多行的情况，计算每行数据占有的行数
                line_num_list = []
                for ele in cell_data_list:
                    num = len(ele[10:])
                    line_num_list.append(num)

                nu = 2
                # 插入数据第一列到第三列
                for line in range(len(cell_data_list)):
                    #将数字字符转化为数字
                    temp_list = copy.deepcopy(cell_data_list[line])
                    for ele in range(len(temp_list)):
                        if temp_list[ele].isdigit():
                            temp_list[ele] = int(temp_list[ele])

                    if line_num_list[line] == 1:
                        self.worksheet_platform.write_row(nu, self.canculate_head_num(12, j, 1), temp_list)
                        if url_list[line]:
                            self.worksheet_platform.write_url(nu, self.canculate_head_num(12, j, 11), url_list[line][0], self.url_format, str(temp_list[10]))
                        nu += 1

                    elif line_num_list[line] > 1:
                        length_merge = line_num_list[line]
                        for i in range(1, 11):
                            self.worksheet_platform.merge_range(nu, self.canculate_head_num(12, j, i),nu + line_num_list[line] - 1,self.canculate_head_num(12, j, i), '')
                        self.worksheet_platform.write_row(nu, self.canculate_head_num(12, j, 1), temp_list[:10], self.titleformat_merage)
                        for m in range(length_merge):
                            self.worksheet_platform.write_url(nu+m, self.canculate_head_num(12, j, 11), url_list[line][m], self.url_format, str(temp_list[10+m]))
                        nu += line_num_list[line]

                self.worksheet_platform.write_row(1, self.canculate_head_num(12, j, 1), header_list, self.header_format)

            except:
                print '爬取第[ %d ]个[ %s ]对应的数据失败' % (j + 1, self.FPGA_Silver_url_list[j])
                fail_url_list.append(self.FPGA_Silver_url_list[j])
        # print '失败的url列表:\t', fail_url_list
        # print '结束获取 Platform Integration Validation Result 数据........\n'

    def insert_save_miss_data(self, k):
        # 获取公式并插入指定位置
        self.get_formula_data('Save-Miss', self.worksheet_save_miss)
        # 标记True为红色
        self.worksheet_save_miss.conditional_format(0, 0, 5000, 1000, {'type': 'cell', 'criteria': '=', 'value': True, 'format': self.format1})

    def insert_test_time_data(self, k):
        # 获取公式并插入指定位置
        self.get_formula_data('TestTime', self.worksheet_test_time)
        # 标记True为红色
        self.worksheet_test_time.conditional_format(0, 0, 5000, 1000, {'type': 'cell', 'criteria': '=', 'value': True,
                                                                       'format': self.format1})

    def return_name(self):
        return self.__class__.__dict__

    def close_workbook(self):
        self.workbook.close()


if __name__ == '__main__':
    start = time.time()
    cache = DiskCache()
    obj = InsertDataIntoExcel(verificate_flag=False, cache=cache)
    # getattr(obj, 'insert_ExistingSi_data')()
    func_name_list = obj.return_name().keys()
    use_func_list = [ func for func in func_name_list if func.startswith('insert') ]
    print use_func_list
    # for func in use_func_list:
    #     getattr(obj, func)()
    # thread_list = []
    # for func in use_func_list:
    #     thread_list.append(threading.Thread(target=getattr(obj, func), args=()))
    # for thread in thread_list:
    #     thread.start()
    # for thread in thread_list:
    #     thread.join()
    # import Queue, random
    # def read(q, obj):
    #     while True:
    #         try:
    #             value = q.get()
    #             print('Get %s from queue.' % value)
    #             getattr(obj, value)()
    #             time.sleep(random.random())
    #         finally:
    #             q.task_done()
    # def main():
    #     q = Queue.Queue()
    #     pw1 = threading.Thread(target=read, args=(q, obj))
    #     # pw2 = threading.Thread(target=read, args=(q, obj))
    #     # pw3 = threading.Thread(target=read, args=(q, obj))
    #     # pw4 = threading.Thread(target=read, args=(q, obj))
    #     # pw1.daemon = True
    #     # pw2.daemon = True
    #     # pw3.daemon = True
    #     # pw4.daemon = True
    #     pw1.start()
    #     # pw2.start()
    #     # pw3.start()
    #     # pw4.start()
    #     for func in use_func_list:
    #         q.put(func)
    #     try:
    #         q.join()
    #     except KeyboardInterrupt:
    #         print("stopped by hand")
    #
    # main()
    # obj.insert_Newsi_data(1)
    obj.insert_ExistingSi_data(1)
    # obj.insert_ClosedSi_data(1)
    # obj.insert_Rework_data()
    # obj.insert_HW_data(1)
    # obj.insert_SW_Original_data()
    # obj.insert_IFWI_Orignal_data()
    # obj.insert_SW_data(1)
    # obj.insert_IFWI_data(1)
    # obj.insert_Platform_data(1)
    # obj.insert_Mapping(1)
    # obj.insert_CaseResult()
    # obj.insert_save_miss_data()
    # obj.insert_test_time_data()
    obj.close_workbook()
    print time.time() - start



